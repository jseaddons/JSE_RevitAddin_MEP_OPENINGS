using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class LinkedFileInfo
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public LinkedFileType FileType { get; set; }
        public bool IsLoaded { get; set; }
        public int ElementCount { get; set; }
        public RevitLinkInstance? LinkInstance { get; set; }
    }

    public class LinkedFileService
    {
        private readonly LinkedFileDetectionService _detectionService;

        public LinkedFileService()
        {
            _detectionService = new LinkedFileDetectionService();
        }

        public List<LinkedFileInfo> GetLinkedFiles(Document document)
        {
            var linkedFiles = new List<LinkedFileInfo>();

            try
            {
                // Get all RevitLinkInstance elements
                var linkInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(link => link.GetLinkDocument() != null)
                    .ToList();

                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null) continue;

                    var fileName = linkDoc.Title;
                    var filePath = linkDoc.PathName;
                    var fileType = LinkedFileDetectionService.DetectFileType(fileName);
                    var isLoaded = true; // Assume loaded if we can access the document
                    
                    // Count elements in the linked document
                    var elementCount = GetElementCountForFileType(linkDoc, fileType);

                    linkedFiles.Add(new LinkedFileInfo
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        FileType = fileType,
                        IsLoaded = isLoaded,
                        ElementCount = elementCount,
                        LinkInstance = linkInstance
                    });
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting linked files: {ex.Message}");
            }

            return linkedFiles;
        }

        private int GetElementCountForFileType(Document linkDoc, LinkedFileType fileType)
        {
            try
            {
                return fileType switch
                {
                    LinkedFileType.Electrical => GetElectricalElementCount(linkDoc),
                    LinkedFileType.Mechanical => GetMechanicalElementCount(linkDoc),
                    LinkedFileType.Plumbing => GetPlumbingElementCount(linkDoc),
                    LinkedFileType.FireProtection => GetPlumbingElementCount(linkDoc), // Same as plumbing
                    LinkedFileType.Architectural => GetArchitecturalElementCount(linkDoc),
                    LinkedFileType.Structural => GetStructuralElementCount(linkDoc),
                    _ => GetTotalElementCount(linkDoc)
                };
            }
            catch
            {
                return 0;
            }
        }

        private int GetElectricalElementCount(Document doc)
        {
            var categories = new[]
            {
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit
            };

            return GetElementCountByCategories(doc, categories);
        }

        private int GetMechanicalElementCount(Document doc)
        {
            var categories = new[]
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory
            };

            return GetElementCountByCategories(doc, categories);
        }

        private int GetPlumbingElementCount(Document doc)
        {
            var categories = new[]
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting
            };

            return GetElementCountByCategories(doc, categories);
        }

        private int GetArchitecturalElementCount(Document doc)
        {
            var categories = new[]
            {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows
            };

            return GetElementCountByCategories(doc, categories);
        }

        private int GetStructuralElementCount(Document doc)
        {
            var categories = new[]
            {
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };

            return GetElementCountByCategories(doc, categories);
        }

        private int GetTotalElementCount(Document doc)
        {
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .GetElementCount();
        }

        private int GetElementCountByCategories(Document doc, BuiltInCategory[] categories)
        {
            var elementIds = new List<ElementId>();
            
            foreach (var category in categories)
            {
                var categoryId = new ElementId(category);
                var elements = new FilteredElementCollector(doc)
                    .OfCategoryId(categoryId)
                    .WhereElementIsNotElementType()
                    .ToElementIds();
                
                elementIds.AddRange(elements);
            }

            return elementIds.Distinct().Count();
        }

        public List<LinkedFileInfo> GetHostElementFiles(List<LinkedFileInfo> allFiles)
        {
            return allFiles.Where(f => f.FileType == LinkedFileType.Architectural || 
                                      f.FileType == LinkedFileType.Structural).ToList();
        }

        public List<LinkedFileInfo> GetReferenceElementFiles(List<LinkedFileInfo> allFiles)
        {
            return allFiles.Where(f => f.FileType == LinkedFileType.Electrical || 
                                      f.FileType == LinkedFileType.Mechanical || 
                                      f.FileType == LinkedFileType.Plumbing ||
                                      f.FileType == LinkedFileType.FireProtection ||
                                      f.FileType == LinkedFileType.Unknown).ToList(); // Include Unknown files
        }
    }
}
