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

                    // ✅ FIX: Extract filename from PathName, fallback to Name or Title
                    // Some link instances have Name as "location Shared" which is not useful
                    // Prefer filename from PathName, then Name (if not "location Shared"), then Title
                    string fileName = string.Empty;
                    if (!string.IsNullOrWhiteSpace(linkDoc.PathName))
                    {
                        fileName = System.IO.Path.GetFileNameWithoutExtension(linkDoc.PathName);
                    }
                    if (string.IsNullOrWhiteSpace(fileName) || fileName.Equals("location Shared", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.IsNullOrWhiteSpace(linkInstance.Name) && 
                            !linkInstance.Name.Equals("location Shared", StringComparison.OrdinalIgnoreCase))
                        {
                            fileName = linkInstance.Name;
                        }
                        else if (!string.IsNullOrWhiteSpace(linkDoc.Title))
                        {
                            fileName = linkDoc.Title;
                        }
                    }
                    // Final fallback - use PathName filename even if it says "location Shared"
                    if (string.IsNullOrWhiteSpace(fileName) && !string.IsNullOrWhiteSpace(linkDoc.PathName))
                    {
                        fileName = System.IO.Path.GetFileNameWithoutExtension(linkDoc.PathName);
                    }
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

        public void AdoptProvisionForVoids(Document doc)
        {
            var settings = ApplicationProfileService.Instance.GetCurrentSettings();

            if (!settings.AdoptProvisionForVoids)
                return;

            var linkedFiles = GetLinkedFiles(doc);

            foreach (var linkedFile in linkedFiles)
            {
                if (linkedFile.LinkInstance == null)
                    continue;

                var linkedDoc = linkedFile.LinkInstance.GetLinkDocument();
                if (linkedDoc == null)
                    continue;

                // Find provisions for voids in the linked document
                var provisions = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .Where(e => e.Name.Contains("Provision for Void"))
                    .Cast<FamilyInstance>()
                    .ToList();

                foreach (var provision in provisions)
                {
                    // Get the geometry of the provision
                    var geometry = provision.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                    if (geometry == null)
                        continue;

                    // Get the transform of the linked document
                    var transform = linkedFile.LinkInstance.GetTotalTransform();

                    // Transform the geometry to the host document's coordinates
                    var transformedGeometry = geometry.GetTransformed(transform);

                    // Find the host element in the host document
                    var hostElement = FindHostElement(doc, transformedGeometry);
                    if (hostElement == null)
                        continue;

                    // Create an opening in the host document
                    var opening = CreateOpeningFromProvision(doc, provision, transformedGeometry, hostElement);
                    if (opening == null)
                        continue;

                    // Link the opening to the provision
                    LinkOpeningToProvision(opening, provision);
                }
            }
        }

        private Element FindHostElement(Document doc, GeometryElement geometry)
        {
            // TODO: Implement logic to find the host element for the provision for void
            return null;
        }

        private FamilyInstance CreateOpeningFromProvision(Document doc, FamilyInstance provision, GeometryElement geometry, Element hostElement)
        {
            // TODO: Implement logic to create an opening from the provision for void
            return null;
        }

        private void LinkOpeningToProvision(FamilyInstance opening, FamilyInstance provision)
        {
            // TODO: Implement logic to link the opening to the provision for void
        }
    }
}
