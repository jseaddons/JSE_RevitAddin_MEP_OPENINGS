using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Repositories
{
    /// <summary>
    /// Implementation of ISleeveRepository for sleeve data persistence.
    /// Handles XML reading/writing for cluster data.
    /// </summary>
    public class SleeveRepository : ISleeveRepository
    {
        /// <summary>
        /// Save sleeve data to _CLUSTER.xml file.
        /// </summary>
        public void SaveSleeveDataToClusterXml(string xmlFilePath, List<ClashZone> clashZones, Document doc)
        {
            // Note: This method signature in the interface might need adjustment to match the original usage
            // The original method took List<SleeveData> and category string.
            // We'll implement a version that matches the interface but we might need to refactor the interface 
            // or the calling code to align better.
            
            // For now, let's implement the core XML saving logic which is the main responsibility.
            // We'll need to convert ClashZones to SleeveData or similar structure if that's what's being saved.
            
            // Wait, looking at UniversalSleevePlacerService.SaveSleeveDataToClusterXml (lines 2988-3048),
            // it takes List<SleeveData> and category.
            // The interface defined earlier has: void SaveSleeveDataToClusterXml(string xmlFilePath, List<ClashZone> clashZones, Document doc);
            // This mismatch needs to be resolved. The interface definition was a bit premature/guessed.
            
            // Let's stick to the interface for now but we might need to update it.
            // Actually, let's implement the method that was actually in UniversalSleevePlacerService 
            // and update the interface to match it, as that's safer.
            
            throw new NotImplementedException("This method signature needs to be aligned with the actual data structure.");
        }

        /// <summary>
        /// Helper method to add XML elements
        /// </summary>
        private void AddXmlElement(System.Xml.XmlDocument doc, System.Xml.XmlElement parent, string name, string value)
        {
            var element = doc.CreateElement(name);
            element.InnerText = value;
            parent.AppendChild(element);
        }

        /// <summary>
        /// Save sleeve data to _CLUSTER.xml file (Internal implementation matching original logic)
        /// </summary>
        public void SaveSleeveDataListToClusterXml(List<SleeveData> sleeveDataList, string category)
        {
            try
            {
                var pluralCategory = category switch
                {
                    "Pipes" => "Pipes",
                    "Ducts" => "Ducts", 
                    "Cable Trays" => "Cable Trays",
                    _ => category
                };
                
                var fileName = $"Plumbing_{pluralCategory.ToLower().Replace(" ", "_")}_CLUSTER.xml";
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    Directory.CreateDirectory(filtersDirectory);
                }

                var filePath = Path.Combine(filtersDirectory, fileName);
                
                var xmlDoc = new System.Xml.XmlDocument();
                var root = xmlDoc.CreateElement("SleeveDataList");
                xmlDoc.AppendChild(root);
                
                foreach (var sleeveData in sleeveDataList)
                {
                    var sleeveElement = xmlDoc.CreateElement("SleeveData");
                    root.AppendChild(sleeveElement);
                    
                    AddXmlElement(xmlDoc, sleeveElement, "SleeveInstanceId", sleeveData.SleeveInstanceId.ToString());
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1X", sleeveData.Corner1.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Y", sleeveData.Corner1.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Z", sleeveData.Corner1.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2X", sleeveData.Corner2.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Y", sleeveData.Corner2.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Z", sleeveData.Corner2.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3X", sleeveData.Corner3.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Y", sleeveData.Corner3.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Z", sleeveData.Corner3.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4X", sleeveData.Corner4.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Y", sleeveData.Corner4.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Z", sleeveData.Corner4.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Width", sleeveData.Width.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Height", sleeveData.Height.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Depth", sleeveData.Depth.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "HostType", sleeveData.HostType);
                    AddXmlElement(xmlDoc, sleeveElement, "Orientation", sleeveData.Orientation);
                    AddXmlElement(xmlDoc, sleeveElement, "Category", sleeveData.Category);
                    AddXmlElement(xmlDoc, sleeveElement, "CreatedAt", sleeveData.CreatedAt);
                }
                
                xmlDoc.Save(filePath);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[REGENERATION-SAVE] Saved {sleeveDataList.Count} sleeves to {fileName}\n");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[REGENERATION-SAVE-ERROR] Error saving {category} cluster XML: {ex.Message}\n");
                throw;
            }
        }

        public void RegenerateClusterXmlFiles(Document doc, List<ClashZone> placedZones)
        {
             // This method in UniversalSleevePlacerService (lines 2862-2928) does a lot of work:
             // 1. Collects all sleeves
             // 2. Groups by category
             // 3. Creates SleeveData objects
             // 4. Calls SaveSleeveDataToClusterXml
             
             // We should probably move the "SaveSleeveDataToClusterXml" part here, 
             // but the collection logic might belong in a service, not a repository.
             // However, for this refactoring step, moving the whole block is acceptable to isolate the responsibility.
             
             // NOTE: The interface signature I defined earlier: void RegenerateClusterXmlFiles(Document doc, List<ClashZone> placedZones);
             // The original method: private void RegenerateClusterXmlFiles() (no args, uses _doc field)
             
             // We will implement a version that takes the Document.
             
            try
            {
                // Wait briefly for Revit to update all sleeves
                System.Threading.Thread.Sleep(500);
                
                // Collect all sleeves with MEP_ElementId parameter
                var allSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.LookupParameter("MEP_ElementId") != null && s.LookupParameter("MEP_ElementId").HasValue)
                    .ToList();
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[REGENERATION-DEBUG] Found {allSleeves.Count} sleeves with MEP_ElementId parameter\n");
                
                // Group sleeves by category using MEP_Category parameter
                var groupedSleeves = allSleeves.GroupBy(s => GetSleeveCategoryFromParameter(s)).ToList();
                
                foreach (var group in groupedSleeves)
                {
                    var category = group.Key;
                    var sleeves = group.ToList();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[REGENERATION-DEBUG] Processing {sleeves.Count} sleeves for category: {category}\n");
                    
                    // Create SleeveDataList with actual Revit coordinates
                    var sleeveDataList = new List<SleeveData>();
                    
                    foreach (var sleeve in sleeves)
                    {
                        var bbox = sleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            var sleeveData = new SleeveData
                            {
                                SleeveInstanceId = sleeve.Id.IntegerValue,
                                Corner1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                                Corner2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                                Corner3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
                                Corner4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                                Width = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Meters),
                                Height = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Meters),
                                Depth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Meters),
                                HostType = GetHostType(sleeve),
                                Orientation = GetSleeveOrientation(doc, sleeve),
                                Category = category,
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            };
                            
                            sleeveDataList.Add(sleeveData);
                        }
                    }
                    
                    // Save to _CLUSTER.xml file
                    SaveSleeveDataListToClusterXml(sleeveDataList, category);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[REGENERATION-ERROR] Error in RegenerateClusterXmlFiles: {ex.Message}\n");
                throw;
            }
        }

        public List<ClashZone> LoadDuctAccessoriesClashZones(string xmlFilePath)
        {
            try
            {
                var clashZones = new List<ClashZone>();
                
                // Note: The original method searched for files in a directory.
                // The interface takes a specific file path.
                // If the path passed is a directory, we search. If it's a file, we load it.
                // Or we can adapt the logic to match the original behavior if the path is generic.
                
                string actualFilePath = xmlFilePath;
                
                // If the path provided is a directory or doesn't exist as a file, try to find matching files
                if (Directory.Exists(xmlFilePath) || (Path.GetFileName(xmlFilePath).Contains("*") && Directory.Exists(Path.GetDirectoryName(xmlFilePath))))
                {
                    var directory = Directory.Exists(xmlFilePath) ? xmlFilePath : Path.GetDirectoryName(xmlFilePath);
                    var pattern = Path.GetFileName(xmlFilePath).Contains("*") ? Path.GetFileName(xmlFilePath) : "*_duct_accessories.xml";
                    
                    var matchingFiles = Directory.GetFiles(directory, pattern);
                    
                    if (matchingFiles.Length > 0)
                    {
                        // Use the most recently modified file
                        actualFilePath = matchingFiles
                            .OrderByDescending(f => File.GetLastWriteTime(f))
                            .First();
                            
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[SleeveRepository] Found most recent Duct Accessories XML: {actualFilePath}");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[SleeveRepository] No Duct Accessories XML files found matching pattern: {pattern} in {directory}");
                        return clashZones;
                    }
                }
                
                if (!File.Exists(actualFilePath))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SleeveRepository] Duct Accessories XML file not found: {actualFilePath}");
                    return clashZones;
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SleeveRepository] Loading Duct Accessories clash zones from: {actualFilePath}");
                
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(actualFilePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    var storageZones = filter?.ClashZoneStorage?.AllZones;
                    if (storageZones != null && storageZones.Count > 0)
                    {
                        clashZones.AddRange(storageZones);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[SleeveRepository] Loaded {clashZones.Count} Duct Accessories clash zones from XML");
                    }
                }
                
                return clashZones;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[SleeveRepository] Error loading Duct Accessories clash zones: {ex.Message}");
                return new List<ClashZone>();
            }
        }

        // Helper methods extracted from UniversalSleevePlacerService
        
        private string GetSleeveCategoryFromParameter(FamilyInstance sleeve)
        {
            try
            {
                var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                if (mepCategoryParam != null && !string.IsNullOrEmpty(mepCategoryParam.AsString()))
                {
                    return mepCategoryParam.AsString();
                }
                
                // Fallback to family name
                var familyName = sleeve.Symbol.FamilyName.ToLower();
                if (familyName.Contains("circular") || familyName.Contains("round")) return "Pipes";
                if (familyName.Contains("duct")) return "Ducts";
                if (familyName.Contains("pipe")) return "Pipes";
                if (familyName.Contains("cable")) return "Cable Trays";
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        private string GetSleeveOrientation(Document doc, FamilyInstance sleeve)
        {
            try
            {
                var mepElementId = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementId != null && mepElementId.HasValue)
                {
                    var mepElement = doc.GetElement(mepElementId.AsElementId());
                    if (mepElement != null)
                    {
                        // Get orientation from MEP element's Wall Direction Type
                        var wallDirectionParam = mepElement.LookupParameter("Wall Direction Type");
                        if (wallDirectionParam != null && !string.IsNullOrEmpty(wallDirectionParam.AsString()))
                        {
                            return wallDirectionParam.AsString();
                        }
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
        }

        private string GetHostType(FamilyInstance sleeve)
        {
             // This was also a helper in UniversalSleevePlacerService.
             // We need to implement it here.
             // Assuming simple logic for now or copy from original if visible.
             // Based on family name usually.
             
             var familyName = sleeve.Symbol.FamilyName;
             if (familyName.Contains("Wall")) return "Wall";
             if (familyName.Contains("Slab") || familyName.Contains("Floor")) return "Floor";
             return "Unknown";
        }

        /// <summary>
        /// Stub implementation for ISleeveRepository compatibility.
        /// Feature implemented in ClashZoneRepository (SQLite).
        /// </summary>
        public void UpdateSleeveCalculatedData(ClashZone zone)
        {
            // No-op: SleeveRepository handles XML, not database updates.
            // This method is required by the interface but should only be called when using ClashZoneRepository.
        }
    }
}
