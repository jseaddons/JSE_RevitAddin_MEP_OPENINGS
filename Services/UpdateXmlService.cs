using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using WinForms = System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to update XML files after manual cluster sleeve adjustments
    /// This is an expensive operation that scans all cluster sleeves and updates XML
    /// </summary>
    public class UpdateXmlService
    {
        /// <summary>
        /// Update XML files with current cluster sleeve information from Revit
        /// </summary>
        public void UpdateXmlFromRevit(Document doc)
        {
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateXmlService] Starting XML update from Revit model");
                
                // Get filters directory using helper
                ProjectPathService.EnsureFiltersDirectory(doc);

                string filtersDirectory = ProjectPathService.GetFiltersDirectory(doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    throw new DirectoryNotFoundException($"Filters directory not found: {filtersDirectory}");
                }

                // Get all XML filter files (excluding CONDITIONS and _global files)
                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                var filterFiles = allXmlFiles.Where(f =>
                {
                    string fileName = Path.GetFileName(f);
                    return !fileName.EndsWith("_global.xml", StringComparison.OrdinalIgnoreCase) &&
                           !fileName.EndsWith("_CONDITIONS.xml", StringComparison.OrdinalIgnoreCase) &&
                           !fileName.EndsWith("_conditions.xml", StringComparison.OrdinalIgnoreCase);
                }).ToList();

                if (filterFiles.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UpdateXmlService] No filter XML files found in {filtersDirectory}");
                    return;
                }

                int totalUpdated = 0;
                
                // Process each filter file
                foreach (var xmlFile in filterFiles)
                {
                    try
                    {
                        int updated = UpdateFilterFile(xmlFile, doc);
                        totalUpdated += updated;
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UpdateXmlService] Updated {updated} clash zones in {Path.GetFileName(xmlFile)}");
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UpdateXmlService] Error processing {Path.GetFileName(xmlFile)}: {ex.Message}");
                        // Continue with other files
                    }
                }

                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateXmlService] ✅ XML update complete: {totalUpdated} total clash zones updated");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UpdateXmlService] Error updating XML from Revit: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        /// <summary>
        /// Update a single filter XML file with current cluster sleeve data
        /// </summary>
        private int UpdateFilterFile(string xmlFilePath, Document doc)
        {
            int updatedCount = 0;
            
            try
            {
                // ✅ Use FilterManagementService helper to load XML
                var filterMgmt = new FilterManagementService(
                    msg => {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UpdateXmlService] {msg}");
                    },
                    msg => { } // No status updates needed
                );
                
                var filter = filterMgmt.LoadFilterFromXmlFile(xmlFilePath);
                var storageZones = filter?.ClashZoneStorage?.AllZones ?? new List<ClashZone>();

                if (storageZones.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UpdateXmlService] No clash zones in {Path.GetFileName(xmlFilePath)}");
                    return 0;
                }

                var clusterSleeveGroups = storageZones
                    .Where(cz => cz.IsClusterResolved && cz.ClusterSleeveInstanceId > 0)
                    .GroupBy(cz => cz.MepElementCategory)
                    .ToList();

                if (clusterSleeveGroups.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UpdateXmlService] No cluster sleeves found in {Path.GetFileName(xmlFilePath)}");
                    return 0;
                }

                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateXmlService] Processing {clusterSleeveGroups.Count} category groups in {Path.GetFileName(xmlFilePath)}");

                // Process each category group
                foreach (var categoryGroup in clusterSleeveGroups)
                {
                    string category = categoryGroup.Key;
                    var clusterSleeveIds = categoryGroup
                        .Select(cz => cz.ClusterSleeveInstanceId)
                        .Distinct()
                        .ToList();

                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UpdateXmlService] Processing {clusterSleeveIds.Count} cluster sleeves for category '{category}'");

                    // Process each cluster sleeve
                    foreach (var clusterSleeveId in clusterSleeveIds)
                    {
                        try
                        {
                            var clusterSleeve = doc.GetElement(new ElementId(clusterSleeveId)) as FamilyInstance;
                            
                            if (clusterSleeve == null)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[UpdateXmlService] Cluster sleeve {clusterSleeveId} not found in Revit - may have been deleted");
                                continue;
                            }

                            // Get MEP elements intersecting this cluster sleeve
                            var mepElements = GetMepElementsIntersectingSleeve(doc, clusterSleeve, category);
                            
                            // Get sleeve dimensions
                            var (width, height, diameter) = GetSleeveDimensions(clusterSleeve);
                            var bbox = clusterSleeve.get_BoundingBox(null);
                            
                            // Update all clash zones for this cluster sleeve
                            var clashZonesForCluster = filter.ClashZoneStorage.AllZones
                                .Where(cz => cz.ClusterSleeveInstanceId == clusterSleeveId)
                                .ToList();

                            foreach (var clashZone in clashZonesForCluster)
                            {
                                // ✅ BUG FIX #2: Only update CLUSTER sleeve bounding box, NOT individual sleeve bounding box
                                // Individual sleeve bounding boxes should be preserved (they represent the original individual sleeves)
                                // Only cluster sleeve bounding boxes should be updated from Revit
                                if (bbox != null)
                                {
                                    // ❌ REMOVED: Do NOT overwrite individual sleeve bounding box with cluster sleeve's bbox
                                    // clashZone.SleeveBoundingBoxMinX = bbox.Min.X;  // This was wrong!
                                    // The individual sleeve bbox should remain as-is (or be zero if individual sleeve was deleted)
                                    
                                    // ✅ CORRECT: Only update cluster sleeve bounding box
                                    clashZone.ClusterSleeveBoundingBoxMinX = bbox.Min.X;
                                    clashZone.ClusterSleeveBoundingBoxMinY = bbox.Min.Y;
                                    clashZone.ClusterSleeveBoundingBoxMinZ = bbox.Min.Z;
                                    clashZone.ClusterSleeveBoundingBoxMaxX = bbox.Max.X;
                                    clashZone.ClusterSleeveBoundingBoxMaxY = bbox.Max.Y;
                                    clashZone.ClusterSleeveBoundingBoxMaxZ = bbox.Max.Z;
                                }
                                
                                clashZone.SleeveWidth = width;
                                clashZone.SleeveHeight = height;
                                clashZone.SleeveDiameter = diameter;
                                
                                // Update MEP element information if found
                                if (mepElements.Count > 0)
                                {
                                    // For cluster sleeves, aggregate MEP parameters from all intersecting elements
                                    var aggregatedParams = AggregateMepParameters(mepElements);
                                    clashZone.MepParameterValues = aggregatedParams;
                                    
                                    // Use first MEP element's ID (or aggregate if multiple)
                                    if (mepElements.Count == 1)
                                    {
                                        clashZone.MepElementIdValue = mepElements[0].Id.GetIntegerValue();
                                    }
                                    else
                                    {
                                        // Multiple MEP elements - keep original or use first
                                        if (clashZone.MepElementIdValue > 0)
                                        {
                                            // Keep original for reference, but note multiple elements
                                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[UpdateXmlService] Cluster sleeve {clusterSleeveId} has {mepElements.Count} MEP elements - keeping original MEP_ElementId");
                                        }
                                        else
                                        {
                                            clashZone.MepElementIdValue = mepElements[0].Id.GetIntegerValue();
                                        }
                                    }
                                }
                                
                                clashZone.LastUpdated = DateTime.Now;
                                updatedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[UpdateXmlService] Error updating cluster sleeve {clusterSleeveId}: {ex.Message}");
                            // Continue with next cluster sleeve
                        }
                    }
                }

                // ✅ Use FilterManagementService helper to save XML
                filterMgmt.SaveFilterToXmlFile(filter, xmlFilePath);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateXmlService] ✅ Saved {updatedCount} updated clash zones to {Path.GetFileName(xmlFilePath)}");
                
                return updatedCount;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UpdateXmlService] Error updating filter file {xmlFilePath}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Get MEP elements that intersect with a cluster sleeve
        /// </summary>
        private List<Element> GetMepElementsIntersectingSleeve(Document doc, FamilyInstance sleeve, string category)
        {
            var mepElements = new List<Element>();
            
            try
            {
                var bbox = sleeve.get_BoundingBox(null);
                if (bbox == null) return mepElements;

                // Expand bounding box slightly to catch nearby elements
                double tolerance = 0.1; // ~30mm in feet
                var expandedMin = new XYZ(bbox.Min.X - tolerance, bbox.Min.Y - tolerance, bbox.Min.Z - tolerance);
                var expandedMax = new XYZ(bbox.Max.X + tolerance, bbox.Max.Y + tolerance, bbox.Max.Z + tolerance);
                var expandedBbox = new BoundingBoxXYZ { Min = expandedMin, Max = expandedMax };
                
                // Determine MEP categories based on filter category
                var mepCategories = new List<BuiltInCategory>();
                if (category.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    mepCategories.Add(BuiltInCategory.OST_DuctCurves);
                    mepCategories.Add(BuiltInCategory.OST_DuctFitting);
                    mepCategories.Add(BuiltInCategory.OST_DuctAccessory);
                }
                else if (category.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    mepCategories.Add(BuiltInCategory.OST_PipeCurves);
                    mepCategories.Add(BuiltInCategory.OST_PipeFitting);
                    mepCategories.Add(BuiltInCategory.OST_PipeAccessory);
                }
                else if (category.IndexOf("Cable", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    mepCategories.Add(BuiltInCategory.OST_CableTray);
                    mepCategories.Add(BuiltInCategory.OST_CableTrayFitting);
                    mepCategories.Add(BuiltInCategory.OST_Conduit);
                    mepCategories.Add(BuiltInCategory.OST_ConduitFitting);
                }

                // Collect MEP elements in expanded bounding box
                var outline = new Outline(expandedMin, expandedMax);
                var filter = new BoundingBoxIntersectsFilter(outline);
                var multiFilter = new ElementMulticategoryFilter(mepCategories);
                var combinedFilter = new LogicalAndFilter(filter, multiFilter);

                mepElements = new FilteredElementCollector(doc)
                    .WherePasses(combinedFilter)
                    .WhereElementIsNotElementType()
                    .ToList();

                // Filter to only elements that actually intersect the sleeve geometry
                var intersectingElements = new List<Element>();
                var sleeveGeometry = sleeve.get_Geometry(new Options { ComputeReferences = true, IncludeNonVisibleObjects = false });
                
                foreach (var element in mepElements)
                {
                    if (DoesElementIntersectSleeve(element, sleeveGeometry, bbox))
                    {
                        intersectingElements.Add(element);
                    }
                }

                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateXmlService] Found {intersectingElements.Count} MEP elements intersecting cluster sleeve {sleeve.Id.GetIntegerValue()}");
                return intersectingElements;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UpdateXmlService] Error getting MEP elements for sleeve: {ex.Message}");
                return mepElements;
            }
        }

        /// <summary>
        /// Check if an MEP element actually intersects the sleeve geometry
        /// </summary>
        private bool DoesElementIntersectSleeve(Element mepElement, GeometryElement sleeveGeometry, BoundingBoxXYZ sleeveBbox)
        {
            try
            {
                // Quick bounding box check first
                var mepBbox = mepElement.get_BoundingBox(null);
                if (mepBbox == null) return false;

                if (!BoundingBoxService.BoundingBoxesIntersect(mepBbox, sleeveBbox))
                {
                    return false;
                }

                // Simple check: if element center is within sleeve bbox, consider it intersecting
                var mepCenter = (mepBbox.Min + mepBbox.Max) / 2.0;
                if (mepCenter.X >= sleeveBbox.Min.X && mepCenter.X <= sleeveBbox.Max.X &&
                    mepCenter.Y >= sleeveBbox.Min.Y && mepCenter.Y <= sleeveBbox.Max.Y &&
                    mepCenter.Z >= sleeveBbox.Min.Z && mepCenter.Z <= sleeveBbox.Max.Z)
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // ✅ OOP REFACTORING: Removed duplicate BoundingBoxesIntersect - now uses BoundingBoxService.BoundingBoxesIntersect()

        /// <summary>
        /// Get sleeve dimensions from Revit element
        /// </summary>
        private (double width, double height, double diameter) GetSleeveDimensions(FamilyInstance sleeve)
        {
            try
            {
                double width = 0;
                double height = 0;
                double diameter = 0;

                var widthParam = sleeve.LookupParameter("Width");
                var heightParam = sleeve.LookupParameter("Height");
                var depthParam = sleeve.LookupParameter("Depth");
                var diameterParam = sleeve.LookupParameter("Diameter");

                if (widthParam != null && widthParam.HasValue)
                    width = widthParam.AsDouble();
                if (heightParam != null && heightParam.HasValue)
                    height = heightParam.AsDouble();
                if (depthParam != null && depthParam.HasValue)
                    diameter = depthParam.AsDouble(); // Diameter stored in Depth parameter for some families
                if (diameterParam != null && diameterParam.HasValue && diameter == 0)
                    diameter = diameterParam.AsDouble();

                return (width, height, diameter);
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UpdateXmlService] Error getting sleeve dimensions: {ex.Message}");
                return (0, 0, 0);
            }
        }

        /// <summary>
        /// Aggregate MEP parameters from multiple elements
        /// </summary>
        private List<SerializableKeyValue> AggregateMepParameters(List<Element> mepElements)
        {
            var aggregatedParams = new Dictionary<string, string>();
            
            try
            {
                // Parameter names to extract
                var paramNames = new[] { "Size", "MEP Size", "Diameter", "System Type", "System Abbreviation", "Service Type" };
                
                foreach (var element in mepElements)
                {
                    foreach (var paramName in paramNames)
                    {
                        var param = element.LookupParameter(paramName);
                        if (param != null && param.HasValue)
                        {
                            string value = param.AsString() ?? param.AsValueString() ?? "";
                            
                            if (!string.IsNullOrEmpty(value))
                            {
                                // For Size parameter: aggregate all values (including duplicates)
                                if (paramName.Equals("Size", StringComparison.OrdinalIgnoreCase) || 
                                    paramName.Equals("MEP Size", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (aggregatedParams.ContainsKey(paramName))
                                    {
                                        aggregatedParams[paramName] = $"{aggregatedParams[paramName]}, {value}";
                                    }
                                    else
                                    {
                                        aggregatedParams[paramName] = value;
                                    }
                                }
                                else
                                {
                                    // For other parameters: only add if not already present (avoid duplicates)
                                    if (!aggregatedParams.ContainsKey(paramName))
                                    {
                                        aggregatedParams[paramName] = value;
                                    }
                                }
                            }
                        }
                    }
                }

                return aggregatedParams.Select(kvp => new SerializableKeyValue { Key = kvp.Key, Value = kvp.Value }).ToList();
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UpdateXmlService] Error aggregating MEP parameters: {ex.Message}");
                return new List<SerializableKeyValue>();
            }
        }
    }
}
