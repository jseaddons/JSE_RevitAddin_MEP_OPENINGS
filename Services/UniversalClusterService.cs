using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Unit = Autodesk.Revit.DB.UnitUtils;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Universal clustering service - extracted from RectangularSleeveClusterCommandV2
    /// Can be called from any context (ICommand, IExternalCommand, etc.)
    /// </summary>
    public class UniversalClusterService
    {
        // Helper struct for grouping key
        private struct SleeveGroupKey
        {
            public string hostType;
            public string systemType;
            public string orientation;
            
            public SleeveGroupKey(string hostType, string systemType, string orientation)
            {
                this.hostType = hostType;
                this.systemType = systemType;
                this.orientation = orientation;
            }
            
            public override bool Equals(object obj)
            {
                if (!(obj is SleeveGroupKey)) return false;
                var other = (SleeveGroupKey)obj;
                return hostType == other.hostType && systemType == other.systemType && orientation == other.orientation;
            }
            
            public override int GetHashCode()
            {
                return (hostType, systemType, orientation).GetHashCode();
            }
        }

        /// <summary>
        /// Cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="targetCategory">Category to cluster (e.g., "Ducts", "Pipes") or null for all</param>
        /// <param name="uiDoc">Optional UIDocument for section box filtering</param>
        /// <returns>Tuple of (placedCount, deletedCount)</returns>
        public (int placedCount, int deletedCount) ClusterSleeves(Document doc, string targetCategory, UIDocument uiDoc = null, string xmlFilePath = null)
        {
            int placedCount = 0;
            int deletedCount = 0;

            try
            {
                // Initialize dedicated cluster debug log
                string clusterLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log";
                File.AppendAllText(clusterLogPath, $"\n===== CLUSTER DEBUG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n");
                File.AppendAllText(clusterLogPath, $"Target Category: {targetCategory ?? "ALL"}\n");
                
                // ⚠️ CRITICAL: Reset cluster flags for deleted cluster sleeves
                ResetClusterFlagsForDeletedSleeves(doc, xmlFilePath);
                
                // Get cluster configuration
                double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
                if (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    toleranceMm = Math.Min(toleranceMm, 100); // clamp pipes to 100mm
                }
                double toleranceDist = UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
                
                DebugLogger.Log($"[UniversalClusterService] Using JoinOpeningsDistance: {toleranceMm}mm (from {ClusterConfigurationManager.Instance.ConfigurationSource})");
                DebugLogger.Log($"[UniversalClusterService] Internal units: {toleranceDist:F6} feet");
                DebugLogger.Log($"[UniversalClusterService] Target category: {targetCategory ?? "ALL"}");
                
                File.AppendAllText(clusterLogPath, $"JoinOpeningsDistance: {toleranceMm}mm\n");
                File.AppendAllText(clusterLogPath, $"Target category: {targetCategory ?? "ALL"}\n");

                // Collect all sleeves using the 4 universal families
                var allSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => 
                    {
                        var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                        return famName == "RectangularOpeningOnWall" ||
                               famName == "CircularOpeningOnWall" ||
                               famName == "RectangularOpeningOnSlab" ||
                               famName == "CircularOpeningOnSlab";
                    })
                    .ToList();
                
                DebugLogger.Log($"[UniversalClusterService] Found {allSleeves.Count} total sleeves (all categories)");
                File.AppendAllText(clusterLogPath, $"Found {allSleeves.Count} total sleeves (all categories)\n");
                
                // Filter by active category when provided
                var rawSleeves = allSleeves;
                if (!string.IsNullOrEmpty(targetCategory))
                {
                    if (targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        rawSleeves = allSleeves.Where(s => GetCategoryFromMepElementId(s).IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                        DebugLogger.Log($"[UniversalClusterService] Filtered sleeves for Pipes only: {rawSleeves.Count}");
                    }
                    else if (targetCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        rawSleeves = allSleeves.Where(s => GetCategoryFromMepElementId(s).IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                        DebugLogger.Log($"[UniversalClusterService] Filtered sleeves for Ducts only: {rawSleeves.Count}");
                    }
                    else if (targetCategory.IndexOf("Cable", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        rawSleeves = allSleeves.Where(s => GetCategoryFromMepElementId(s).IndexOf("Cable", StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                        DebugLogger.Log($"[UniversalClusterService] Filtered sleeves for Cable Trays only: {rawSleeves.Count}");
                    }
                }
                
                DebugLogger.Log($"[UniversalClusterService] Processing {rawSleeves.Count} sleeves" + 
                               (string.IsNullOrEmpty(targetCategory) ? " (all categories)" : $" (filtering will happen during grouping)"));
                File.AppendAllText(clusterLogPath, $"Processing {rawSleeves.Count} sleeves\n");

                // Use SectionBoxHelper to reduce to only elements visible in the active 3D section box (if UIDocument provided)
                List<FamilyInstance> sleeves;
                if (uiDoc != null)
                {
                    try
                    {
                        var rawElements = rawSleeves.Cast<Element>().Select(e => (element: (Element)e, transform: (Transform?)null)).ToList();
                        var filtered = SectionBoxHelper.FilterElementsBySectionBox(uiDoc, rawElements);
                        sleeves = filtered.Select(t => t.element).OfType<FamilyInstance>().ToList();
                        DebugLogger.Log($"[UniversalClusterService] Raw sleeves={rawSleeves.Count}, Filtered by section box={sleeves.Count}");
                        
                        // Fallback to raw collection if section box filtering yields zero results
                        if (sleeves.Count == 0 && rawSleeves.Count > 0)
                        {
                            DebugLogger.Log($"[UniversalClusterService] Section-box filtering yielded 0 results; falling back to raw collection of {rawSleeves.Count} sleeves.");
                            sleeves = rawSleeves;
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[UniversalClusterService] SectionBox filtering failed: {ex.Message}; falling back to raw collection");
                        sleeves = rawSleeves;
                    }
                }
                else
                {
                    sleeves = rawSleeves;
                    DebugLogger.Log($"[UniversalClusterService] No UIDocument provided, skipping section box filtering");
                }

                if (sleeves.Count == 0)
                {
                    DebugLogger.Log($"[UniversalClusterService] No sleeves to cluster");
                    return (0, 0);
                }

                // ⚠️ CRITICAL: Transaction must be started by caller
                if (!doc.IsModifiable)
                {
                    DebugLogger.Error($"[UniversalClusterService] Document is not modifiable - transaction must be started by caller");
                    throw new InvalidOperationException("Document must be in a transaction before calling ClusterSleeves");
                }

                // Group sleeves by host type, system type, and orientation
                var sleeveGroups = sleeves.GroupBy(sleeve => {
                    var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
                    string effectiveOrientation = hostOrientationParam != null ? hostOrientationParam.AsString() : "";
                    
                    var famName = sleeve.Symbol.Family.Name.ToLower();
                    string hostType = famName.Contains("onwall") ? "Wall" : (famName.Contains("onslab") || famName.Contains("onfloor") ? "Floor" : "Unknown");
                    
                    // Get category from MEP_ElementId by looking up in XML files (same approach as MarkParameterService)
                    string systemType = GetCategoryFromMepElementId(sleeve);
                    
                    // Log to dedicated cluster debug file
                    File.AppendAllText(clusterLogPath, $"Sleeve {sleeve.Id}: hostType={hostType}, systemType={systemType}, orientation={effectiveOrientation}\n");
                    
                    return new SleeveGroupKey(hostType, systemType, effectiveOrientation);
                });

                // Log group diagnostics
                foreach (var g in sleeveGroups)
                {
                    var list = g.ToList();
                    var ids = list.Select(fi => fi.Id.IntegerValue.ToString()).Take(10).ToList();
                    DebugLogger.Log($"[ClusterService] Group hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}");
                    File.AppendAllText(clusterLogPath, $"Group: hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}\n");
                }

                // Form clusters using spatial hashing
                var clustersByGroup = FormClusters(sleeveGroups, toleranceDist);

                // Process each cluster
                foreach (var groupEntry in clustersByGroup)
                {
                    var groupKey = groupEntry.Key;
                    var clusters = groupEntry.Value;

                    foreach (var cluster in clusters)
                    {
                        if (cluster.Count <= 1) continue; // Skip individual sleeves

                        try
                        {
                            // ⚠️ CRITICAL: Check if cluster already exists using flag management
                            if (IsClusterAlreadyExists(cluster, groupKey))
                            {
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"SKIP: Cluster already exists for {cluster.Count} sleeves - skipping placement\n");
                                continue;
                            }

                            // Place cluster sleeve
                            PlaceClusterSleeve(doc, cluster, groupKey, targetCategory, out int placed1, out int deleted1, xmlFilePath);
                            placedCount += placed1;
                            deletedCount += deleted1;
                            
                            // Marking happens inside PlaceClusterSleeve with actual cluster sleeve ID
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Error($"[UniversalClusterService] Error placing cluster: {ex.Message}");
                        }
                    }
                }

                DebugLogger.Log($"[UniversalClusterService] Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.\n");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"===== CLUSTER DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Clustering failed: {ex.Message}");
                DebugLogger.Error($"[UniversalClusterService] Stack trace: {ex.StackTrace}");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"ERROR: {ex.Message}\n");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Stack trace: {ex.StackTrace}\n");
                throw;
            }

            return (placedCount, deletedCount);
        }

        /// <summary>
        /// Get category from MEP Element ID by searching XML files (same as MarkParameterService)
        /// </summary>
        private string GetCategoryFromMepElementId(FamilyInstance sleeve)
        {
            try
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();

                    // Find matching clash zone in XML files to get category
                    var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");

                    if (Directory.Exists(filtersDirectory))
                    {
                        var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                        foreach (var xmlFile in xmlFiles)
                        {
                            try
                            {
                                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                                using (var reader = new StreamReader(xmlFile))
                                {
                                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                                    if (filter?.ClashZoneStorage?.ClashZones != null)
                                    {
                                        foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                        {
                                            if (clashZone.MepElementId.IntegerValue == mepElementId)
                                            {
                                                return clashZone.MepElementCategory;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // Ignore deserialization errors for individual files
                                continue;
                            }
                        }
                    }

                    DebugLogger.Warning($"[UniversalClusterService] No category found for MEP Element ID {mepElementId}");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"WARNING: No category found for MEP Element ID {mepElementId} on sleeve {sleeve.Id}\n");
                    return "Unknown";
                }
                else
                {
                    DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id} missing MEP_ElementId parameter");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"WARNING: Sleeve {sleeve.Id} missing MEP_ElementId parameter\n");
                    return "Unknown";
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting category for MEP Element ID on sleeve {sleeve.Id}: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// Check if cluster already exists using flag management (efficient)
        /// </summary>
        private bool IsClusterAlreadyExists(List<FamilyInstance> cluster, SleeveGroupKey groupKey)
        {
            try
            {
                // Check if any sleeve in the cluster is already cluster-resolved
                foreach (var sleeve in cluster)
                {
                    var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                    if (mepElementIdParam != null)
                    {
                        long mepElementId = mepElementIdParam.AsInteger();
                        
                        // Find clash zone for this MEP element
                        var clashZone = GetClashZoneByMepElementId(mepElementId);
                        if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                        {
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Cluster already exists: ClashZone {clashZone.Id} is cluster-resolved (cluster sleeve ID: {clashZone.ClusterSleeveInstanceId})\n");
                            return true;
                        }
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error checking if cluster exists: {ex.Message}");
                return false; // Default to allowing placement if check fails
            }
        }

        /// <summary>
        /// Mark clash zones as cluster-resolved after successful cluster placement
        /// </summary>
        private void MarkClashZonesAsClusterResolved(List<FamilyInstance> cluster)
        {
            // This method is called from the main clustering loop - actual marking happens in PlaceClusterSleeve
            // with the correct cluster sleeve ID
        }

        /// <summary>
        /// Mark clash zones as cluster-resolved with the actual cluster sleeve ID
        /// </summary>
        private void MarkClashZonesAsClusterResolvedWithSleeveId(List<FamilyInstance> cluster, ElementId clusterSleeveId, string xmlFilePath = null)
        {
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                    return;

                // ✅ FIX: Only process specific XML file if provided (ONE SOURCE OF TRUTH)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")  // Backward compatibility
                    : new[] { xmlFilePath };  // Only the specific file
                    
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[MarkClusterResolved] Updating cluster flags in {xmlFiles.Length} XML file(s): {(string.IsNullOrEmpty(xmlFilePath) ? "ALL" : Path.GetFileName(xmlFilePath))} for cluster sleeve {clusterSleeveId.IntegerValue}\n");
                
                int markedCount = 0;

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        OpeningFilter filter = null;
                        
                        // Read existing filter
                        using (var reader = new StreamReader(xmlFile))
                        {
                            filter = (OpeningFilter)serializer.Deserialize(reader);
                        }

                        if (filter?.ClashZoneStorage?.ClashZones != null)
                        {
                            bool updated = false;
                            
                            foreach (var sleeve in cluster)
                            {
                                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                                if (mepElementIdParam != null)
                                {
                                    long mepElementId = mepElementIdParam.AsInteger();
                                    
                                    // Find and mark clash zone as cluster-resolved
                                    var clashZone = filter.ClashZoneStorage.ClashZones.FirstOrDefault(cz => cz.MepElementId.IntegerValue == mepElementId);
                                    if (clashZone != null)
                                    {
                                        clashZone.IsClusterResolved = true;
                                        clashZone.ClusterSleeveId = clusterSleeveId;
                                        clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue; // ✅ FIX: Store integer for XML serialization
                                        clashZone.LastUpdated = DateTime.Now;
                                        updated = true;
                                        markedCount++;
                                        
                                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Marked ClashZone {clashZone.Id} as cluster-resolved with cluster sleeve {clusterSleeveId.IntegerValue}\n");
                                    }
                                }
                            }
                            
                            // Save updated XML if changes were made
                            if (updated)
                            {
                                using (var writer = new StreamWriter(xmlFile))
                                {
                                    serializer.Serialize(writer, filter);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalClusterService] Error processing XML file {xmlFile}: {ex.Message}");
                    }
                }

                if (markedCount > 0)
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Marked {markedCount} clash zones as cluster-resolved\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error marking clash zones as cluster-resolved: {ex.Message}");
            }
        }

        /// <summary>
        /// Reset cluster flags for deleted cluster sleeves
        /// </summary>
        private void ResetClusterFlagsForDeletedSleeves(Document doc, string xmlFilePath = null)
        {
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                    return;

                // ✅ FIX: Only process specific XML file if provided (ONE SOURCE OF TRUTH)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")  // Backward compatibility
                    : new[] { xmlFilePath };  // Only the specific file
                    
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[ResetFlags] Processing {xmlFiles.Length} XML file(s): {(string.IsNullOrEmpty(xmlFilePath) ? "ALL" : Path.GetFileName(xmlFilePath))}\n");
                
                int resetCount = 0;

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        OpeningFilter filter = null;
                        
                        // ⚠️ CRITICAL: Read XML first, then close reader before writing
                        using (var reader = new StreamReader(xmlFile))
                        {
                            filter = (OpeningFilter)serializer.Deserialize(reader);
                        } // Reader is now closed
                        
                        if (filter?.ClashZoneStorage?.ClashZones != null)
                        {
                            bool modified = false;
                            
                            foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                            {
                                if (clashZone.IsClusterResolved)
                                {
                                    // ✅ FIX: Check integer version (ClusterSleeveInstanceId) which IS serialized to XML
                                    if (clashZone.ClusterSleeveInstanceId <= 0)
                                    {
                                        clashZone.IsClusterResolved = false;
                                        clashZone.ClusterSleeveId = null;
                                        clashZone.ClusterSleeveInstanceId = -1;
                                        clashZone.LastUpdated = DateTime.Now;
                                        resetCount++;
                                        modified = true;
                                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Reset cluster flag for ClashZone {clashZone.Id} - ClusterSleeveInstanceId was {clashZone.ClusterSleeveInstanceId} (invalid)\n");
                                    }
                                    else
                                    {
                                        // Check if cluster sleeve still exists in Revit
                                        var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
                                        var clusterSleeve = doc.GetElement(clusterSleeveId);
                                        if (clusterSleeve == null)
                                        {
                                            // Cluster sleeve was deleted - reset flag
                                            clashZone.IsClusterResolved = false;
                                            clashZone.ClusterSleeveId = null;
                                            clashZone.ClusterSleeveInstanceId = -1;
                                            clashZone.LastUpdated = DateTime.Now;
                                            resetCount++;
                                            modified = true;
                                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Reset cluster flag for ClashZone {clashZone.Id} - cluster sleeve {clashZone.ClusterSleeveInstanceId} was deleted\n");
                                        }
                                    }
                                }
                            }
                            
                            // ✅ FIX: Only save if modifications were made, and reader is already closed
                            if (modified)
                            {
                                using (var writer = new StreamWriter(xmlFile))
                                {
                                    serializer.Serialize(writer, filter);
                                }
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✓ Saved reset flags to {Path.GetFileName(xmlFile)}\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalClusterService] Error processing XML file {xmlFile}: {ex.Message}");
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✗ Error resetting flags in {Path.GetFileName(xmlFile)}: {ex.Message}\n");
                    }
                }

                if (resetCount > 0)
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Reset cluster flags for {resetCount} deleted cluster sleeves\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error resetting cluster flags: {ex.Message}");
            }
        }

        /// <summary>
        /// Get clash zone by MEP element ID from XML files
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, string xmlFilePath = null)
        {
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                    return null;

                // ✅ FIX: Only search specific XML file if provided (ONE SOURCE OF TRUTH)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")  // Backward compatibility
                    : new[] { xmlFilePath };  // Only the specific file

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                {
                                    if (clashZone.MepElementId.IntegerValue == mepElementId)
                                    {
                                        return clashZone;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting clash zone for MEP Element ID {mepElementId}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Load a universal opening family from the Resources folder
        /// </summary>
        private bool LoadUniversalFamily(Document doc, string familyName)
        {
            try
            {
                string resourcesPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources";
                string familyPath = Path.Combine(resourcesPath, $"{familyName}.rfa");

                DebugLogger.Log($"[ClusterService] Attempting to load family from: {familyPath}");

                if (!File.Exists(familyPath))
                {
                    DebugLogger.Error($"[ClusterService] Universal family file not found: {familyPath}");
                    return false;
                }

                // Load the family
                bool loaded = doc.LoadFamily(familyPath);

                if (loaded)
                {
                    DebugLogger.Log($"[ClusterService] Successfully loaded universal family: {familyName}");
                    return true;
                }
                else
                {
                    DebugLogger.Error($"[ClusterService] Failed to load universal family: {familyName}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ClusterService] Error loading universal family '{familyName}': {ex.Message}");
                return false;
            }
        }

        private Dictionary<SleeveGroupKey, List<List<FamilyInstance>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, FamilyInstance>> sleeveGroups, 
            double toleranceDist)
        {
            var clustersByGroup = new Dictionary<SleeveGroupKey, List<List<FamilyInstance>>>();
            double cellSize = toleranceDist > 0.0 ? toleranceDist : Unit.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

            foreach (var group in sleeveGroups)
            {
                var familySleeves = group.ToList();
                var centers = new Dictionary<FamilyInstance, XYZ>(familySleeves.Count);
                var bboxes = new Dictionary<FamilyInstance, BoundingBoxXYZ>(familySleeves.Count);

                // Precompute centers and bboxes
                foreach (var s in familySleeves)
                {
                    var center = (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin;
                    centers[s] = center;
                    try { var bb = s.get_BoundingBox(null); if (bb != null) bboxes[s] = bb; } catch { }
                }

                // Build spatial hash grid
                var grid = BuildSpatialGrid(familySleeves, bboxes, centers, cellSize);

                // Form clusters using BFS
                var groupClusters = FormClustersFromGrid(familySleeves, grid, bboxes, centers, cellSize, toleranceDist);
                clustersByGroup[group.Key] = groupClusters;
            }

            return clustersByGroup;
        }

        private Dictionary<(int, int, int), List<FamilyInstance>> BuildSpatialGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize)
        {
            var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();

            foreach (var s in familySleeves)
            {
                BoundingBoxXYZ? bb = null;
                try { bb = bboxes.ContainsKey(s) ? bboxes[s] : s.get_BoundingBox(null); } catch { }

                if (bb != null)
                {
                    int min_ix = (int)Math.Floor(bb.Min.X / cellSize);
                    int max_ix = (int)Math.Floor(bb.Max.X / cellSize);
                    int min_iy = (int)Math.Floor(bb.Min.Y / cellSize);
                    int max_iy = (int)Math.Floor(bb.Max.Y / cellSize);
                    int min_iz = (int)Math.Floor(bb.Min.Z / cellSize);
                    int max_iz = (int)Math.Floor(bb.Max.Z / cellSize);

                    for (int gx = min_ix; gx <= max_ix; gx++)
                        for (int gy = min_iy; gy <= max_iy; gy++)
                            for (int gz = min_iz; gz <= max_iz; gz++)
                            {
                                var key = (gx, gy, gz);
                                if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                                list.Add(s);
                            }
                }
                else
                {
                    // Fallback to center-based bucketing
                    var c = centers[s];
                    int ix = (int)Math.Floor(c.X / cellSize);
                    int iy = (int)Math.Floor(c.Y / cellSize);
                    int iz = (int)Math.Floor(c.Z / cellSize);
                    var key = (ix, iy, iz);
                    if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                    list.Add(s);
                }
            }

            return grid;
        }

        private List<List<FamilyInstance>> FormClustersFromGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<(int, int, int), List<FamilyInstance>> grid,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize,
            double toleranceDist)
        {
            var groupClusters = new List<List<FamilyInstance>>();
            var unprocessedSet = new HashSet<FamilyInstance>(familySleeves);

            while (unprocessedSet.Count > 0)
            {
                var start = unprocessedSet.First();
                var queue = new Queue<FamilyInstance>();
                var cluster = new List<FamilyInstance>();
                queue.Enqueue(start);
                unprocessedSet.Remove(start);

                while (queue.Count > 0)
                {
                    var inst = queue.Dequeue();
                    cluster.Add(inst);

                    BoundingBoxXYZ o1_bbox = bboxes.ContainsKey(inst) ? bboxes[inst] : inst.get_BoundingBox(null);
                    var candidates = GetCandidatesFromGrid(inst, o1_bbox, centers, grid, cellSize, toleranceDist);
                    var neighbors = FilterNeighborsByBoundingBox(inst, candidates, o1_bbox, bboxes, unprocessedSet, toleranceDist);

                    foreach (var n in neighbors)
                    {
                        if (unprocessedSet.Remove(n)) queue.Enqueue(n);
                    }
                }

                groupClusters.Add(cluster);
            }

            return groupClusters;
        }

        private List<FamilyInstance> GetCandidatesFromGrid(
            FamilyInstance inst,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, XYZ> centers,
            Dictionary<(int, int, int), List<FamilyInstance>> grid,
            double cellSize,
            double toleranceDist)
        {
            var candidates = new List<FamilyInstance>();

            if (o1_bbox != null)
            {
                double exMinX = o1_bbox.Min.X - toleranceDist;
                double exMaxX = o1_bbox.Max.X + toleranceDist;
                double exMinY = o1_bbox.Min.Y - toleranceDist;
                double exMaxY = o1_bbox.Max.Y + toleranceDist;
                double exMinZ = o1_bbox.Min.Z - toleranceDist;
                double exMaxZ = o1_bbox.Max.Z + toleranceDist;

                int min_ix = (int)Math.Floor(exMinX / cellSize);
                int max_ix = (int)Math.Floor(exMaxX / cellSize);
                int min_iy = (int)Math.Floor(exMinY / cellSize);
                int max_iy = (int)Math.Floor(exMaxY / cellSize);
                int min_iz = (int)Math.Floor(exMinZ / cellSize);
                int max_iz = (int)Math.Floor(exMaxZ / cellSize);

                for (int gx = min_ix; gx <= max_ix; gx++)
                    for (int gy = min_iy; gy <= max_iy; gy++)
                        for (int gz = min_iz; gz <= max_iz; gz++)
                        {
                            var key = (gx, gy, gz);
                            if (grid.TryGetValue(key, out var bucket)) candidates.AddRange(bucket);
                        }
            }
            else
            {
                // Fallback to 3x3x3 neighbor search
                var c = centers[inst];
                int ix = (int)Math.Floor(c.X / cellSize);
                int iy = (int)Math.Floor(c.Y / cellSize);
                int iz = (int)Math.Floor(c.Z / cellSize);

                for (int dx = -1; dx <= 1; dx++)
                    for (int dy = -1; dy <= 1; dy++)
                        for (int dz = -1; dz <= 1; dz++)
                        {
                            var key = (ix + dx, iy + dy, iz + dz);
                            if (grid.TryGetValue(key, out var bucket)) candidates.AddRange(bucket);
                        }
            }

            return candidates;
        }

        private List<FamilyInstance> FilterNeighborsByBoundingBox(
            FamilyInstance inst,
            List<FamilyInstance> candidates,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            HashSet<FamilyInstance> unprocessedSet,
            double toleranceDist)
        {
            var neighbors = new List<FamilyInstance>();

            foreach (var s in candidates)
            {
                if (!unprocessedSet.Contains(s)) continue;
                if (s == inst) continue;

                var o2_bbox = bboxes.ContainsKey(s) ? bboxes[s] : s.get_BoundingBox(null);
                if (o1_bbox == null || o2_bbox == null) continue;

                bool xOverlap = o1_bbox.Max.X >= o2_bbox.Min.X - toleranceDist && o1_bbox.Min.X <= o2_bbox.Max.X + toleranceDist;
                bool yOverlap = o1_bbox.Max.Y >= o2_bbox.Min.Y - toleranceDist && o1_bbox.Min.Y <= o2_bbox.Max.Y + toleranceDist;
                bool zOverlap = o1_bbox.Max.Z >= o2_bbox.Min.Z - toleranceDist && o1_bbox.Min.Z <= o2_bbox.Max.Z + toleranceDist;

                if (xOverlap && yOverlap && zOverlap) neighbors.Add(s);
            }

            return neighbors;
        }

        private void PlaceClusterSleeve(
            Document doc,
            List<FamilyInstance> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            out int placed,
            out int deleted,
            string xmlFilePath = null)
        {
            placed = 0;
            deleted = 0;

            // Determine if cluster is circular or rectangular
            bool isCircular = cluster.All(s => 
            {
                var fam = s.Symbol?.Family?.Name ?? "";
                return fam.Contains("Circular");
            });
            // PIPE CLUSTERS: Always rectangular regardless of member shape (legacy parity)
            // Be robust to different labels (e.g., "Pipe", "Pipes", localized)
            bool isPipeCategory = (!string.IsNullOrEmpty(groupKey.systemType) && groupKey.systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0) ||
                                   (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0);
            if (isPipeCategory)
            {
                isCircular = false;
                DebugLogger.Log("[ClusterService] For Pipes category: forcing rectangular cluster shape (legacy parity)");
            }
            
            // Select universal family based on host type and shape
            string familyName = "";
            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                familyName = isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
            }
            else if (groupKey.hostType == "Floor")
            {
                familyName = isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }
            else
            {
                DebugLogger.Log($"Unknown host type for cluster group, skipping. HostType={groupKey.hostType}");
                return;
            }

            DebugLogger.Log($"[ClusterService] Creating {groupKey.systemType} cluster using family '{familyName}' (shape: {(isCircular ? "Circular" : "Rectangular")})");

            // Find FamilySymbol
            var allClusterSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            DebugLogger.Log($"[ClusterService] Looking for family '{familyName}' - Found {allClusterSymbols.Count} symbols");

            if (allClusterSymbols.Count == 0)
            {
                DebugLogger.Error($"[ClusterService] No suitable cluster family found for name '{familyName}'");
                DebugLogger.Error($"[ClusterService] Available families in project:");
                var allFamilies = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Select(sym => sym.Family.Name)
                    .Distinct()
                    .ToList();
                foreach (var famName in allFamilies.OrderBy(f => f))
                {
                    DebugLogger.Error($"[ClusterService]   - {famName}");
                }

                // Try to load the missing universal family
                if (!LoadUniversalFamily(doc, familyName))
                {
                    DebugLogger.Error($"[ClusterService] Failed to load universal family '{familyName}'");
                    return;
                }

                // Try again to find the family after loading
                allClusterSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (allClusterSymbols.Count == 0)
                {
                    DebugLogger.Error($"[ClusterService] Still no family found for '{familyName}' after loading attempt");
                    return;
                }
            }

            var clusterSymbol = allClusterSymbols.First();
            if (!clusterSymbol.IsActive) clusterSymbol.Activate();

            // Get bounding box and midpoint
            var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(cluster);

            // Get reference level
            Level? refLevel = HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]);
            if (refLevel == null)
            {
                DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                return;
            }

            // Place cluster sleeve
            FamilyInstance inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, refLevel!, StructuralType.NonStructural);

            // Apply rotation if needed
            double rotationAngle = 0.0;
            if (groupKey.hostType == "Wall" && groupKey.orientation == "Y")
            {
                rotationAngle = Math.PI / 2;
            }

            if (rotationAngle != 0.0)
            {
                XYZ axisOrigin = mid;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
            }

            // Set size parameters
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth);

            // ⚠️ CRITICAL: Set MEP_ElementId on cluster sleeve (use first individual sleeve's MEP_ElementId)
            var firstSleeve = cluster.FirstOrDefault();
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"[MEP_ElementId] First sleeve: {firstSleeve?.Id}\n");
            
            if (firstSleeve != null)
            {
                var mepElementIdParam = firstSleeve.LookupParameter("MEP_ElementId");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"[MEP_ElementId] Parameter found on first sleeve: {mepElementIdParam != null}\n");
                
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"[MEP_ElementId] Value from first sleeve: {mepElementId}\n");
                    
                    var clusterMepElementIdParam = inst.LookupParameter("MEP_ElementId");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"[MEP_ElementId] Parameter found on cluster: {clusterMepElementIdParam != null}, ReadOnly: {clusterMepElementIdParam?.IsReadOnly}\n");
                    
                    if (clusterMepElementIdParam != null && !clusterMepElementIdParam.IsReadOnly)
                    {
                        clusterMepElementIdParam.Set(mepElementId);
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✓ Set MEP_ElementId = {mepElementId} on cluster sleeve {inst.Id}\n");
                    }
                    else
                    {
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✗ Cannot set MEP_ElementId on cluster sleeve {inst.Id} (param null or readonly)\n");
                    }
                }
                else
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✗ MEP_ElementId parameter not found on first sleeve {firstSleeve.Id}\n");
                }
            }

            placed++;

            // ⚠️ CRITICAL: Mark clash zones as cluster-resolved with the actual cluster sleeve ID
            MarkClashZonesAsClusterResolvedWithSleeveId(cluster, inst.Id, xmlFilePath);

            // Delete originals
            foreach (var s in cluster)
            {
                doc.Delete(s.Id);
                deleted++;
            }
        }

        private void SetClusterSizeParameters(
            Document doc,
            FamilyInstance inst,
            List<FamilyInstance> cluster,
            SleeveGroupKey groupKey,
            double width,
            double height,
            double depth)
        {
            var widthParam = inst.LookupParameter("Width");
            var heightParam = inst.LookupParameter("Height");
            var depthParam = inst.LookupParameter("Depth");

            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                // Map to match family created in Left view
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(height);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(depth);

                // Use host thickness for Depth
                double hostThickness = width;
                if (groupKey.hostType == "Wall")
                {
                    var wall = cluster[0].Host as Wall;
                    if (wall != null)
                    {
                        hostThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                        DebugLogger.Log($"[ClusterService] Wall thickness used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                    }
                }
                else if (groupKey.hostType == "Structural Framing")
                {
                    var framing = cluster[0].Host as FamilyInstance;
                    if (framing != null)
                    {
                        var framingType = framing.Symbol;
                        var bParam = framingType.LookupParameter("b");
                        if (bParam != null && bParam.StorageType == StorageType.Double)
                        {
                            hostThickness = bParam.AsDouble();
                            DebugLogger.Log($"[ClusterService] Framing 'b' parameter used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                        }
                    }
                }
                if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(hostThickness);
            }
            else
            {
                // For other hosts (Floor), use bounding box values
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(width);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(height);
                if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(depth);
            }
        }

        /// <summary>
        /// Load sleeve ID to category mapping from ClashZone XML files
        /// </summary>
        private Dictionary<int, string> LoadSleeveCategoryMapping(string targetCategory)
        {
            var mapping = new Dictionary<int, string>();
            
            try
            {
                var filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return mapping;
                }
                
                // Load XML files for the target category (or all if null)
                var searchPattern = string.IsNullOrEmpty(targetCategory) 
                    ? "*.xml" 
                    : $"*_{GetCategoryXmlSuffix(targetCategory)}.xml";
                
                var xmlFiles = Directory.GetFiles(filtersDirectory, searchPattern);
                
                DebugLogger.Log($"[UniversalClusterService] Loading sleeve mapping from {xmlFiles.Length} XML files (pattern: {searchPattern})");
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                {
                                    // Map SleeveInstanceId to category
                                    if (clashZone.SleeveInstanceId > 0)
                                    {
                                        mapping[clashZone.SleeveInstanceId] = clashZone.MepElementCategory;
                                    }
                                    
                                    // Also map ClusterSleeveId if exists
                                    if (clashZone.ClusterSleeveId != null && clashZone.ClusterSleeveId.IntegerValue > 0)
                                    {
                                        mapping[clashZone.ClusterSleeveId.IntegerValue] = clashZone.MepElementCategory;
                                    }
                                }
                                
                                DebugLogger.Log($"[UniversalClusterService] Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones from {Path.GetFileName(xmlFile)}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[UniversalClusterService] Error loading {Path.GetFileName(xmlFile)}: {ex.Message}");
                    }
                }
                
                DebugLogger.Log($"[UniversalClusterService] Total sleeve-to-category mappings loaded: {mapping.Count}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading sleeve category mapping: {ex.Message}");
            }
            
            return mapping;
        }
        
        /// <summary>
        /// Get XML file suffix for category (e.g., "Ducts" → "ducts")
        /// </summary>
        private string GetCategoryXmlSuffix(string category)
        {
            return MepCategoryConstants.GetXmlSuffix(category);
        }
    }
}

