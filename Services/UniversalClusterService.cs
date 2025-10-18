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
            
            // ✅ PERFORMANCE: Start timing
            var startTime = DateTime.Now;

            try
            {
                // ✅ PERFORMANCE FIX: Minimal logging - only log session start and end
                string clusterLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log";
                File.AppendAllText(clusterLogPath, $"\n===== CLUSTER SESSION STARTED {DateTime.Now:HH:mm:ss} =====\n");
                
                // 🔥 CRITICAL DEBUG: Log which XML file cluster service is working with
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 CLUSTER SERVICE WORKING WITH XML FILE: {xmlFilePath ?? "NULL"}\n");
                
                // ⚠️ CRITICAL: Reset cluster flags for deleted cluster sleeves
                ResetClusterFlagsForDeletedSleeves(doc, xmlFilePath);
                
                // Load clash zone cache once (calculate once, use many times)
                LoadClashZoneCache(xmlFilePath);
                
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
                    var famName = sleeve.Symbol.Family.Name.ToLower();
                    string hostType = famName.Contains("onwall") ? "Wall" : (famName.Contains("onslab") || famName.Contains("onfloor") ? "Floor" : "Unknown");
                    
                    // Get category from MEP_ElementId by looking up in XML files (same approach as MarkParameterService)
                    string systemType = GetCategoryFromMepElementId(sleeve);
                    
                    // ⚠️ FIX: Get orientation from ClashZone (pre-calculated during refresh)
                    string effectiveOrientation = GetOrientationFromClashZone(sleeve);
                    
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
                            
                            // ✅ CRITICAL: Update ClashZone flags after cluster placement
                            if (placed1 > 0)
                            {
                                DebugLogger.Info($"[UniversalClusterService] 🔥 FLAG UPDATE: placed1={placed1}, about to find cluster sleeve 🔥");
                                
                                // Find the cluster sleeve that was just placed (it should be the newest family instance)
                                var clusterSleeve = FindNewestClusterSleeve(doc, cluster, groupKey);
                                if (clusterSleeve != null)
                                {
                                    DebugLogger.Info($"[UniversalClusterService] 🔥 FLAG UPDATE: Found cluster sleeve {clusterSleeve.Id.IntegerValue}, calling UpdateClashZoneFlagsForCluster 🔥");
                                    UpdateClashZoneFlagsForCluster(clusterSleeve, cluster, groupKey.systemType);
                                }
                                else
                                {
                                    DebugLogger.Warning($"[UniversalClusterService] ⚠️ FLAG UPDATE: No cluster sleeve found, skipping flag update");
                                }
                            }
                            else
                            {
                                DebugLogger.Warning($"[UniversalClusterService] ⚠️ FLAG UPDATE: placed1={placed1}, no cluster sleeve placed, skipping flag update");
                            }
                            
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
                
                // ✅ PERFORMANCE: Log clustering performance
                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                DebugLogger.Info($"[UniversalClusterService] ⚡ Clustering completed in {duration.TotalSeconds:F1} seconds");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"⚡ PERFORMANCE: Clustering completed in {duration.TotalSeconds:F1} seconds\n");
                
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
        /// Get host orientation from ClashZone (pre-calculated during refresh)
        /// </summary>
        private string GetOrientationFromClashZone(FamilyInstance sleeve)
        {
            try
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();
                    
                    // Find clash zone for this MEP element
                    var clashZone = GetClashZoneByMepElementId(mepElementId);
                    if (clashZone != null)
                    {
                        return clashZone.HostOrientation ?? "";
                    }
                }
                
                return "";
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting orientation from ClashZone: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Get category from MEP Element ID using clash cache (one time calculate, use many times)
        /// ✅ PERFORMANCE: Uses cache to avoid O(n·m) XML deserialization
        /// </summary>
        private string GetCategoryFromMepElementId(FamilyInstance sleeve)
        {
            try
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();

                    // ✅ PERFORMANCE: Fast O(1) lookup using clash cache
                    if (_clashZoneCache != null && _clashZoneCache.TryGetValue(mepElementId, out ClashZone clashZone))
                    {
                        return clashZone.MepElementCategory ?? "Unknown";
                    }

                    DebugLogger.Warning($"[UniversalClusterService] No category found for MEP Element ID {mepElementId} in cache");
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"WARNING: No category found for MEP Element ID {mepElementId} on sleeve {sleeve.Id} (not in cache)\n");
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
                                        
                                        // ✅ CORRECT: Keep individual sleeve flag as TRUE when deleted by clustering
                                        // This ensures clash zone is avoided next time (both flags true = avoid)
                                        clashZone.IsResolved = true; // ✅ FIX: Set to true - individual sleeve was deleted by clustering
                                        clashZone.SleeveInstanceId = -1; // Clear individual sleeve ID
                                        clashZone.SleeveFamilyName = string.Empty; // Clear individual sleeve family
                                        
                                        updated = true;
                                        markedCount++;
                                        
                                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Marked ClashZone {clashZone.Id} as cluster-resolved with cluster sleeve {clusterSleeveId.IntegerValue} (cleared individual flags)\n");
                                    }
                                }
                            }
                            
        // Save updated XML if changes were made
        if (updated)
        {
            try
            {
                using (var writer = new StreamWriter(xmlFile))
                {
                    serializer.Serialize(writer, filter);
                }
                
                // 🔥 CRITICAL DEBUG: Log XML file save
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 💾 XML FILE SAVED: {xmlFile} with {markedCount} cluster-resolved clash zones\n");
            }
            catch (Exception ex)
            {
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ ERROR SAVING XML FILE: {xmlFile} - {ex.Message}\n");
            }
        }
        else
        {
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] ⚠️ NO CHANGES MADE - XML FILE NOT SAVED: {xmlFile}\n");
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
                                // Check cluster sleeves
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
                                
                                // ⚠️ CRITICAL: Also check individual sleeves (IsResolved) - BUT ONLY if NOT cluster-resolved
                                // If cluster-resolved, keep individual flag as true even if individual sleeve is missing
                                if (clashZone.IsResolved && !clashZone.IsClusterResolved)
                                {
                                    if (clashZone.SleeveInstanceId <= 0)
                                    {
                                        clashZone.IsResolved = false;
                                        clashZone.ResolvedSleeveId = null;
                                        clashZone.SleeveInstanceId = -1;
                                        clashZone.SleeveFamilyName = string.Empty;
                                        clashZone.LastUpdated = DateTime.Now;
                                        resetCount++;
                                        modified = true;
                                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Reset individual sleeve flag for ClashZone {clashZone.Id} - SleeveInstanceId was {clashZone.SleeveInstanceId} (invalid)\n");
                                    }
                                    else
                                    {
                                        // Check if individual sleeve still exists in Revit
                                        var sleeveId = new ElementId(clashZone.SleeveInstanceId);
                                        var sleeve = doc.GetElement(sleeveId);
                                        if (sleeve == null)
                                        {
                                            // Individual sleeve was deleted - reset flag
                                            clashZone.IsResolved = false;
                                            clashZone.ResolvedSleeveId = null;
                                            clashZone.SleeveInstanceId = -1;
                                            clashZone.SleeveFamilyName = string.Empty;
                                            clashZone.LastUpdated = DateTime.Now;
                                            resetCount++;
                                            modified = true;
                                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"Reset individual sleeve flag for ClashZone {clashZone.Id} - sleeve {clashZone.SleeveInstanceId} was deleted\n");
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
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✓ Reset flags for {resetCount} deleted sleeves (cluster + individual)\n");
                }
                else
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", $"✓ No deleted sleeves found - all flags preserved\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error resetting cluster flags: {ex.Message}");
            }
        }

        /// <summary>
        /// Get clash zone by MEP element ID using clash cache (one time calculate, use many times)
        /// ✅ PERFORMANCE: Uses cache to avoid O(n·m) XML deserialization
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, string xmlFilePath = null)
        {
            try
            {
                // ✅ PERFORMANCE: Fast O(1) lookup using clash cache
                if (_clashZoneCache != null && _clashZoneCache.TryGetValue(mepElementId, out ClashZone clashZone))
                {
                    return clashZone;
                }

                // Fallback to XML search only if not in cache (should rarely happen)
                DebugLogger.Warning($"[UniversalClusterService] MEP Element ID {mepElementId} not found in cache, falling back to XML search");
                
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
                                foreach (var cz in filter.ClashZoneStorage.ClashZones)
                                {
                                    if (cz.MepElementId.IntegerValue == mepElementId)
                                    {
                                        return cz;
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

        /// <summary>
        /// Cache for clash zone data to avoid expensive lookups during clustering
        /// Key: MEP Element ID, Value: ClashZone data
        /// </summary>
        private Dictionary<long, ClashZone> _clashZoneCache = new Dictionary<long, ClashZone>();

        /// <summary>
        /// Load clash zone cache from XML files once at the start of clustering
        /// This follows "calculate once, use many times" principle
        /// </summary>
        private void LoadClashZoneCache(string xmlFilePath)
        {
            _clashZoneCache.Clear();
            
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                    return;

                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                foreach (var xmlFile in xmlFiles)
                {
                    if (File.Exists(xmlFile))
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var cz in filter.ClashZoneStorage.ClashZones)
                                {
                                    // Cache by MEP element ID for fast lookup
                                    if (cz.MepElementIdValue > 0)
                                    {
                                        _clashZoneCache[cz.MepElementIdValue] = cz;
                                    }
                                }
                            }
                        }
                    }
                }
                
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[CACHE] Loaded {_clashZoneCache.Count} clash zones into cache\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading clash zone cache: {ex.Message}");
            }
        }

        /// <summary>
        /// Get clash zone from cache by MEP element ID (from sleeve parameter)
        /// </summary>
        private ClashZone GetClashZoneFromCache(FamilyInstance sleeve)
        {
            try
            {
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepIdParam != null && !string.IsNullOrEmpty(mepIdParam.AsString()))
                {
                    if (long.TryParse(mepIdParam.AsString(), out long mepId))
                    {
                        if (_clashZoneCache.TryGetValue(mepId, out ClashZone cz))
                        {
                            return cz;
                        }
                    }
                }
            }
            catch { }
            return null;
        }
        
        /// <summary>
        /// Get pipe information string from cached clash zone
        /// </summary>
        private string GetPipeInfoFromCache(FamilyInstance sleeve)
        {
            var cz = GetClashZoneFromCache(sleeve);
            if (cz != null)
            {
                return $"pipeId={cz.MepElementIdValue}, pipe={cz.MepElementSize:F1}mm";
            }
            return "pipeId=Unknown, pipe=Unknown";
        }

        /// <summary>
        /// Get MEP_ElementId parameter value
        /// </summary>
        private string GetMepElementId(FamilyInstance sleeve)
        {
            try
            {
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepIdParam != null && !mepIdParam.IsReadOnly)
                {
                    return mepIdParam.AsString();
                }
            }
            catch { }
            return "Unknown";
        }

        /// <summary>
        /// Get pipe size from MEP_ElementId parameter
        /// </summary>
        private string GetPipeSizeFromMepElementId(FamilyInstance sleeve)
        {
            try
            {
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepIdParam != null && !mepIdParam.IsReadOnly)
                {
                    string mepElementId = mepIdParam.AsString();
                    if (long.TryParse(mepElementId, out long mepId))
                    {
                        var clashZone = GetClashZoneByMepElementId(mepId);
                        if (clashZone != null)
                        {
                            return $"{clashZone.MepElementSize:F1}mm";
                        }
                    }
                }
            }
            catch { }
            return "Unknown";
        }

        /// <summary>
        /// Calculate sleeve outer diameter from bounding box
        /// </summary>
        private double GetSleeveOuterDiameter(BoundingBoxXYZ bbox)
        {
            if (bbox == null) return 0.0;
            
            // For circular sleeves, diameter is the larger of width or height
            double width = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Millimeters);
            double height = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Millimeters);
            
            return Math.Max(width, height);
        }

        /// <summary>
        /// Check if a sleeve is circular based on its family name
        /// </summary>
        private bool IsCircularSleeve(FamilyInstance sleeve)
        {
            string familyName = sleeve.Symbol?.FamilyName ?? "";
            return familyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   familyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// Calculate the minimum distance between two circular sleeves using cached XML data
        /// Returns the clearance between the outer edges (center-to-center minus radii)
        /// Follows "calculate once, use many times" - uses MepElementSize from XML cache
        /// </summary>
        private double GetMinimumDistanceBetweenCircularSleeves(FamilyInstance sleeve1, FamilyInstance sleeve2, BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            // Calculate centers from bounding boxes
            XYZ center1 = new XYZ(
                (bbox1.Min.X + bbox1.Max.X) / 2.0,
                (bbox1.Min.Y + bbox1.Max.Y) / 2.0,
                (bbox1.Min.Z + bbox1.Max.Z) / 2.0
            );
            
            XYZ center2 = new XYZ(
                (bbox2.Min.X + bbox2.Max.X) / 2.0,
                (bbox2.Min.Y + bbox2.Max.Y) / 2.0,
                (bbox2.Min.Z + bbox2.Max.Z) / 2.0
            );
            
            // Calculate center-to-center distance (in Revit internal units - feet)
            double centerDistance = center1.DistanceTo(center2);
            
            // Get MEP element sizes from cached clash zones (already calculated during refresh)
            var cz1 = GetClashZoneFromCache(sleeve1);
            var cz2 = GetClashZoneFromCache(sleeve2);
            
            double diameter1_mm = cz1?.MepElementSize ?? GetSleeveOuterDiameter(bbox1);
            double diameter2_mm = cz2?.MepElementSize ?? GetSleeveOuterDiameter(bbox2);
            
            // Convert diameters to radii in feet
            double radius1 = UnitUtils.ConvertToInternalUnits(diameter1_mm / 2.0, UnitTypeId.Millimeters);
            double radius2 = UnitUtils.ConvertToInternalUnits(diameter2_mm / 2.0, UnitTypeId.Millimeters);
            
            // Clearance = center-to-center - (radius1 + radius2) (all in feet)
            double clearance = centerDistance - (radius1 + radius2);
            
            return Math.Max(0.0, clearance); // Return 0 if overlapping
        }

        /// <summary>
        /// Calculate the minimum distance between two bounding boxes
        /// Returns 0 if they overlap, otherwise the shortest distance between any two points
        /// </summary>
        private double GetMinimumDistanceBetweenBoundingBoxes(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            // Check if bounding boxes overlap
            bool xOverlap = bbox1.Max.X >= bbox2.Min.X && bbox1.Min.X <= bbox2.Max.X;
            bool yOverlap = bbox1.Max.Y >= bbox2.Min.Y && bbox1.Min.Y <= bbox2.Max.Y;
            bool zOverlap = bbox1.Max.Z >= bbox2.Min.Z && bbox1.Min.Z <= bbox2.Max.Z;
            
            if (xOverlap && yOverlap && zOverlap)
            {
                return 0.0; // Bounding boxes overlap
            }
            
            // Calculate minimum distance between non-overlapping bounding boxes
            double dx = Math.Max(0, Math.Max(bbox1.Min.X - bbox2.Max.X, bbox2.Min.X - bbox1.Max.X));
            double dy = Math.Max(0, Math.Max(bbox1.Min.Y - bbox2.Max.Y, bbox2.Min.Y - bbox1.Max.Y));
            double dz = Math.Max(0, Math.Max(bbox1.Min.Z - bbox2.Max.Z, bbox2.Min.Z - bbox1.Max.Z));
            
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
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

                // Calculate distance based on sleeve type
                // For circular sleeves: Use center-to-center minus radii (actual clearance)
                // For rectangular sleeves: Use bounding box boundary distance
                bool isCircular1 = IsCircularSleeve(inst);
                bool isCircular2 = IsCircularSleeve(s);
                
                double minDistance;
                string distanceMethod;
                
                if (isCircular1 && isCircular2)
                {
                    // Both circular: Use proper circular distance (center-to-center - radii)
                    // Uses MepElementSize from XML cache (calculate once, use many times)
                    minDistance = GetMinimumDistanceBetweenCircularSleeves(inst, s, o1_bbox, o2_bbox);
                    distanceMethod = "circular";
                }
                else
                {
                    // At least one rectangular: Use bounding box distance
                    minDistance = GetMinimumDistanceBetweenBoundingBoxes(o1_bbox, o2_bbox);
                    distanceMethod = "bbox";
                }
                
                // Debug logging for clustering distance with pipe sizes
                double minDistanceMm = UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters);
                double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
                
                // Get pipe information from cache (fast lookup from XML)
                string pipe1Info = GetPipeInfoFromCache(inst);
                string pipe2Info = GetPipeInfoFromCache(s);
                
                // Calculate individual sleeve outer diameter
                double sleeve1OD = GetSleeveOuterDiameter(o1_bbox);
                double sleeve2OD = GetSleeveOuterDiameter(o2_bbox);
                
                if (minDistance <= toleranceDist)
                {
                    // ✅ PERFORMANCE FIX: No logging during clustering - only log summary at end
                    neighbors.Add(s);
                }
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
            // ⚠️ FIX: X-orientation walls need rotation, not Y-orientation walls
            double rotationAngle = 0.0;
            if (groupKey.hostType == "Wall" && groupKey.orientation == "X")
            {
                rotationAngle = Math.PI / 2;  // 90° rotation for X-oriented walls
            }

            if (rotationAngle != 0.0)
            {
                XYZ axisOrigin = mid;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
            }

            // Set size parameters (swap dimensions if rotated for orientation alignment)
            bool shouldSwapDimensions = (groupKey.hostType == "Wall" && rotationAngle != 0.0);
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);

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
            double depth,
            bool shouldSwapDimensions = false)
        {
            var widthParam = inst.LookupParameter("Width");
            var heightParam = inst.LookupParameter("Height");
            var depthParam = inst.LookupParameter("Depth");

            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                // Step 1: Assign all from bbox
                double openingWidth = width;
                double openingHeight = height;
                double openingDepth = depth;
                
                // Step 2: Map world coordinates to sleeve parameters
                if (groupKey.orientation == "Y")
                {
                    // Y-wall: Width=H, Height=D, Depth=W
                    double tempWidth = openingWidth;
                    double tempHeight = openingHeight;
                    double tempDepth = openingDepth;
                    openingWidth = tempHeight;  // H (608.6mm)
                    openingHeight = tempDepth;  // D (270mm)
                    openingDepth = tempWidth;   // W (200mm)
                }
                else if (shouldSwapDimensions) // X-walls
                {
                    // X-wall: Width=W, Height=D, Depth=H
                    double tempWidth = openingWidth;
                    double tempHeight = openingHeight;
                    double tempDepth = openingDepth;
                    openingWidth = tempWidth;   // W (478.8mm)
                    openingHeight = tempDepth;  // D (238.9mm)
                    openingDepth = tempHeight;  // H (200mm)
                }
                
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[CLUSTER-DIM] Sleeve {inst.Id}: Orientation={groupKey.orientation}, Swap={shouldSwapDimensions}, " +
                    $"BBox(W={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, " +
                    $"H={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, " +
                    $"D={UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm), " +
                    $"Final(W={UnitUtils.ConvertFromInternalUnits(openingWidth, UnitTypeId.Millimeters):F1}mm, " +
                    $"H={UnitUtils.ConvertFromInternalUnits(openingHeight, UnitTypeId.Millimeters):F1}mm, " +
                    $"D={UnitUtils.ConvertFromInternalUnits(openingDepth, UnitTypeId.Millimeters):F1}mm)\n");
                
                // Set Width and Height from bbox dimensions (after swap if needed)
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(openingWidth);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(openingHeight);

                // Get actual host thickness for Depth parameter (through-wall dimension)
                // ⚠️ CRITICAL: Use ClashZone.StructuralElementThickness (pre-calculated during refresh)
                var firstSleeve = cluster[0];
                var mepIdParam = firstSleeve.LookupParameter("MEP_ElementId");
                double hostThickness = openingDepth;  // Default fallback
                
                if (mepIdParam != null)
                {
                    long mepId = mepIdParam.AsInteger();
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                        $"[CLUSTER-DEPTH] Sleeve {inst.Id}: Looking up ClashZone for MEP_ElementId={mepId}\n");
                    
                    var clashZone = GetClashZoneByMepElementId(mepId);
                    if (clashZone != null)
                    {
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                            $"[CLUSTER-DEPTH] Found ClashZone {clashZone.Id}, StructThickness={clashZone.StructuralElementThickness:F6}ft ({UnitUtils.ConvertFromInternalUnits(clashZone.StructuralElementThickness, UnitTypeId.Millimeters):F1}mm)\n");
                        
                        if (clashZone.StructuralElementThickness > 0)
                        {
                            hostThickness = clashZone.StructuralElementThickness;
                        }
                    }
                    else
                    {
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                            $"[CLUSTER-DEPTH] ClashZone NOT FOUND for MEP_ElementId={mepId}\n");
                    }
                }
                
                // Fallback: Try to get from host element directly if ClashZone thickness was not available
                if (hostThickness <= 0.0)
                {
                    if (groupKey.hostType == "Wall")
                    {
                        var wall = cluster[0].Host as Wall;
                        if (wall != null)
                        {
                            hostThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                            DebugLogger.Log($"[ClusterService] Wall thickness from host element: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
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
                            else
                            {
                                // ⚠️ BUG FIX: Use bounding box calculation instead of hardcoded 500mm
                                var framingBbox = framing.get_BoundingBox(null);
                                if (framingBbox != null)
                                {
                                    // Calculate thickness from bounding box
                                    var thickness = Math.Max(
                                        Math.Max(
                                            framingBbox.Max.X - framingBbox.Min.X,
                                            framingBbox.Max.Y - framingBbox.Min.Y
                                        ),
                                        framingBbox.Max.Z - framingBbox.Min.Z
                                    );
                                    hostThickness = thickness;
                                    DebugLogger.Log($"[ClusterService] FRAMING FIX: Using calculated thickness {UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters):F1}mm instead of 500mm fallback");
                                }
                                else
                                {
                                    hostThickness = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters);
                                    DebugLogger.Log($"[ClusterService] FRAMING FALLBACK: Using 500mm fallback (no bounding box available)");
                                }
                            }
                        }
                    }
                }
                
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[CLUSTER-DIM] Sleeve {inst.Id}: Mapped Depth = {UnitUtils.ConvertFromInternalUnits(openingDepth, UnitTypeId.Millimeters):F1}mm, HostThickness = {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm\n");
                    
                // Set the mapped dimensions (use hostThickness for depth instead of openingDepth)
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(openingWidth);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(openingHeight);
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

        /// <summary>
        /// ✅ CRITICAL: Update ClashZone flags after placing cluster sleeve
        /// This implements the flag management system to prevent individual sleeves over cluster sleeves
        /// </summary>
        private void UpdateClashZoneFlagsForCluster(
            FamilyInstance clusterInstance, 
            List<FamilyInstance> originalSleeves, 
            string systemType)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterService] 🔥 UpdateClashZoneFlagsForCluster CALLED 🔥");
                DebugLogger.Info($"[UniversalClusterService] Cluster ID: {clusterInstance.Id.IntegerValue}, Original sleeves: {originalSleeves.Count}, SystemType: {systemType}");
                
                // ✅ DEBUG: Log the category mapping
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_management_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] UpdateClashZoneFlagsForCluster: systemType='{systemType}', clusterId={clusterInstance.Id.IntegerValue}\n");

                // Map systemType to category (Ducts → ducts, Pipes → pipes, etc.)
                // systemType comes from GetCategoryFromMepElementId which returns clashZone.MepElementCategory (plural)
                string category = systemType switch
                {
                    "Ducts" => "ducts",
                    "Pipes" => "pipes", 
                    "Cable Trays" => "cabletrays",
                    "Duct Accessories" => "duct_accessories",
                    _ => "ducts"
                };

                // Get cluster mark
                string clusterMark = clusterInstance.LookupParameter("Mark")?.AsString() 
                                  ?? $"CO-{clusterInstance.Id.IntegerValue}";

                DebugLogger.Info($"[UniversalClusterService] Loading clash zones for category: {category}");
                
                // ✅ DEBUG: Log the category mapping result
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_management_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] Category mapping: '{systemType}' → '{category}'\n");

                // Load clash zones from XML
                var clashZones = LoadClashZonesFromXml(category);
                DebugLogger.Info($"[UniversalClusterService] Loaded {clashZones.Count} clash zones from XML");

                // Update each clash zone for deleted sleeves
                int updatedCount = 0;
                foreach (var sleeve in originalSleeves)
                {
                    var clashZone = clashZones.FirstOrDefault(cz => 
                        cz.SleeveInstanceId == sleeve.Id.IntegerValue);
                    
                    if (clashZone != null)
                    {
                        // Mark as clustered
                        clashZone.IsClustered = true;
                        clashZone.IsResolved = true;
                        clashZone.ClusterSleeveInstanceId = clusterInstance.Id.IntegerValue;
                        clashZone.SleeveInstanceId = -1;  // Deleted
                        clashZone.SleeveFamilyName = string.Empty;
                        
                        updatedCount++;
                        DebugLogger.Info($"[UniversalClusterService] ✓ Updated ClashZone {clashZone.Id} as clustered");
                    }
                    else
                    {
                        DebugLogger.Warning($"[UniversalClusterService] ⚠️ ClashZone not found for sleeve {sleeve.Id.IntegerValue}");
                    }
                }

                // Save updated clash zones back to XML
                SaveClashZonesToXml(clashZones, category);
                DebugLogger.Info($"[UniversalClusterService] ✅ Updated {updatedCount} clash zones as clustered, saved to XML");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error updating clash zone flags: {ex.Message}");
            }
        }

        /// <summary>
        /// Load clash zones from category-specific XML file
        /// </summary>
        private List<Models.ClashZone> LoadClashZonesFromXml(string category)
        {
            try
            {
                string filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");

                // Find pattern: *_{category}.xml (e.g., Ventilation_ducts.xml)
                var pattern = $"*_{category}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                
                if (matchingFiles.Length == 0)
                {
                    DebugLogger.Warning($"[UniversalClusterService] No XML files found for pattern: {pattern}");
                    return new List<Models.ClashZone>();
                }

                // Use most recently modified file
                var xmlFile = matchingFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();
                DebugLogger.Info($"[UniversalClusterService] Loading clash zones from: {xmlFile}");

                var serializer = new XmlSerializer(typeof(Models.OpeningFilter));
                using (var reader = new StreamReader(xmlFile))
                {
                    var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                    return filter.ClashZoneStorage?.ClashZones ?? new List<Models.ClashZone>();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading clash zones from XML: {ex.Message}");
                return new List<Models.ClashZone>();
            }
        }

        /// <summary>
        /// Find the newest cluster sleeve that was just placed
        /// </summary>
        private FamilyInstance FindNewestClusterSleeve(Document doc, List<FamilyInstance> originalSleeves, SleeveGroupKey groupKey)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterService] 🔥 FindNewestClusterSleeve CALLED 🔥");
                DebugLogger.Info($"[UniversalClusterService] Host Type: {groupKey.hostType}, System Type: {groupKey.systemType}");

                // ✅ CRITICAL FIX: Use the same logic as PlaceClusterSleeve to determine family name
                // Determine if cluster is circular or rectangular (same logic as PlaceClusterSleeve)
                bool isCircular = originalSleeves.All(s => 
                {
                    var fam = s.Symbol?.Family?.Name ?? "";
                    return fam.Contains("Circular");
                });
                
                // PIPE CLUSTERS: Always rectangular regardless of member shape (legacy parity)
                bool isPipeCategory = (!string.IsNullOrEmpty(groupKey.systemType) && groupKey.systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0);
                if (isPipeCategory)
                {
                    isCircular = false;
                    DebugLogger.Info($"[UniversalClusterService] For Pipes category: forcing rectangular cluster shape");
                }

                // Select universal family based on host type and shape (same logic as PlaceClusterSleeve)
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
                    DebugLogger.Error($"[UniversalClusterService] Unknown host type: {groupKey.hostType}");
                    return null;
                }

                DebugLogger.Info($"[UniversalClusterService] Looking for cluster family: {familyName} (shape: {(isCircular ? "Circular" : "Rectangular")})");

                // Find all family instances of the cluster family type
                var clusterSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(fi => fi.Id.IntegerValue) // Newest will have highest ID
                    .ToList();

                DebugLogger.Info($"[UniversalClusterService] Found {clusterSleeves.Count} cluster sleeves of type {familyName}");

                // Return the newest one (highest ID)
                var newestSleeve = clusterSleeves.FirstOrDefault();
                if (newestSleeve != null)
                {
                    DebugLogger.Info($"[UniversalClusterService] ✅ Found newest cluster sleeve: ID {newestSleeve.Id.IntegerValue}");
                }
                else
                {
                    DebugLogger.Warning($"[UniversalClusterService] ⚠️ No cluster sleeve found for family {familyName}");
                }

                return newestSleeve;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error finding newest cluster sleeve: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Save updated clash zones back to XML file
        /// </summary>
        private void SaveClashZonesToXml(List<Models.ClashZone> clashZones, string category)
        {
            try
            {
                string filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");

                // Find pattern: *_{category}.xml (e.g., Ventilation_ducts.xml)
                var pattern = $"*_{category}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                
                if (matchingFiles.Length == 0)
                {
                    DebugLogger.Warning($"[UniversalClusterService] No XML files found for pattern: {pattern}");
                    return;
                }

                // Use most recently modified file
                var xmlFile = matchingFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();
                DebugLogger.Info($"[UniversalClusterService] Saving clash zones to: {xmlFile}");

                var serializer = new XmlSerializer(typeof(Models.OpeningFilter));
                Models.OpeningFilter filter;
                
                // Load existing filter
                using (var reader = new StreamReader(xmlFile))
                {
                    filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                }
                
                // Update clash zone storage
                if (filter.ClashZoneStorage == null)
                {
                    filter.ClashZoneStorage = new Models.ClashZoneStorage();
                }
                filter.ClashZoneStorage.ClashZones = clashZones;
                filter.LastModified = DateTime.Now;
                
                // Save back
                using (var writer = new StreamWriter(xmlFile))
                {
                    serializer.Serialize(writer, filter);
                }
                
                DebugLogger.Info($"[UniversalClusterService] ✅ Successfully saved {clashZones.Count} clash zones to XML");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error saving clash zones to XML: {ex.Message}");
            }
        }
    }
}

