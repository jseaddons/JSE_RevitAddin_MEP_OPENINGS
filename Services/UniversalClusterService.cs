using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using static JSE_RevitAddin_MEP_OPENINGS.Models.MepCategoryConstants;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Universal clustering service - extracted from RectangularSleeveClusterCommandV2
    /// Can be called from any context (ICommand, IExternalCommand, etc.)
    /// </summary>
    public class UniversalClusterService
    {
        private string _filterName; // Store filter name for use in GetFilterNameForCategory
        private Document _doc; // Active document context for path resolution
        
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
        public (int placedCount, int deletedCount) ClusterSleeves(Document doc, string targetCategory, UIDocument uiDoc = null, string xmlFilePath = null, string filterName = null, List<FamilyInstance> placedClusterSleevesOut = null)
        {
            // Store document for helper methods
            _doc = doc;
            // ✅ CRITICAL: Store filterName in class field for use in GetFilterNameForCategory
            _filterName = filterName;
            
            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>(); // Track placed cluster sleeves for cleanup
            
            // ✅ PERFORMANCE: Start timing
            var startTime = DateTime.Now;

            try
            {
                // ✅ PERFORMANCE FIX: Minimal logging - only log session start and end
                string clusterLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log";
                bool __suppressClusterLogs = DeploymentConfiguration.DeploymentMode;
                if (!DeploymentConfiguration.DeploymentMode)
                    File.AppendAllText(clusterLogPath, $"\n===== CLUSTER SESSION STARTED {DateTime.Now:HH:mm:ss} =====\n");
                
                // 🔥 CRITICAL DEBUG: Log which XML file cluster service is working with
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 CLUSTER SERVICE WORKING WITH XML FILE: {xmlFilePath ?? "NULL"}\n");
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 CLUSTER SERVICE FILTER NAME: {filterName ?? "NULL"}\n");
                
                // ⚠️ CRITICAL: Reset cluster flags for deleted cluster sleeves
                ResetClusterFlagsForDeletedSleeves(doc, xmlFilePath);
                
                // ✅ CRITICAL: Small delay to ensure XML files are fully written to disk
                System.Threading.Thread.Sleep(100);
                
                // ✅ DYNAMIC: Load clash zone cache from regular XML files (not _CLUSTER.xml)
                LoadClashZoneCacheFromRegularXml(xmlFilePath, targetCategory, doc, filterName);
                
                // ✅ NEW: Use cheaper bounding box overlap algorithm instead of expensive PreCalculatedClusterService
                // This uses saved bounding box coordinates from sleeve placement (no Revit API calls needed)
                
                // Load user settings first
                var settingsService = new SettingsService();
                var currentProfile = ApplicationProfileService.Instance.GetCurrentProfile();
                if (currentProfile != null)
                {
                    var settings = settingsService.GetSettings(currentProfile);
                    ClusterConfigurationManager.Instance.LoadFromSettings(settings, "User Settings");
                }
                
                double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
                var toleranceDist = UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
                DebugLogger.Info($"[UniversalClusterService] Using cheaper bounding box overlap algorithm with tolerance: {toleranceMm}mm");
                
                // ✅ DYNAMIC: Reload cache with updated cluster flags from regular XML
                LoadClashZoneCacheFromRegularXml(xmlFilePath, targetCategory, doc, filterName);
                
                // Get cluster configuration
                if (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    toleranceMm = Math.Min(toleranceMm, 100); // clamp pipes to 100mm
                }
                toleranceDist = UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
                
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
                
                // ✅ FIX: Read sleeves directly from Revit instead of relying on clash zone cache
                // This ensures clustering works even when no clash zones exist (like for pipes)
                var cacheSleeves = allSleeves.Where(s => 
                {
                    var mepElementIdParam = s.LookupParameter("MEP_ElementId");
                    if (mepElementIdParam != null)
                    {
                        long mepElementId = mepElementIdParam.AsInteger();
                        
                        // ✅ CRITICAL FIX: Process ALL sleeves that have MEP_ElementId parameter
                        // Don't filter by clash zone cache - read directly from Revit
                        DebugLogger.Log($"[UniversalClusterService] PROCESS: Sleeve {s.Id} (MEP {mepElementId}) found in Revit - proceeding with clustering");
                            return true;
                    }
                    return false;
                }).ToList();
                
                // ✅ DEBUG: Log current vs old clash filtering statistics
                int currentClashCount = 0;
                int oldClashCount = 0;
                if (_clashZoneCache != null)
                {
                    foreach (var clashZone in _clashZoneCache.Values)
                    {
                        if (clashZone.IsCurrentClash)
                            currentClashCount++;
                        else
                            oldClashCount++;
                    }
                }
                
                DebugLogger.Log($"[UniversalClusterService] Found {cacheSleeves.Count} sleeves in Revit model (out of {allSleeves.Count} total)");
                DebugLogger.Log($"[UniversalClusterService] DIRECT REVIT READING: Processing all sleeves with MEP_ElementId parameter");
                File.AppendAllText(clusterLogPath, $"Found {cacheSleeves.Count} sleeves in Revit model (out of {allSleeves.Count} total)\n");
                File.AppendAllText(clusterLogPath, $"DIRECT REVIT READING: Processing all sleeves with MEP_ElementId parameter\n");
                
                // ✅ CRITICAL: Filter by MEP_Category parameter to prevent cross-category clustering
                // ✅ FIXED: Use XML data directly for clustering - no Revit API calls needed
                // Get sleeves from XML data only (SleeveInstanceId > 0 means placed sleeves)
                var rawSleeves = _clashZoneCache.Values
                    .Where(cz => cz.SleeveInstanceId > 0) // Only placed sleeves
                    .Where(cz => !cz.IsClusterResolved) // ✅ CRITICAL: Skip sleeves already resolved by cluster
                    .Where(cz => string.IsNullOrEmpty(targetCategory) || 
                                string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    .Select(cz => new { 
                        SleeveInstanceId = cz.SleeveInstanceId,
                        Category = cz.MepElementCategory,
                        HostType = GetHostTypeFromClashZone(cz),
                        Orientation = GetEffectiveOrientationForClustering(cz),
                        BoundingBox = GetBoundingBoxFromClashZone(cz)
                    })
                    .ToList();
                
                DebugLogger.Log($"[UniversalClusterService] Found {rawSleeves.Count} sleeves from XML data for category '{targetCategory}'");
                
                DebugLogger.Log($"[UniversalClusterService] Processing {rawSleeves.Count} sleeves" + 
                               (string.IsNullOrEmpty(targetCategory) ? " (all categories)" : $" (category-filtered for {targetCategory})"));
                File.AppendAllText(clusterLogPath, $"Processing {rawSleeves.Count} sleeves (category-filtered for {targetCategory ?? "ALL"})\n");
                
                // ⚠️ CRITICAL: Log flag states of sleeves being processed for clustering (AFTER filtering)
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_state_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] Processing {rawSleeves.Count} sleeves for clustering (filtered for {targetCategory ?? "ALL"})\n");
                
                foreach (var sleeve in rawSleeves.Take(5)) // Log first 5 filtered sleeves
                {
                    // ✅ FIXED: Use XML data instead of Revit API calls
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    var clashZone = GetClashZoneBySleeveInstanceId(sleeveInstanceId);
                    if (clashZone != null)
                    {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] Sleeve {sleeveInstanceId}: MEP={clashZone.MepElementIdValue}, ClashZone={clashZone.Id}\n");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                    }
                }

                // ✅ FIXED: Work with XML data directly - no Revit API calls needed
                if (rawSleeves.Count == 0)
                {
                    DebugLogger.Log($"[UniversalClusterService] No sleeves to cluster from XML data");
                    return (0, 0);
                }

                // ⚠️ CRITICAL: Transaction must be started by caller
                if (!doc.IsModifiable)
                {
                    DebugLogger.Error($"[UniversalClusterService] Document is not modifiable - transaction must be started by caller");
                    throw new InvalidOperationException("Document must be in a transaction before calling ClusterSleeves");
                }

                // Group sleeves by host type, system type, and orientation using XML data
                var sleeveGroups = rawSleeves.GroupBy(sleeve => {
                    // All data comes from XML - no Revit API calls needed
                    string hostType = sleeve.HostType;
                    string systemType = sleeve.Category;
                    string effectiveOrientation = sleeve.Orientation;
                    
                    // Log to dedicated cluster debug file
                    File.AppendAllText(clusterLogPath, $"Sleeve {sleeve.SleeveInstanceId}: hostType={hostType}, systemType={systemType}, orientation={effectiveOrientation}\n");
                    
                    return new SleeveGroupKey(hostType, systemType, effectiveOrientation);
                });

                // Log group diagnostics
                foreach (var g in sleeveGroups)
                {
                    var list = g.ToList();
                    var ids = list.Select(s => s.SleeveInstanceId.ToString()).Take(10).ToList();
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
                                DebugLogger.Info($"[DEBUG] SKIP: Cluster already exists for {cluster.Count} sleeves - skipping placement\n");
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
                                    UpdateClashZoneFlagsForCluster(clusterSleeve, cluster, groupKey.systemType, doc);
                                    
                                    // ✅ Track placed cluster sleeve for cleanup
                                    placedClusters.Add(clusterSleeve);
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
                DebugLogger.Info($"[DEBUG] Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.\n");
                
                // ⚠️ DISABLED: Cleanup will be called AFTER XML save in orchestrator to use cached bounding boxes
                // Cleanup requires cluster sleeve bounding boxes to be saved to XML first
                // The orchestrator will call CleanupSleevesWithinClustersAfterXmlSave() after XML update
                
                DebugLogger.Log($"[UniversalClusterService] Final Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                DebugLogger.Info($"[DEBUG] Final Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.\n");
                
                // ✅ RETURN: Populate the output list with placed cluster sleeves
                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.Clear();
                    placedClusterSleevesOut.AddRange(placedClusters);
                    DebugLogger.Log($"[UniversalClusterService] Returning {placedClusters.Count} placed cluster sleeves for coordinate update");
                }
                
                // ✅ PERFORMANCE: Log clustering performance
                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                DebugLogger.Info($"[UniversalClusterService] ⚡ Clustering completed in {duration.TotalSeconds:F1} seconds");
                DebugLogger.Info($"[DEBUG] ⚡ PERFORMANCE: Clustering completed in {duration.TotalSeconds:F1} seconds\n");
                
                DebugLogger.Info($"[DEBUG] ===== CLUSTER DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Clustering failed: {ex.Message}");
                DebugLogger.Error($"[UniversalClusterService] Stack trace: {ex.StackTrace}");
                DebugLogger.Info($"[DEBUG] ERROR: {ex.Message}\n");
                DebugLogger.Info($"[DEBUG] Stack trace: {ex.StackTrace}\n");
                throw;
            }

            return (placedCount, deletedCount);
        }

        /// <summary>
        /// ✅ FIX: Get actual HostType from SleeveData XML instead of hardcoding from family name
        /// </summary>
        private string GetHostTypeFromSleeveData(FamilyInstance sleeve)
        {
            try
            {
                // ✅ DYNAMIC: Get host type directly from sleeve instance using UniversalSleevePlacerService logic
                var host = sleeve.Host;
                if (host != null)
                {
                    var hostType = host.GetType().Name;
                    if (hostType.Contains("Wall")) return "Wall";
                    if (hostType.Contains("Floor")) return "Floor";
                    if (hostType.Contains("Ceiling")) return "Ceiling";
                    if (hostType.Contains("StructuralFraming")) return "Structural Framing";
                    return hostType;
                }
                
                // ✅ DYNAMIC: Fallback to HostOrientation parameter if available
                var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
                if (hostOrientationParam != null && !string.IsNullOrEmpty(hostOrientationParam.AsString()))
                {
                    string orientation = hostOrientationParam.AsString();
                    if (orientation == "FloorHosted") return "Floor";
                    if (orientation == "WallHosted") return "Wall";
                    if (orientation == "X" || orientation == "Y") return "Wall"; // X/Y typically means wall
                }
                
                // ✅ DYNAMIC: Fallback to family name analysis
                var familyName = sleeve.Symbol?.Family?.Name ?? "";
                if (familyName.Contains("Wall")) return "Wall";
                if (familyName.Contains("Slab")) return "Floor";
                
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine host type dynamically");
                return "Unknown";
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting HostType for sleeve {sleeve.Id.IntegerValue}: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// ✅ FIX: Get orientation from SleeveData XML instead of ClashZone cache
        /// </summary>
        private string GetOrientationFromClashZone(FamilyInstance sleeve)
        {
            try
            {
                // ✅ DYNAMIC: Get orientation directly from MEP element using UniversalSleevePlacerService logic
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null && mepElementIdParam.HasValue)
                {
                    var mepElementId = mepElementIdParam.AsElementId();
                    var mepElement = sleeve.Document.GetElement(mepElementId);
                    if (mepElement != null)
                    {
                        // Get orientation from MEP element's Wall Direction Type
                        var wallDirectionParam = mepElement.LookupParameter("Wall Direction Type");
                        if (wallDirectionParam != null && !string.IsNullOrEmpty(wallDirectionParam.AsString()))
                        {
                            string orientation = wallDirectionParam.AsString();
                            DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Orientation='{orientation}' from MEP element Wall Direction Type");
                            return orientation;
                        }
                    }
                }
                
                // ✅ DYNAMIC: Fallback to HostOrientation parameter
                var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
                if (hostOrientationParam != null && !string.IsNullOrEmpty(hostOrientationParam.AsString()))
                {
                    string orientation = hostOrientationParam.AsString();
                    DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Orientation='{orientation}' from HostOrientation parameter");
                    return orientation;
                }
                
                // ✅ DYNAMIC: Fallback to geometric analysis
                var bbox = sleeve.get_BoundingBox(null);
                if (bbox != null)
                {
                    double width = bbox.Max.X - bbox.Min.X;
                    double height = bbox.Max.Y - bbox.Min.Y;
                    double depth = bbox.Max.Z - bbox.Min.Z;
                    
                    // Determine orientation based on largest dimension
                    if (width > height && width > depth) return "X";
                    if (height > width && height > depth) return "Y";
                    if (depth > width && depth > height) return "Z";
                }
                
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine orientation dynamically");
                return "";
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting orientation from MEP element: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// ✅ FIX: Get category from SleeveData XML instead of ClashZone cache
        /// </summary>
        private string GetCategoryFromMepElementId(FamilyInstance sleeve)
        {
            try
            {
                // ✅ DYNAMIC: Get category directly from MEP_Category parameter (set by UniversalSleevePlacerService)
                var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                if (mepCategoryParam != null && !string.IsNullOrEmpty(mepCategoryParam.AsString()))
                {
                    string category = mepCategoryParam.AsString();
                    DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Category='{category}' from MEP_Category parameter");
                    return category;
                }
                
                // ✅ DYNAMIC: Fallback to MEP element lookup using UniversalSleevePlacerService logic
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null && mepElementIdParam.HasValue)
                {
                    var mepElementId = mepElementIdParam.AsElementId();
                    var mepElement = sleeve.Document.GetElement(mepElementId);
                    if (mepElement != null)
                    {
                        var category = mepElement.Category?.Name;
                        if (!string.IsNullOrEmpty(category))
                        {
                            // Map Revit categories to our system categories
                            if (category.Contains("Pipe")) return "Pipes";
                            if (category.Contains("Duct")) return "Ducts";
                            if (category.Contains("Cable")) return "Cable Trays";
                            if (category.Contains("Duct Accessory")) return "Duct Accessories";
                            
                            DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Category='{category}' from MEP element");
                            return category;
                        }
                    }
                }
                
                // ✅ DYNAMIC: Fallback to family name analysis
                var familyName = sleeve.Symbol?.Family?.Name ?? "";
                if (familyName.Contains("Circular")) return "Pipes"; // Circular openings are typically pipes
                if (familyName.Contains("Rectangular")) return "Ducts"; // Rectangular openings are typically ducts
                
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine category dynamically");
                return "Unknown";
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting category from MEP element: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// Check if cluster already exists using flag management (efficient)
        /// </summary>
        private bool IsClusterAlreadyExists(List<dynamic> cluster, SleeveGroupKey groupKey)
        {
            try
            {
                // Check if any sleeve in the cluster is already cluster-resolved
                foreach (var sleeve in cluster)
                {
                    // ✅ FIXED: Use XML data instead of Revit API calls
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    
                    // Find clash zone for this sleeve instance ID
                    var clashZone = GetClashZoneBySleeveInstanceId(sleeveInstanceId);
                    if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                    {
                        DebugLogger.Info($"[DEBUG] Cluster already exists: ClashZone {clashZone.Id} is cluster-resolved (cluster sleeve ID: {clashZone.ClusterSleeveInstanceId})\n");
                        return true;
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
        private void MarkClashZonesAsClusterResolvedWithSleeveId(List<dynamic> cluster, ElementId clusterSleeveId, string xmlFilePath = null)
        {
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                    return;

                // ✅ FIX: Only process specific XML file if provided (ONE SOURCE OF TRUTH)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")  // Backward compatibility
                    : new[] { xmlFilePath };  // Only the specific file
                    
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
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
                                // ✅ FIXED: Use sleeve instance ID from XML data
                                int sleeveInstanceId = sleeve.SleeveInstanceId;
                                
                                // Find and mark clash zone as cluster-resolved using sleeve instance ID
                                var clashZone = filter.ClashZoneStorage.ClashZones.FirstOrDefault(cz => cz.SleeveInstanceId == sleeveInstanceId);
                                if (clashZone != null)
                                {
                                    clashZone.IsClusterResolved = true;
                                    clashZone.ClusterSleeveId = clusterSleeveId;
                                    clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue; // ✅ FIX: Store integer for XML serialization
                                    clashZone.LastUpdated = DateTime.Now;
                                    
                                    // ✅ CORRECT: Individual sleeve was placed then deleted during clustering
                                    // Set clustering history flag to distinguish from non-proximity sleeves
                                    clashZone.MarkedForClusteringSleeveProcess = true; // ✅ FIX: Mark as part of clustering history
                                    clashZone.IsResolved = true; // Individual sleeve was placed (then deleted)
                                    
                                    // ✅ STORAGE: Save original SleeveInstanceId BEFORE clearing it
                                    clashZone.AfterClusterSleevePlacedSleeveInstanceId = clashZone.SleeveInstanceId;
                                    
                                    clashZone.SleeveInstanceId = -1; // Individual sleeve was deleted
                                    clashZone.SleeveFamilyName = string.Empty; // Individual sleeve family cleared
                                    
                                    // ✅ GLOBAL XML: Record cluster placement in global XML
                                    try
                                    {
                                        var categoryName = clashZone.MepElementCategory;
                                        // ✅ MEMORY OPTIMIZATION: Use singleton to avoid reloading XML multiple times
                                        var globalManager = GlobalFlagManager.GetOrCreate(categoryName);
                                        
                                        // Get filter filename from XML file path
                                        string filterName = Path.GetFileName(xmlFile ?? "unknown_filter.xml");
                                        
                                        globalManager.RecordPlacement(
                                            clashZone.MepElementId,
                                            clashZone.StructuralElementId,
                                            null, // Individual sleeve was deleted
                                            clusterSleeveId, // Cluster sleeve ID
                                            filterName
                                        );
                                        
                                        DebugLogger.Info($"[GLOBAL-XML] Recorded cluster sleeve {clusterSleeveId.IntegerValue} for MEP={clashZone.MepElementId.IntegerValue}, Host={clashZone.StructuralElementId.IntegerValue}");
                                    }
                                    catch (Exception globalEx)
                                    {
                                        DebugLogger.Warning($"[GLOBAL-XML] Error recording cluster placement: {globalEx.Message}");
                                    }
                                    
                                    // ⚠️ CRITICAL: Log flag state AFTER cluster sleeve placement
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] ClashZone {clashZone.Id}: Cluster sleeve {clusterSleeveId.IntegerValue} placed\n");
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                                    
                                    updated = true;
                                    markedCount++;
                                    
                                    DebugLogger.Info($"[DEBUG] Marked ClashZone {clashZone.Id} as cluster-resolved with cluster sleeve {clusterSleeveId.IntegerValue} (cleared individual flags)\n");
                                }
                                else
                                {
                                    DebugLogger.Info($"[DEBUG] ✗ Clash zone not found for SleeveInstanceId={sleeveInstanceId}\n");
                                }
                            }
                            
                            // Save updated XML if changes were made
                            if (updated)
                            {
                                // ⚠️ CRITICAL: Log flag states BEFORE XML save after clustering
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-XML-SAVE-BEFORE] About to save XML after clustering\n");
                                
                                try
                                {
                                    // Normalize coordinates before saving to avoid 0,0,0 in XML
                                    if (filter?.ClashZoneStorage?.ClashZones != null)
                                    {
                                        foreach (var z in filter.ClashZoneStorage.ClashZones)
                                        {
                                            if (z == null) continue;
                                            bool hasIP = z.IntersectionPoint != null;
                                            bool isIPZero = hasIP && Math.Abs(z.IntersectionPoint.X) < 1e-9 && Math.Abs(z.IntersectionPoint.Y) < 1e-9 && Math.Abs(z.IntersectionPoint.Z) < 1e-9;
                                            bool hasSPP = z.SleevePlacementPoint != null;
                                            bool isSPPZero = hasSPP && Math.Abs(z.SleevePlacementPoint.X) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Y) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Z) < 1e-9;
                                            
                                            if (hasIP && !isIPZero)
                                            {
                                                z.IntersectionPointX = z.IntersectionPoint.X;
                                                z.IntersectionPointY = z.IntersectionPoint.Y;
                                                z.IntersectionPointZ = z.IntersectionPoint.Z;
                                            }
                                            else if (hasSPP && !isSPPZero)
                                            {
                                                z.IntersectionPointX = z.SleevePlacementPoint.X;
                                                z.IntersectionPointY = z.SleevePlacementPoint.Y;
                                                z.IntersectionPointZ = z.SleevePlacementPoint.Z;
                                            }
                                            else if (z.ClashBoundingBox != null)
                                            {
                                                var center = (z.ClashBoundingBox.Min + z.ClashBoundingBox.Max) / 2.0;
                                                z.IntersectionPointX = center.X;
                                                z.IntersectionPointY = center.Y;
                                                z.IntersectionPointZ = center.Z;
                                            }
                                        }
                                    }
                                    using (var writer = new StreamWriter(xmlFile))
                                    {
                                        serializer.Serialize(writer, filter);
                                    }
                                    
                                    // ⚠️ CRITICAL: Log flag states AFTER XML save after clustering
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-XML-SAVE-AFTER] XML save completed after clustering\n");
                                    
                                    // 🔥 CRITICAL DEBUG: Log XML file save
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 💾 XML FILE SAVED: {xmlFile} with {markedCount} cluster-resolved clash zones\n");
                                }
                                catch (Exception ex)
                                {
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR SAVING XML FILE: {xmlFile} - {ex.Message}\n");
                                }
                            }
                            else
                            {
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ NO CHANGES MADE - XML FILE NOT SAVED: {xmlFile}\n");
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
                    DebugLogger.Info($"[DEBUG] Marked {markedCount} clash zones as cluster-resolved\n");
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
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                    return;

                // ✅ FIX: Only process specific XML file if provided (ONE SOURCE OF TRUTH)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")  // Backward compatibility
                    : new[] { xmlFilePath };  // Only the specific file
                    
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
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
                                        DebugLogger.Info($"[DEBUG] Reset cluster flag for ClashZone {clashZone.Id} - ClusterSleeveInstanceId was {clashZone.ClusterSleeveInstanceId} (invalid)\n");
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
                                            DebugLogger.Info($"[DEBUG] Reset cluster flag for ClashZone {clashZone.Id} - cluster sleeve {clashZone.ClusterSleeveInstanceId} was deleted\n");
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
                                        DebugLogger.Info($"[DEBUG] Reset individual sleeve flag for ClashZone {clashZone.Id} - SleeveInstanceId was {clashZone.SleeveInstanceId} (invalid)\n");
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
                                            DebugLogger.Info($"[DEBUG] Reset individual sleeve flag for ClashZone {clashZone.Id} - sleeve {clashZone.SleeveInstanceId} was deleted\n");
                                        }
                                    }
                                }
                            }
                            
                            // ✅ FIX: Only save if modifications were made, and reader is already closed
                            if (modified)
                            {
                                // Normalize coordinates before saving to avoid 0,0,0 in XML
                                if (filter?.ClashZoneStorage?.ClashZones != null)
                                {
                                    foreach (var z in filter.ClashZoneStorage.ClashZones)
                                    {
                                        if (z == null) continue;
                                        bool hasIP = z.IntersectionPoint != null;
                                        bool isIPZero = hasIP && Math.Abs(z.IntersectionPoint.X) < 1e-9 && Math.Abs(z.IntersectionPoint.Y) < 1e-9 && Math.Abs(z.IntersectionPoint.Z) < 1e-9;
                                        bool hasSPP = z.SleevePlacementPoint != null;
                                        bool isSPPZero = hasSPP && Math.Abs(z.SleevePlacementPoint.X) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Y) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Z) < 1e-9;
                                        
                                        if (hasIP && !isIPZero)
                                        {
                                            z.IntersectionPointX = z.IntersectionPoint.X;
                                            z.IntersectionPointY = z.IntersectionPoint.Y;
                                            z.IntersectionPointZ = z.IntersectionPoint.Z;
                                        }
                                        else if (hasSPP && !isSPPZero)
                                        {
                                            z.IntersectionPointX = z.SleevePlacementPoint.X;
                                            z.IntersectionPointY = z.SleevePlacementPoint.Y;
                                            z.IntersectionPointZ = z.SleevePlacementPoint.Z;
                                        }
                                        else if (z.ClashBoundingBox != null)
                                        {
                                            var center = (z.ClashBoundingBox.Min + z.ClashBoundingBox.Max) / 2.0;
                                            z.IntersectionPointX = center.X;
                                            z.IntersectionPointY = center.Y;
                                            z.IntersectionPointZ = center.Z;
                                        }
                                    }
                                }
                                using (var writer = new StreamWriter(xmlFile))
                                {
                                    serializer.Serialize(writer, filter);
                                }
                                DebugLogger.Info($"[DEBUG] ✓ Saved reset flags to {Path.GetFileName(xmlFile)}\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalClusterService] Error processing XML file {xmlFile}: {ex.Message}");
                        DebugLogger.Info($"[DEBUG] ✗ Error resetting flags in {Path.GetFileName(xmlFile)}: {ex.Message}\n");
                    }
                }

                if (resetCount > 0)
                {
                    DebugLogger.Info($"[DEBUG] ✓ Reset flags for {resetCount} deleted sleeves (cluster + individual)\n");
                }
                else
                {
                    DebugLogger.Info($"[DEBUG] ✓ No deleted sleeves found - all flags preserved\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error resetting cluster flags: {ex.Message}");
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

        private Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, dynamic>> sleeveGroups, 
            double toleranceDist)
        {
            var clustersByGroup = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
            
            // ✅ FIXED: Use XML data directly for clustering - no Revit API calls needed

            foreach (var group in sleeveGroups)
            {
                var xmlSleeves = group.ToList();
                var groupClusters = new List<List<dynamic>>();
                
                DebugLogger.Log($"[UniversalClusterService] Processing group {group.Key.hostType}_{group.Key.systemType} with {xmlSleeves.Count} sleeves");
                DebugLogger.Info($"[DEBUG] GROUP: {group.Key.hostType}_{group.Key.systemType} - {xmlSleeves.Count} sleeves\n");
                
                // ✅ FIXED: Use bounding box overlap algorithm with XML data
                var clusters = CalculateClustersUsingXmlData(xmlSleeves, toleranceDist, group.Key.orientation);
                
                if (clusters.Count > 0)
                {
                    groupClusters.AddRange(clusters);
                    DebugLogger.Log($"[UniversalClusterService] Formed {clusters.Count} clusters using XML data");
                    DebugLogger.Info($"[DEBUG] ✓ CLUSTERS FORMED: {clusters.Count} clusters using XML data\n");
                    
                    // Log cluster details
                    foreach (var cluster in clusters)
                    {
                        if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log",
                            $"  - Cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.SleeveInstanceId))}\n");
                    }
                }
                else
                {
                    DebugLogger.Info($"[DEBUG] ℹ️ NO CLUSTERS: No proximate sleeves found in this group\n");
                }
                
                clustersByGroup[group.Key] = groupClusters;
            }

            return clustersByGroup;
        }
        
        /// <summary>
        /// ✅ FIXED: Calculate clusters using XML data directly - no Revit API calls needed
        /// </summary>
        private List<List<dynamic>> CalculateClustersUsingXmlData(List<dynamic> xmlSleeves, double toleranceDist, string orientation)
        {
            var clusters = new List<List<dynamic>>();
            var processed = new HashSet<int>();
            
            DebugLogger.Log($"[UniversalClusterService] Starting XML clustering for {xmlSleeves.Count} sleeves with tolerance {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm");
            
            foreach (var sleeve in xmlSleeves)
            {
                if (processed.Contains(sleeve.SleeveInstanceId))
                    continue;
                    
                var cluster = new List<dynamic> { sleeve };
                processed.Add(sleeve.SleeveInstanceId);
                
                // ✅ IMPROVED: Use iterative expansion to find all connected sleeves
                bool foundNewNeighbors = true;
                while (foundNewNeighbors)
                {
                    foundNewNeighbors = false;
                    var currentClusterSize = cluster.Count;
                    
                    // Find neighbors for all sleeves in current cluster
                    foreach (var clusterSleeve in cluster.ToList())
                    {
                        foreach (var otherSleeve in xmlSleeves)
                        {
                            if (processed.Contains(otherSleeve.SleeveInstanceId))
                                continue;
                                
                            if (BoundingBoxesOverlapFromXml(clusterSleeve, otherSleeve, toleranceDist))
                            {
                                cluster.Add(otherSleeve);
                                processed.Add(otherSleeve.SleeveInstanceId);
                                foundNewNeighbors = true;
                                DebugLogger.Log($"[UniversalClusterService] Added sleeve {otherSleeve.SleeveInstanceId} to cluster (now {cluster.Count} sleeves)");
                            }
                        }
                    }
                }
                
                if (cluster.Count > 1) // Only add clusters with multiple sleeves
                {
                    clusters.Add(cluster);
                    DebugLogger.Log($"[UniversalClusterService] Final cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.SleeveInstanceId))}");
                }
            }
            
            DebugLogger.Log($"[UniversalClusterService] Formed {clusters.Count} clusters from {xmlSleeves.Count} sleeves");
            return clusters;
        }
        
        /// <summary>
        /// Calculate cluster bounding box from XML data
        /// </summary>
        private (double width, double height, double depth, XYZ mid) GetClusterBoundingBoxFromXml(List<dynamic> cluster)
        {
            try
            {
                if (cluster.Count == 0)
                    return (0, 0, 0, XYZ.Zero);

                // Get all bounding boxes from XML data
                var boundingBoxes = cluster.Select(s => s.BoundingBox).Where(bbox => bbox != null).ToList();
                
                if (boundingBoxes.Count == 0)
                    return (0, 0, 0, XYZ.Zero);

                // ✅ DEBUG: Log individual sleeve bounding boxes
                DebugLogger.Info($"\n[CLUSTER-BBOX] Individual sleeves in cluster ({cluster.Count} sleeves):\n");
                
                foreach (var sleeve in cluster)
                {
                    var bbox = sleeve.BoundingBox;
                    if (bbox != null)
                    {
                        double sleeveWidth = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Millimeters);
                        double sleeveHeight = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Millimeters);
                        double sleeveDepth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Millimeters);
                        
                        if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                            $"  Sleeve {sleeve.SleeveInstanceId}: W={sleeveWidth:F1}mm, H={sleeveHeight:F1}mm, D={sleeveDepth:F1}mm, " +
                            $"Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}), Max=({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})\n");
                    }
                }

                // ✅ NEW: Calculate total volume accounting for overlaps
                double totalVolume = 0.0;
                var processedPairs = new HashSet<string>();
                
                // Calculate individual volumes and subtract overlaps
                for (int i = 0; i < boundingBoxes.Count; i++)
                {
                    var bbox1 = boundingBoxes[i];
                    double vol1 = (bbox1.Max.X - bbox1.Min.X) * (bbox1.Max.Y - bbox1.Min.Y) * (bbox1.Max.Z - bbox1.Min.Z);
                    totalVolume += vol1;
                    
                    // Subtract overlaps with previously processed boxes
                    for (int j = 0; j < i; j++)
                    {
                        var bbox2 = boundingBoxes[j];
                        var pairKey = $"{Math.Min(i, j)}-{Math.Max(i, j)}";
                        
                        if (!processedPairs.Contains(pairKey))
                        {
                            // Calculate overlap volume
                            double overlapX = Math.Max(0, Math.Min(bbox1.Max.X, bbox2.Max.X) - Math.Max(bbox1.Min.X, bbox2.Min.X));
                            double overlapY = Math.Max(0, Math.Min(bbox1.Max.Y, bbox2.Max.Y) - Math.Max(bbox1.Min.Y, bbox2.Min.Y));
                            double overlapZ = Math.Max(0, Math.Min(bbox1.Max.Z, bbox2.Max.Z) - Math.Max(bbox1.Min.Z, bbox2.Min.Z));
                            double overlapVol = overlapX * overlapY * overlapZ;
                            
                            if (overlapVol > 0)
                            {
                                totalVolume -= overlapVol;
                                
                                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                                    $"[OVERLAP] Sleeves {i} and {j}: Overlap Vol={UnitUtils.ConvertFromInternalUnits(overlapVol, UnitTypeId.CubicMillimeters):F2}mm³, " +
                                    $"Overlap Dim=({UnitUtils.ConvertFromInternalUnits(overlapX, UnitTypeId.Millimeters):F1}mm, " +
                                    $"{UnitUtils.ConvertFromInternalUnits(overlapY, UnitTypeId.Millimeters):F1}mm, " +
                                    $"{UnitUtils.ConvertFromInternalUnits(overlapZ, UnitTypeId.Millimeters):F1}mm)\n");
                            }
                            
                            processedPairs.Add(pairKey);
                        }
                    }
                }
                
                // Calculate overall bounding box (union of all boxes)
                double minX = boundingBoxes.Min(bbox => bbox.Min.X);
                double minY = boundingBoxes.Min(bbox => bbox.Min.Y);
                double minZ = boundingBoxes.Min(bbox => bbox.Min.Z);
                double maxX = boundingBoxes.Max(bbox => bbox.Max.X);
                double maxY = boundingBoxes.Max(bbox => bbox.Max.Y);
                double maxZ = boundingBoxes.Max(bbox => bbox.Max.Z);

                double width = maxX - minX;
                double height = maxY - minY;
                double depth = maxZ - minZ;
                XYZ mid = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);

                return (width, height, depth, mid);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error calculating cluster bounding box from XML: {ex.Message}");
                return (0, 0, 0, XYZ.Zero);
            }
        }

        /// <summary>
        /// Get reference level from XML data
        /// </summary>
        private Level GetReferenceLevelFromXml(Document doc, dynamic sleeve)
        {
            try
            {
                // For now, get the first level in the document as a fallback
                // In a real implementation, this would need to be stored in the XML data
                var levels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                return levels.FirstOrDefault() ?? levels.First();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting reference level from XML: {ex.Message}");
                return null;
            }
        }
        private bool BoundingBoxesOverlapFromXml(dynamic sleeve1, dynamic sleeve2, double toleranceDist)
        {
            try
            {
                // Use bounding box coordinates from XML data
                var bbox1 = sleeve1.BoundingBox;
                var bbox2 = sleeve2.BoundingBox;
                
                if (bbox1 == null || bbox2 == null)
                    return false;
                
                // Check overlap with tolerance
                bool xOverlap = (bbox1.Min.X - toleranceDist) <= bbox2.Max.X && (bbox1.Max.X + toleranceDist) >= bbox2.Min.X;
                bool yOverlap = (bbox1.Min.Y - toleranceDist) <= bbox2.Max.Y && (bbox1.Max.Y + toleranceDist) >= bbox2.Min.Y;
                bool zOverlap = (bbox1.Min.Z - toleranceDist) <= bbox2.Max.Z && (bbox1.Max.Z + toleranceDist) >= bbox2.Min.Z;
                
                return xOverlap && yOverlap && zOverlap;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error checking bounding box overlap: {ex.Message}");
                return false;
            }
        }
        private List<List<FamilyInstance>> CalculateClustersUsingBoundingBoxOverlap(List<FamilyInstance> sleeves, double toleranceDist, string orientation)
        {
            var clusters = new List<List<FamilyInstance>>();
            var unprocessedSet = new HashSet<FamilyInstance>(sleeves);
            
            DebugLogger.Log($"[UniversalClusterService] Starting bounding box overlap clustering for {sleeves.Count} sleeves with tolerance {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm");
            
            while (unprocessedSet.Count > 0)
            {
                var startSleeve = unprocessedSet.First();
                var queue = new Queue<FamilyInstance>();
                var cluster = new List<FamilyInstance>();
                
                queue.Enqueue(startSleeve);
                unprocessedSet.Remove(startSleeve);
                
                while (queue.Count > 0)
                {
                    var currentSleeve = queue.Dequeue();
                    cluster.Add(currentSleeve);
                    
                    // Find proximate sleeves using bounding box overlap
                    var proximateSleeves = FindProximateSleevesUsingBoundingBox(currentSleeve, unprocessedSet, toleranceDist, orientation);
                    
                    foreach (var proximateSleeve in proximateSleeves)
                    {
                        if (unprocessedSet.Remove(proximateSleeve))
                        {
                            queue.Enqueue(proximateSleeve);
                        }
                    }
                }
                
                // Only add clusters with more than 1 sleeve
                if (cluster.Count > 1)
                {
                    clusters.Add(cluster);
                    DebugLogger.Log($"[UniversalClusterService] Formed cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.Id.IntegerValue))}");
                    if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                        $"✓ CLUSTER: {string.Join(", ", cluster.Select(s => s.Id.IntegerValue))}\n");
                }
                else
                {
                    DebugLogger.Log($"[UniversalClusterService] Individual sleeve {cluster[0].Id.IntegerValue} (no proximate neighbors)");
                    DebugLogger.Info($"✗ NO CLUSTER: Individual sleeve {cluster[0].Id.IntegerValue} (no proximate neighbors within {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm)\n");
                }
            }
            
            return clusters;
        }
        
        /// <summary>
        /// ✅ NEW: Find proximate sleeves using bounding box overlap with saved coordinates
        /// This uses the working bounding box overlap algorithm but with pre-saved coordinates
        /// </summary>
        private List<FamilyInstance> FindProximateSleevesUsingBoundingBox(FamilyInstance currentSleeve, HashSet<FamilyInstance> candidateSleeves, double toleranceDist, string orientation)
        {
            var proximateSleeves = new List<FamilyInstance>();
            
            // Get current sleeve's clash zone
            var currentMepElementIdParam = currentSleeve.LookupParameter("MEP_ElementId");
            if (currentMepElementIdParam == null) return proximateSleeves;
            
            long currentMepElementId = currentMepElementIdParam.AsInteger();
            if (!_clashZoneCache.ContainsKey(currentMepElementId)) return proximateSleeves;
            
            var currentClashZone = _clashZoneCache[currentMepElementId];
            
            foreach (var candidateSleeve in candidateSleeves)
            {
                if (candidateSleeve == currentSleeve) continue;
                
                // Get candidate sleeve's clash zone
                var candidateMepElementIdParam = candidateSleeve.LookupParameter("MEP_ElementId");
                if (candidateMepElementIdParam == null) continue;
                
                long candidateMepElementId = candidateMepElementIdParam.AsInteger();
                if (!_clashZoneCache.ContainsKey(candidateMepElementId)) continue;
                
                var candidateClashZone = _clashZoneCache[candidateMepElementId];
                
                // Check bounding box overlap using saved coordinates
                // ✅ FIX: Pass orientation from grouping logic instead of relying on ClashZone orientation
                bool isOverlapping = CheckBoundingBoxOverlapWithOrientation(
                    currentClashZone, candidateClashZone, toleranceDist, orientation);
                
                if (isOverlapping)
                {
                    proximateSleeves.Add(candidateSleeve);
                    
                    // Log proximity analysis
                    DebugLogger.Info($"[DEBUG] ✓ BOUNDING-BOX-PROXIMITY: Sleeve {currentSleeve.Id.IntegerValue} -> {candidateSleeve.Id.IntegerValue}: Bounding boxes overlap within {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
                }
                else
                {
                    // Log non-proximate analysis (only for specific sleeves to reduce noise)
                    if (currentSleeve.Id.IntegerValue == 897149 || currentSleeve.Id.IntegerValue == 897154 || currentSleeve.Id.IntegerValue == 897195 ||
                        candidateSleeve.Id.IntegerValue == 897149 || candidateSleeve.Id.IntegerValue == 897154 || candidateSleeve.Id.IntegerValue == 897195)
                    {
                        DebugLogger.Info($"[DEBUG] ✗ NO PROXIMITY: Sleeve {currentSleeve.Id.IntegerValue} <-> {candidateSleeve.Id.IntegerValue}: Bounding boxes do not overlap within {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
                    }
                }
            }
            
            return proximateSleeves;
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Check bounding box overlap using orientation from grouping logic
        /// This bypasses the ClashZone orientation logic and uses the correct coordinates directly
        /// 
        /// ⚠️ CRITICAL PROTECTION: DO NOT MODIFY THIS METHOD WITHOUT UNDERSTANDING THE IMPACT ⚠️
        /// This method is essential for correct floor sleeve clustering:
        /// - Floor sleeves MUST use 2D X,Y distance calculation (ignoring Z)
        /// - The orientation parameter from GetEffectiveOrientationForClustering() is "Floor" for all floor sleeves
        /// - DO NOT use MepElementOrientationDirection here - that's only for individual sleeve rotation
        /// 
        /// Bug Fixed: 2025-10-27 - Floor sleeves were not clustering because they were split into X/Y groups
        /// Root Cause: GetEffectiveOrientationForClustering() was using MepElementOrientationDirection for floors
        /// Fix: Return unified "Floor" orientation for ALL floor sleeves in GetEffectiveOrientationForClustering()
        /// Status: WORKING - VERIFIED with cable trays (clusters form correctly), but ducts still need investigation
        /// </summary>
        private bool CheckBoundingBoxOverlapWithOrientation(ClashZone current, ClashZone other, double toleranceDist, string orientation)
        {
            if (other == null) return false;
            
            double minDistance;
            
            // ✅ DEBUG: Log coordinates and orientation
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect1: Min=({current.SleeveBoundingBoxMinX:F3}, {current.SleeveBoundingBoxMinY:F3}), Max=({current.SleeveBoundingBoxMaxX:F3}, {current.SleeveBoundingBoxMaxY:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect2: Min=({other.SleeveBoundingBoxMinX:F3}, {other.SleeveBoundingBoxMinY:F3}), Max=({other.SleeveBoundingBoxMaxX:F3}, {other.SleeveBoundingBoxMaxY:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] HostType={current.StructuralElementType}, Orientation={orientation}\n");
            
            // ✅ CRITICAL FIX: Use orientation from grouping logic instead of ClashZone orientation
            // For floor sleeves, orientation="Floor" (unified for all floor sleeves)
            if (current.StructuralElementType == "Floor")
            {
                // Floor sleeves: Use X,Y distance only (ignore Z coordinate)
                minDistance = CalculateMinimumDistance2D(
                    current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxY,
                    other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY);
            }
            else if (current.StructuralElementType == "Wall" || current.StructuralElementType == "Structural Framing")
            {
                if (orientation == "X")
                {
                    // Wall/Framing sleeves (X orientation): Use X,Z distance only (ignore Y coordinate)
                    minDistance = CalculateMinimumDistance2D(
                        current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxZ);
                }
                else if (orientation == "Y")
                {
                    // Wall/Framing sleeves (Y orientation): Use Y,Z distance only (ignore X coordinate)
                    DebugLogger.Info($"[DISTANCE-DEBUG] Using Y,Z coordinates for Y-oriented walls\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect1 YZ: Min=({current.SleeveBoundingBoxMinY:F3}, {current.SleeveBoundingBoxMinZ:F3}), Max=({current.SleeveBoundingBoxMaxY:F3}, {current.SleeveBoundingBoxMaxZ:F3})\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect2 YZ: Min=({other.SleeveBoundingBoxMinY:F3}, {other.SleeveBoundingBoxMinZ:F3}), Max=({other.SleeveBoundingBoxMaxY:F3}, {other.SleeveBoundingBoxMaxZ:F3})\n");
                    minDistance = CalculateMinimumDistance2D(
                        current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                }
                else
                {
                    // Default for walls: Use Y,Z distance (most walls are Y-oriented)
                    minDistance = CalculateMinimumDistance2D(
                        current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                }
            }
            else
            {
                // Fallback: Use 3D distance for unknown host types
                minDistance = CalculateMinimumDistance3D(
                    current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, 
                    current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                    other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ,
                    other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
            }
            
            // Log the calculated distance (only for specific sleeves to reduce noise)
            bool shouldLogDistance = current.SleeveInstanceId == 897149 || current.SleeveInstanceId == 897154 || current.SleeveInstanceId == 897195 ||
                                     other.SleeveInstanceId == 897149 || other.SleeveInstanceId == 897154 || other.SleeveInstanceId == 897195;
            
            if (shouldLogDistance)
            {
                DebugLogger.Info($"[DISTANCE-DEBUG] Calculated distance: {UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm, Tolerance: {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
            }
            
            return minDistance <= toleranceDist;
        }

        /// <summary>
        /// Calculate minimum distance between two 2D rectangles
        /// </summary>
        private double CalculateMinimumDistance2D(double minX1, double minY1, double maxX1, double maxY1,
            double minX2, double minY2, double maxX2, double maxY2)
        {
            // Check if rectangles overlap
            bool xOverlap = !(maxX1 < minX2 || maxX2 < minX1);
            bool yOverlap = !(maxY1 < minY2 || maxY2 < minY1);
            
            if (xOverlap && yOverlap)
            {
                return 0; // Rectangles overlap
            }
            
            // Calculate minimum distance
            double dx = 0;
            double dy = 0;
            
            if (!xOverlap)
            {
                dx = Math.Min(Math.Abs(maxX1 - minX2), Math.Abs(maxX2 - minX1));
            }
            
            if (!yOverlap)
            {
                dy = Math.Min(Math.Abs(maxY1 - minY2), Math.Abs(maxY2 - minY1));
            }
            
            return Math.Sqrt(dx * dx + dy * dy);
        }

        /// <summary>
        /// Calculate minimum distance between two 3D bounding boxes
        /// </summary>
        private double CalculateMinimumDistance3D(double minX1, double minY1, double minZ1, double maxX1, double maxY1, double maxZ1,
            double minX2, double minY2, double minZ2, double maxX2, double maxY2, double maxZ2)
        {
            // Check if bounding boxes overlap
            bool xOverlap = !(maxX1 < minX2 || maxX2 < minX1);
            bool yOverlap = !(maxY1 < minY2 || maxY2 < minY1);
            bool zOverlap = !(maxZ1 < minZ2 || maxZ2 < minZ1);
            
            if (xOverlap && yOverlap && zOverlap)
            {
                return 0; // Bounding boxes overlap
            }
            
            // Calculate minimum distance
            double dx = 0;
            double dy = 0;
            double dz = 0;
            
            if (!xOverlap)
            {
                dx = Math.Min(Math.Abs(maxX1 - minX2), Math.Abs(maxX2 - minX1));
            }
            
            if (!yOverlap)
            {
                dy = Math.Min(Math.Abs(maxY1 - minY2), Math.Abs(maxY2 - minY1));
            }
            
            if (!zOverlap)
            {
                dz = Math.Min(Math.Abs(maxZ1 - minZ2), Math.Abs(maxZ2 - minZ1));
            }
            
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
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
            double toleranceDist,
            SleeveGroupKey groupKey)
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
                    var neighbors = FilterNeighborsByBoundingBox(inst, candidates, o1_bbox, bboxes, unprocessedSet, toleranceDist, groupKey);

                    foreach (var n in neighbors)
                    {
                        if (unprocessedSet.Remove(n)) queue.Enqueue(n);
                    }
                }

                // 🔥 PROXIMITY DEBUG: Log cluster results
                string clusterLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log";
                if (cluster.Count > 1)
                {
                    var clusterIds = string.Join(", ", cluster.Select(s => s.Id.IntegerValue));
                    File.AppendAllText(clusterLogPath, $"✅ CLUSTER FORMED: {cluster.Count} sleeves [{clusterIds}]\n");
                }
                else
                {
                    File.AppendAllText(clusterLogPath, $"❌ INDIVIDUAL SLEEVE: {cluster[0].Id.IntegerValue} (no proximate neighbors)\n");
                }

                groupClusters.Add(cluster);
            }

            // 🔥 PROXIMITY DEBUG: Log final summary
            string clusterLogPath2 = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log";
            bool __suppressClusterLogs2 = DeploymentConfiguration.DeploymentMode;
            int totalClusters = groupClusters.Count(c => c.Count > 1);
            int totalIndividuals = groupClusters.Count(c => c.Count == 1);
            File.AppendAllText(clusterLogPath2, $"\n=== CLUSTER SUMMARY ===\n");
            File.AppendAllText(clusterLogPath2, $"Total clusters formed: {totalClusters}\n");
            File.AppendAllText(clusterLogPath2, $"Total individual sleeves: {totalIndividuals}\n");
            File.AppendAllText(clusterLogPath2, $"Total sleeves processed: {groupClusters.Sum(c => c.Count)}\n\n");

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
        /// Filter neighbors based on bounding box overlap and clustering criteria
        /// </summary>
        private List<FamilyInstance> FilterNeighborsByBoundingBox(
            FamilyInstance inst,
            List<FamilyInstance> candidates,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            HashSet<FamilyInstance> unprocessedSet,
            double toleranceDist,
            SleeveGroupKey groupKey)
        {
            var neighbors = new List<FamilyInstance>();

            foreach (var candidate in candidates)
            {
                if (candidate == inst || !unprocessedSet.Contains(candidate))
                    continue;

                // Check if candidate matches the group criteria
                if (!MatchesGroupCriteria(candidate, groupKey))
                    continue;

                // Get bounding box for candidate
                BoundingBoxXYZ o2_bbox = bboxes.ContainsKey(candidate) ? bboxes[candidate] : candidate.get_BoundingBox(null);
                if (o2_bbox == null) continue;

                // Check bounding box overlap with tolerance
                if (BoundingBoxesOverlap(o1_bbox, o2_bbox, toleranceDist))
                {
                    neighbors.Add(candidate);
                }
            }

            return neighbors;
        }
        
        /// <summary>
        /// Check if two bounding boxes overlap within tolerance
        /// </summary>
        private bool BoundingBoxesOverlap(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2, double tolerance)
        {
            if (bbox1 == null || bbox2 == null) return false;

            return bbox1.Min.X - tolerance <= bbox2.Max.X &&
                   bbox1.Max.X + tolerance >= bbox2.Min.X &&
                   bbox1.Min.Y - tolerance <= bbox2.Max.Y &&
                   bbox1.Max.Y + tolerance >= bbox2.Min.Y &&
                   bbox1.Min.Z - tolerance <= bbox2.Max.Z &&
                   bbox1.Max.Z + tolerance >= bbox2.Min.Z;
        }

        /// <summary>
        /// Load clash zones from regular XML files and return them as a list
        /// </summary>
        private List<ClashZone> LoadClashZonesFromRegularXml(string xmlFilePath, string targetCategory, Document doc)
        {
            var clashZones = new List<ClashZone>();
            
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return clashZones;
                }

                // Load from regular XML files
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                foreach (var xmlFile in xmlFiles)
                {
                    // Skip CONDITIONS files
                    if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                        continue;
                        
                    // Skip _CLUSTER.xml files - we want regular XML files
                    if (Path.GetFileName(xmlFile).Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                        continue;

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
                                    // Filter by target category during loading
                                    if (!string.IsNullOrEmpty(targetCategory))
                                    {
                                        if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                continue;
                            }
                                    
                                    // Only process clash zones with valid SleeveInstanceId (placed sleeves)
                                    if (cz.SleeveInstanceId > 0)
                                    {
                                        // Reconstruct SleevePlacementPoint from XML-serializable properties
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        clashZones.Add(cz);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading clash zones from regular XML: {ex.Message}");
            }
            
            return clashZones;
        }

        /// <summary>
        /// Get host type from clash zone data
        /// </summary>
        private string GetHostTypeFromClashZone(ClashZone clashZone)
        {
            try
            {
                // Get host type from structural element type
                if (!string.IsNullOrEmpty(clashZone.StructuralElementType))
                {
                    if (clashZone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase))
                        return "Wall";
                    else if (clashZone.StructuralElementType.Contains("Floor", StringComparison.OrdinalIgnoreCase) || 
                             clashZone.StructuralElementType.Contains("Slab", StringComparison.OrdinalIgnoreCase))
                        return "Floor";
                    else if (clashZone.StructuralElementType.Contains("Structural Framing", StringComparison.OrdinalIgnoreCase))
                        return "Structural Framing";
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting host type from clash zone: {ex.Message}");
            }
            return "Unknown";
        }

        /// <summary>
        /// Get effective orientation for clustering based on host type
        /// </summary>
        private string GetEffectiveOrientationForClustering(ClashZone clashZone)
        {
            try
            {
                string hostType = GetHostTypeFromClashZone(clashZone);
                
                // For walls and framing, use HostOrientation (X/Y)
                if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    return clashZone.HostOrientation ?? "Unknown";
                }
                // ✅ CRITICAL FIX: For floors, return "Floor" to group ALL floor sleeves together for clustering
                // DO NOT use MepElementOrientationDirection - that's only for individual sleeve rotation
                // For clustering, ALL floor sleeves should be grouped by category and host type only
                else if (hostType == "Floor")
                {
                    return "Floor"; // ✅ FIXED: Return unified "Floor" instead of X/Y split
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting effective orientation for clustering: {ex.Message}");
            }
            return "Unknown";
        }

        /// <summary>
        /// Get bounding box from clash zone data
        /// </summary>
        private BoundingBoxXYZ GetBoundingBoxFromClashZone(ClashZone clashZone)
        {
            try
            {
                // Create bounding box from stored coordinates
                var bbox = new BoundingBoxXYZ();
                bbox.Min = new XYZ(clashZone.SleeveBoundingBoxMinX, clashZone.SleeveBoundingBoxMinY, clashZone.SleeveBoundingBoxMinZ);
                bbox.Max = new XYZ(clashZone.SleeveBoundingBoxMaxX, clashZone.SleeveBoundingBoxMaxY, clashZone.SleeveBoundingBoxMaxZ);
                return bbox;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting bounding box from clash zone: {ex.Message}");
            return null;
            }
        }
        
        /// <summary>
        /// Check if a sleeve matches the group criteria for clustering
        /// </summary>
        private bool MatchesGroupCriteria(FamilyInstance sleeve, SleeveGroupKey groupKey)
        {
            try
            {
                // ✅ SIMPLIFIED: For clustering, we only need host type and orientation from sleeve directly
                // Category is already filtered at the sleeve collection level
                
                // Get host type from sleeve
                string hostType = GetHostTypeFromSleeve(sleeve);
                if (!string.Equals(hostType, groupKey.hostType, StringComparison.OrdinalIgnoreCase))
                    return false;

                // Get orientation from sleeve
                string orientation = GetOrientationFromSleeve(sleeve);
                if (!string.Equals(orientation, groupKey.orientation, StringComparison.OrdinalIgnoreCase))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error checking group criteria: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Get host type from sleeve element
        /// </summary>
        private string GetHostTypeFromSleeve(FamilyInstance sleeve)
        {
            try
            {
                var host = sleeve.Host;
                if (host != null)
                {
                    var category = host.Category;
                    if (category != null)
                    {
                        string categoryName = category.Name;
                        if (categoryName.Contains("Wall", StringComparison.OrdinalIgnoreCase))
                            return "Wall";
                        else if (categoryName.Contains("Floor", StringComparison.OrdinalIgnoreCase) || categoryName.Contains("Slab", StringComparison.OrdinalIgnoreCase))
                            return "Floor";
                        else if (categoryName.Contains("Structural Framing", StringComparison.OrdinalIgnoreCase))
                            return "Structural Framing";
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting host type from sleeve: {ex.Message}");
            }
            return "Unknown";
        }

        /// <summary>
        /// Get orientation from sleeve element directly
        /// 
        /// ⚠️ CRITICAL PROTECTION: DO NOT MODIFY THIS METHOD WITHOUT UNDERSTANDING THE IMPACT ⚠️
        /// This method is essential for correct sleeve grouping during clustering:
        /// - For floor sleeves, return "Floor" to match GetEffectiveOrientationForClustering()
        /// - For wall/framing sleeves, return X/Y based on Wall Direction Type parameter
        /// 
        /// Bug Fixed: 2025-10-27 - Floor sleeves were returning "Unknown" from Wall Direction Type parameter
        /// Root Cause: GetOrientationFromSleeve() only checked "Wall Direction Type" parameter (for walls/framing)
        /// Fix: Added host type check - if host is Floor, return "Floor"; otherwise use Wall Direction Type
        /// Status: WORKING - This ensures MatchesGroupCriteria() can correctly match sleeves to groups
        /// </summary>
        private string GetOrientationFromSleeve(FamilyInstance sleeve)
        {
            try
            {
                // ✅ CRITICAL FIX: Check host type first - floor sleeves have different orientation logic
                string hostType = GetHostTypeFromSleeve(sleeve);
                if (hostType == "Floor")
                {
                    // For floor sleeves, return "Floor" to match GetEffectiveOrientationForClustering()
                    return "Floor";
                }
                
                // For wall/framing sleeves, get orientation from Wall Direction Type parameter
                var wallDirectionParam = sleeve.LookupParameter("Wall Direction Type");
                if (wallDirectionParam != null)
                {
                    string wallDirection = wallDirectionParam.AsString();
                    if (!string.IsNullOrEmpty(wallDirection))
                    {
                        // Convert wall direction to orientation
                        if (wallDirection.Contains("X", StringComparison.OrdinalIgnoreCase))
                            return "X";
                        else if (wallDirection.Contains("Y", StringComparison.OrdinalIgnoreCase))
                            return "Y";
                        else if (wallDirection.Contains("Z", StringComparison.OrdinalIgnoreCase))
                            return "Z";
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting orientation from sleeve: {ex.Message}");
            }
            return "Unknown";
        }

        /// <summary>
        /// Cache for clash zone data to avoid expensive lookups during clustering
        /// Key: MEP Element ID, Value: ClashZone data
        /// </summary>
        private static Dictionary<long, ClashZone> _clashZoneCache = new Dictionary<long, ClashZone>();
        private static List<SleeveData> _loadedSleeveDataCache = new List<SleeveData>();

        /// <summary>
        /// Pre-calculated cluster service for "calculate once, use many times" approach
        /// </summary>
        // ✅ REMOVED: PreCalculatedClusterService is no longer used - replaced with cheaper bounding box overlap algorithm

        /// <summary>
        /// ✅ DYNAMIC: Load clash zone cache from regular XML files (not _CLUSTER.xml)
        /// This reads from the XML files that UniversalSleevePlacerService saves to
        /// </summary>
        /// <summary>
        /// Public method to reload cache for cleanup after XML update
        /// </summary>
        public void LoadClashZoneCacheForCleanup(string xmlFilePath, string targetCategory, Document doc, string filterName)
        {
            // ✅ CRITICAL: Clear cache before reloading to get fresh data from XML
            _clashZoneCache.Clear();
            
            DebugLogger.Info($"[UniversalClusterService] Loading cache for cleanup with xmlFilePath={xmlFilePath}, targetCategory={targetCategory}");
            LoadClashZoneCacheFromRegularXml(xmlFilePath, targetCategory, doc, filterName);
            
            // Log what was loaded
            DebugLogger.Info($"[UniversalClusterService] Cache loaded with {_clashZoneCache.Count} clash zones");
            foreach (var kvp in _clashZoneCache.Take(5))
            {
                var cz = kvp.Value;
                DebugLogger.Info($"[UniversalClusterService] Cached: SleeveId={cz.SleeveInstanceId}, AfterClusterId={cz.AfterClusterSleevePlacedSleeveInstanceId}, ClusterId={cz.ClusterSleeveInstanceId}");
            }
        }
        
        /// <summary>
        /// Public method to run cleanup after XML save (uses cache instead of Revit API)
        /// </summary>
        public int CleanupSleevesWithinClustersAfterXmlSave(Document doc, List<FamilyInstance> placedClusters, string xmlFilePath = null)
        {
            int deleted = CleanupSleevesWithinClusters(doc, placedClusters);
            
            // ✅ CRITICAL: Save updated flags to XML after cleanup
            if (deleted > 0 && !string.IsNullOrEmpty(xmlFilePath))
            {
                SaveUpdatedFlagsToXml(xmlFilePath);
            }
            
            return deleted;
        }
        
        private void LoadClashZoneCacheFromRegularXml(string xmlFilePath, string targetCategory = null, Document doc = null, string filterName = null)
        {
            _clashZoneCache.Clear();
            
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return;
                }

                // ✅ DYNAMIC: Load from regular XML files (the ones UniversalSleevePlacerService saves to)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                foreach (var xmlFile in xmlFiles)
                {
                    // Skip CONDITIONS files
                    if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                        continue;
                        
                    // Skip _CLUSTER.xml files - we want regular XML files
                    if (Path.GetFileName(xmlFile).Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (File.Exists(xmlFile))
                    {
                        DebugLogger.Info($"[UniversalClusterService] Loading clash zones from regular XML: {Path.GetFileName(xmlFile)}");
                        
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var cz in filter.ClashZoneStorage.ClashZones)
                                {
                                    // ✅ DYNAMIC: Filter by target category during loading
                                    if (!string.IsNullOrEmpty(targetCategory))
                                    {
                                        if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                        {
                                            DebugLogger.Log($"[UniversalClusterService] SKIP: ClashZone {cz.Id} category '{cz.MepElementCategory}' doesn't match target '{targetCategory}'");
                                continue;
                            }
                                    }
                                    
                                    // ✅ DYNAMIC: Process clash zones with valid SleeveInstanceId (placed sleeves) OR valid ClusterSleeveInstanceId OR valid AfterClusterSleevePlacedSleeveInstanceId (for cleanup)
                                    if (cz.SleeveInstanceId > 0 || cz.ClusterSleeveInstanceId > 0 || cz.AfterClusterSleevePlacedSleeveInstanceId > 0)
                                    {
                                    // Cache by MEP element ID for fast lookup
                                    if (cz.MepElementIdValue > 0)
                                    {
                                            // ✅ CRITICAL: Reconstruct SleevePlacementPoint from XML-serializable properties
                                        cz.EnsureSleevePlacementPointReconstructed();
                                            
                                            // ✅ DYNAMIC: Mark as current clash for clustering
                                            cz.IsCurrentClash = true;
                                        
                                        _clashZoneCache[cz.MepElementIdValue] = cz;
                                            
                                            DebugLogger.Log($"[UniversalClusterService] LOADED: ClashZone {cz.Id} with SleeveInstanceId={cz.SleeveInstanceId}, ClusterSleeveInstanceId={cz.ClusterSleeveInstanceId} from {Path.GetFileName(xmlFile)}");
                                        }
                                    }
                                    else
                                    {
                                        DebugLogger.Log($"[UniversalClusterService] SKIP: ClashZone {cz.Id} has no SleeveInstanceId or ClusterSleeveInstanceId (not placed yet)");
                                    }
                                }
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[UniversalClusterService] Loaded {_clashZoneCache.Count} clash zones from regular XML files");
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[CACHE] Loaded {_clashZoneCache.Count} clash zones from regular XML files (filtered for {targetCategory ?? "ALL"})\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading clash zone cache from regular XML: {ex.Message}");
            }
        }

        /// <summary>
        /// Get clash zone from cache by sleeve instance ID
        /// </summary>
        private ClashZone GetClashZoneBySleeveInstanceId(int sleeveInstanceId)
        {
            try
            {
                // Search through all loaded clash zones for matching SleeveInstanceId
                foreach (var clashZone in _clashZoneCache.Values)
                {
                    if (clashZone.SleeveInstanceId == sleeveInstanceId)
                    {
                        DebugLogger.Log($"[UniversalClusterService] Found clash zone for Sleeve Instance ID {sleeveInstanceId} in cache");
                        return clashZone;
                        }
                    }
                
                DebugLogger.Warning($"[UniversalClusterService] Sleeve Instance ID {sleeveInstanceId} not found in cache (cache size: {_clashZoneCache?.Count ?? 0})");
            return null;
        }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting clash zone by sleeve instance ID: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ PERFORMANCE: Fast O(1) lookup using clash cache
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, string xmlFilePath = null)
        {
            try
            {
                // ✅ PERFORMANCE: Fast O(1) lookup using clash cache
                if (_clashZoneCache != null && _clashZoneCache.TryGetValue(mepElementId, out ClashZone clashZone))
                {
                    DebugLogger.Log($"[UniversalClusterService] Found clash zone for MEP Element ID {mepElementId} in cache");
                    return clashZone;
                }
                
                DebugLogger.Warning($"[UniversalClusterService] MEP Element ID {mepElementId} not found in cache (cache size: {_clashZoneCache?.Count ?? 0})");

                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);

                if (!Directory.Exists(filtersDirectory))
                    return null;

                // ✅ FIX: Only search specific XML file if provided (ONE SOURCE OF TRUTH)
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
                                    if (cz.MepElementIdValue == mepElementId)
                                    {
                                        // ✅ CRITICAL FIX: Reconstruct SleevePlacementPoint from XML-serializable properties
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        cz.EnsureSleevePlacementPointActiveDocumentReconstructed();

                                        return cz;
                                    }
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error getting clash zone by MEP element ID: {ex.Message}");
                return null;
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
        /// Load clash zones from XML for a specific category
        /// </summary>
        private List<ClashZone> LoadClashZonesFromXml(string category, Document doc)
        {
            var clashZones = new List<ClashZone>();

            try
            {
                var filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");

                if (!Directory.Exists(filtersDirectory))
                    return clashZones;

                var pattern = $"*_{category}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);

                if (matchingFiles.Length == 0)
                    return clashZones;

                var xmlFile = matchingFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();

                var serializer = new XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFile))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    if (filter?.ClashZoneStorage?.ClashZones != null)
                    {
                        clashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error loading clash zones from XML: {ex.Message}");
            }

            return clashZones;
        }

        private void PlaceClusterSleeve(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            out int placed,
            out int deleted,
            string xmlFilePath = null)
        {
            placed = 0;
            deleted = 0;

            // ✅ FIXED: Work with XML data directly - determine cluster properties from XML
            // Determine if cluster is circular or rectangular based on XML data
            bool isCircular = cluster.All(s => 
            {
                // Check if any sleeve in cluster has circular properties
                // For now, default to rectangular for pipes
                return false; // Simplified for XML data
            });
            
            // PIPE CLUSTERS: Always rectangular regardless of member shape (legacy parity)
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

                DebugLogger.Log($"[ClusterService] Creating {groupKey.systemType} cluster using family '{familyName}' (shape: {(isCircular ? "Circular" : "Rectangular")}, orientation: {groupKey.orientation})");

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

            // ✅ HYBRID APPROACH: Use XML for clustering, Revit API for accurate bounding box
            // Step 1: Convert XML cluster to actual Revit sleeves
            List<FamilyInstance> actualSleeves = new List<FamilyInstance>();
            foreach (var xmlSleeve in cluster)
            {
                try
                {
                    var sleeveId = new ElementId(xmlSleeve.SleeveInstanceId);
                    var sleeve = doc.GetElement(sleeveId) as FamilyInstance;
                    if (sleeve != null)
                    {
                        actualSleeves.Add(sleeve);
                    }
                    else
                    {
                        DebugLogger.Warning($"[ClusterService] Sleeve ID {xmlSleeve.SleeveInstanceId} not found in document for cluster");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ClusterService] Error getting sleeve {xmlSleeve.SleeveInstanceId} from document: {ex.Message}");
                }
            }
            
            if (actualSleeves.Count == 0)
            {
                DebugLogger.Error($"[ClusterService] No actual sleeves found for cluster group, skipping placement.");
                return;
            }
            
            // Step 2: Use ClusterBoundingBoxServices to get accurate bounding box from actual sleeves
            var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(actualSleeves);
            
            DebugLogger.Log($"[ClusterService] ✅ HYBRID APPROACH: Cluster bounding box from Revit API - Width={width:F3}, Height={height:F3}, Depth={depth:F3}");

            // ✅ FIXED: Get reference level from XML data (use first sleeve's level)
            Level? refLevel = GetReferenceLevelFromXml(doc, cluster[0]);
            if (refLevel == null)
            {
                DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                return;
            }

            // Place cluster sleeve
            FamilyInstance inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, refLevel!, StructuralType.NonStructural);

            // ✅ FIXED: Apply rotation based on XML data orientation (for both Walls and Framing)
            double rotationAngle = 0.0;
            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                // Get orientation from XML data (stored in MepElementOrientationDirection)
                string xmlOrientation = groupKey.orientation ?? "Unknown";
                
                // Apply rotation based on XML orientation data
                if (xmlOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                {
                    rotationAngle = Math.PI / 2;  // 90° rotation for X-oriented walls/framing
                }
                else if (xmlOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                {
                    rotationAngle = 0.0;  // No rotation for Y-oriented walls/framing
                }
                
                DebugLogger.Log($"[ClusterService] {groupKey.hostType} orientation from XML: '{xmlOrientation}', rotation angle: {rotationAngle * 180 / Math.PI}°");
                DebugLogger.Info($"[ROTATION] {groupKey.hostType} orientation: {xmlOrientation}, rotation: {rotationAngle * 180 / Math.PI}°\n");
            }

            if (rotationAngle != 0.0)
            {
                XYZ axisOrigin = mid;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
            }

            // Set size parameters (swap dimensions if rotated for orientation alignment)
            // ✅ FIX: Apply swap for both walls and framing when X-oriented
            bool shouldSwapDimensions = ((groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing") && rotationAngle != 0.0);
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);
            
            // ✅ DEBUG: Log before calling SetClusterSleeveMetadata
            DebugLogger.Info($"[PlaceClusterSleeve] About to call SetClusterSleeveMetadata for cluster sleeve {inst.Id}, targetCategory='{targetCategory}'");
            DebugLogger.Info($"[PlaceClusterSleeve] About to call SetClusterSleeveMetadata for cluster sleeve {inst.Id}, targetCategory='{targetCategory}'\n");
            
            // CRITICAL FIX: Set metadata parameters for cluster sleeve
            SetClusterSleeveMetadata(inst, targetCategory);

            // ✅ FIXED: Set MEP_ElementId on cluster sleeve using XML data
            var firstSleeve = cluster.FirstOrDefault();
            DebugLogger.Info($"[MEP_ElementId] First sleeve ID: {firstSleeve?.SleeveInstanceId}\n");
            
            if (firstSleeve != null)
            {
                // Get MEP Element ID from XML data (stored in clash zone)
                var clashZone = GetClashZoneBySleeveInstanceId(firstSleeve.SleeveInstanceId);
                if (clashZone != null && clashZone.MepElementIdValue > 0)
                {
                    long mepElementId = clashZone.MepElementIdValue;
                    DebugLogger.Info($"[MEP_ElementId] Value from XML: {mepElementId}\n");
                    
                    var clusterMepElementIdParam = inst.LookupParameter("MEP_ElementId");
                    DebugLogger.Info($"[MEP_ElementId] Parameter found on cluster: {clusterMepElementIdParam != null}, ReadOnly: {clusterMepElementIdParam?.IsReadOnly}\n");
                    
                    if (clusterMepElementIdParam != null && !clusterMepElementIdParam.IsReadOnly)
                    {
                        clusterMepElementIdParam.Set(mepElementId);
                        DebugLogger.Info($"[DEBUG] ✓ Set MEP_ElementId = {mepElementId} on cluster sleeve {inst.Id}\n");
                    }
                    else
                    {
                        DebugLogger.Info($"[DEBUG] ✗ Cannot set MEP_ElementId on cluster sleeve {inst.Id} (param null or readonly)\n");
                    }
                }
                else
                {
                    DebugLogger.Info($"[DEBUG] ✗ MEP_ElementId not found in XML for sleeve {firstSleeve.SleeveInstanceId}\n");
                }
            }

            placed++;

            // ⚠️ CRITICAL: Mark clash zones as cluster-resolved with the actual cluster sleeve ID
            MarkClashZonesAsClusterResolvedWithSleeveId(cluster, inst.Id, xmlFilePath);

            // ✅ GLOBAL FLAGS: Upsert IsResolved/IsClusterResolved into per-category global index after clustering
            try
            {
                var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved)>();
                foreach (var s in cluster)
                {
                    var cz = GetClashZoneBySleeveInstanceId(s.SleeveInstanceId);
                    if (cz != null)
                    {
                        updates.Add((cz.Id, cz.IsResolved, cz.IsClusterResolved));
                    }
                }

                if (updates.Count > 0)
                {
                    var updatesWithIds = updates.Select(u => (u.Id, u.IsResolved, u.IsClusterResolved, 0, inst.Id.IntegerValue)).ToList();
                    GlobalIndexService.UpsertFlagsWithIds(_doc, targetCategory, updatesWithIds);
                }
            }
            catch (Exception upEx)
            {
                DebugLogger.Warning($"[GLOBAL_INDEX] Upsert after clustering failed: {upEx.Message}");
            }

            DebugLogger.Log($"[ClusterService] About to delete {cluster.Count} individual sleeves for cluster sleeve {inst.Id.IntegerValue}");
            DebugLogger.Info($"[DELETE] About to delete {cluster.Count} individual sleeves for cluster sleeve {inst.Id.IntegerValue} (hostType={groupKey.hostType})\n");

            // Delete originals - collect ElementIds first, then delete in batch
            var sleevesToDelete = new List<ElementId>();
            foreach (var s in cluster)
            {
                try
                {
                    // Get the actual Revit sleeve by SleeveInstanceId
                    var sleeveElementId = new ElementId(s.SleeveInstanceId);
                    var sleeveElement = doc.GetElement(sleeveElementId);
                    
                    DebugLogger.Info($"[DELETE] Checking sleeve {sleeveElementId.IntegerValue}: found={sleeveElement != null}, isFamilyInstance={sleeveElement is FamilyInstance}\n");
                    
                    if (sleeveElement != null && sleeveElement is FamilyInstance sleeveInstance)
                    {
                        sleevesToDelete.Add(sleeveElementId);
                        DebugLogger.Log($"[ClusterService] Queued for deletion: individual sleeve {sleeveElementId.IntegerValue}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ClusterService] Could not find sleeve {s.SleeveInstanceId} for deletion");
                        DebugLogger.Info($"[DELETE] ✗ Sleeve {sleeveElementId.IntegerValue} not found or not a FamilyInstance\n");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ClusterService] Error preparing sleeve {s.SleeveInstanceId} for deletion: {ex.Message}");
                    DebugLogger.Info($"[DELETE] ✗ Exception preparing sleeve: {ex.Message}\n");
                }
            }
            
            DebugLogger.Info($"[DELETE] Total sleeves queued for deletion: {sleevesToDelete.Count}\n");
            
            // Delete all sleeves in batch (within the same transaction)
            if (sleevesToDelete.Count > 0)
            {
                try
                {
                    DebugLogger.Info($"[DELETE] Calling doc.Delete() for {sleevesToDelete.Count} sleeves (hostType={groupKey.hostType})\n");
                    DebugLogger.Info($"[DELETE] Document modifiable: {doc.IsModifiable}\n");
                    doc.Delete(sleevesToDelete);
                    deleted = sleevesToDelete.Count;
                    DebugLogger.Log($"[ClusterService] Successfully deleted {deleted} individual sleeves in batch");
                    DebugLogger.Info($"[DELETE] ✓ Successfully deleted {deleted} individual sleeves in batch (hostType={groupKey.hostType})\n");
                    
                    // ✅ VERIFY: Check if sleeves were actually deleted
                    int stillExists = 0;
                    foreach (var sleeveId in sleevesToDelete)
                    {
                        var stillThere = doc.GetElement(sleeveId);
                        if (stillThere != null)
                        {
                            stillExists++;
                            DebugLogger.Info($"[DELETE] ⚠️ WARNING: Sleeve {sleeveId.IntegerValue} still exists after deletion attempt!\n");
                        }
                    }
                    if (stillExists > 0)
                    {
                        DebugLogger.Info($"[DELETE] ⚠️ WARNING: {stillExists} sleeves still exist after deletion attempt!\n");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ClusterService] Error deleting sleeves in batch: {ex.Message}");
                    DebugLogger.Info($"[DELETE] ✗ Error deleting sleeves: {ex.Message} (hostType={groupKey.hostType})\n");
                    DebugLogger.Info($"[DELETE] Stack trace: {ex.StackTrace}\n");
                    deleted = 0;
                }
            }
            else
            {
                DebugLogger.Info($"[DELETE] ⚠️ No sleeves to delete (sleevesToDelete.Count=0) for hostType={groupKey.hostType}\n");
            }
        }

        private void SetClusterSizeParameters(
            Document doc,
            FamilyInstance inst,
            List<dynamic> cluster,
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
                
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
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
                // ✅ FIXED: Use XML data instead of Revit API calls
                var firstSleeve = cluster[0];
                double hostThickness = openingDepth;  // Default fallback to calculated depth
                
                // ✅ FIXED: Cluster sleeve should use individual sleeve thickness, not structural element thickness
                // For all host types, use the calculated depth from individual sleeves
                hostThickness = openingDepth;
                DebugLogger.Log($"[ClusterService] Cluster sleeve: Using individual sleeve depth {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm for {groupKey.hostType}");
                
                // ✅ FIXED: Use calculated depth as final fallback (no Revit API calls needed)
                if (hostThickness <= 0.0)
                {
                    hostThickness = openingDepth;  // Use calculated depth from bounding box
                    DebugLogger.Log($"[ClusterService] Using calculated depth as fallback: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                }
                
                DebugLogger.Info($"[CLUSTER-DIM] Sleeve {inst.Id}: Calculated Depth = {UnitUtils.ConvertFromInternalUnits(openingDepth, UnitTypeId.Millimeters):F1}mm, Final HostThickness = {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm\n");
                    
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
            List<dynamic> originalSleeves, 
            string systemType,
            Document doc)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterService] 🔥 UpdateClashZoneFlagsForCluster CALLED 🔥");
                DebugLogger.Info($"[UniversalClusterService] Cluster ID: {clusterInstance.Id.IntegerValue}, Original sleeves: {originalSleeves.Count}, SystemType: {systemType}");
                
                // ✅ DEBUG: Log the category mapping
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UpdateClashZoneFlagsForCluster: systemType='{systemType}', clusterId={clusterInstance.Id.IntegerValue}\n");

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
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Category mapping: '{systemType}' → '{category}'\n");

                // Load clash zones from XML
                var clashZones = LoadClashZonesFromRegularXml(null, category, doc);
                DebugLogger.Info($"[UniversalClusterService] Loaded {clashZones.Count} clash zones from XML");

                // Update each clash zone for deleted sleeves
                int updatedCount = 0;
                foreach (var sleeve in originalSleeves)
                {
                    // ✅ FIXED: Use XML data instead of Revit API calls
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    var clashZone = clashZones.FirstOrDefault(cz => 
                        cz.SleeveInstanceId == sleeveInstanceId);
                    
                    if (clashZone != null)
                    {
                        // Mark as clustered (using IsClusterResolved instead of IsClustered)
                        clashZone.IsClusterResolved = true;
                        clashZone.IsResolved = true;
                        clashZone.ClusterSleeveInstanceId = clusterInstance.Id.IntegerValue;
                        
                        // ✅ STORAGE: Save original SleeveInstanceId BEFORE clearing it
                        clashZone.AfterClusterSleevePlacedSleeveInstanceId = clashZone.SleeveInstanceId;
                        
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
        /// Find the newest cluster sleeve that was just placed
        /// </summary>
        private FamilyInstance FindNewestClusterSleeve(Document doc, List<dynamic> originalSleeves, SleeveGroupKey groupKey)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterService] 🔥 FindNewestClusterSleeve CALLED 🔥");
                DebugLogger.Info($"[UniversalClusterService] Host Type: {groupKey.hostType}, System Type: {groupKey.systemType}");

                // ✅ FIXED: Use XML data instead of Revit API calls
                // Determine if cluster is circular or rectangular (simplified for XML data)
                bool isCircular = false; // Default to rectangular for XML data
                
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
        /// Save cluster flags to XML immediately after pre-calculation
        /// This ensures the flags are available for the clustering logic
        /// </summary>
        private void SaveClusterFlagsToXml(string xmlFilePath, string targetCategory)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterService] Saving MarkedForClusteringSleeveProcess flags to XML: {Path.GetFileName(xmlFilePath)}");
                
                // Load current XML
                var serializer = new XmlSerializer(typeof(OpeningFilter));
                OpeningFilter filter;
                using (var reader = new StreamReader(xmlFilePath))
                {
                    filter = (OpeningFilter)serializer.Deserialize(reader);
                }
                
                if (filter?.ClashZoneStorage?.ClashZones == null)
                {
                    DebugLogger.Warning($"[UniversalClusterService] No clash zones found in XML for flag saving");
                    return;
                }
                
                // Update cluster flags from cache AND preserve IsCurrentClash flags
                int updatedCount = 0;
                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                {
                    if (_clashZoneCache != null && _clashZoneCache.ContainsKey(clashZone.MepElementId.IntegerValue))
                    {
                        var cachedClashZone = _clashZoneCache[clashZone.MepElementId.IntegerValue];
                        
                        // Update MarkedForClusteringSleeveProcess flag
                        if (cachedClashZone.MarkedForClusteringSleeveProcess != null)
                        {
                            clashZone.MarkedForClusteringSleeveProcess = cachedClashZone.MarkedForClusteringSleeveProcess;
                            updatedCount++;
                            
                            DebugLogger.Info($"[UniversalClusterService] Updated MarkedForClusteringSleeveProcess flag for ClashZone {clashZone.Id}: {clashZone.MarkedForClusteringSleeveProcess}");
                        }
                        
                        // ✅ CRITICAL: Preserve IsCurrentClash flag from cache
                        clashZone.IsCurrentClash = cachedClashZone.IsCurrentClash;
                        
                        DebugLogger.Info($"[UniversalClusterService] Preserved IsCurrentClash flag for ClashZone {clashZone.Id}: {clashZone.IsCurrentClash}");
                    }
                }
                
                // Save updated XML
                using (var writer = new StreamWriter(xmlFilePath))
                {
                    serializer.Serialize(writer, filter);
                }
                
                DebugLogger.Info($"[UniversalClusterService] Saved {updatedCount} MarkedForClusteringSleeveProcess flags and preserved IsCurrentClash flags to XML");
                DebugLogger.Info($"[DEBUG] SAVED FLAGS: {updatedCount} MarkedForClusteringSleeveProcess flags saved, IsCurrentClash flags preserved\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Error saving cluster flags: {ex.Message}");
                DebugLogger.Info($"[DEBUG] ERROR SAVING FLAGS: {ex.Message}\n");
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

        /// <summary>
        /// CRITICAL FIX: Set metadata parameters for cluster sleeve
        /// This ensures parameter transfer can find cluster sleeves
        /// </summary>
        private void SetClusterSleeveMetadata(FamilyInstance clusterSleeve, string category)
    {
        try
        {
            DebugLogger.Info($"[SetClusterSleeveMetadata] Setting metadata for cluster sleeve {clusterSleeve.Id}, category='{category}'");
            DebugLogger.Info($"[SetClusterSleeveMetadata] Setting metadata for cluster sleeve {clusterSleeve.Id}, category='{category}'\n");
            
            // ✅ OPTIONAL: Try to set MEP_Category parameter if it exists (not critical - XML has category info per sleeve)
            // Note: Opening families don't have MEP_Category parameter - XML stores category per clash zone
            // Parameter transfer uses MEP Element ID and XML category to find correct XML file
            var mepCategoryParam = clusterSleeve.LookupParameter("MEP_Category");
            if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
            {
                // Parameter exists and is writable - set it for convenience
                mepCategoryParam.Set(category);
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set MEP_Category = '{category}' for cluster sleeve {clusterSleeve.Id}");
                DebugLogger.Info($"[SetClusterSleeveMetadata] ✓ Set MEP_Category = '{category}' for cluster sleeve {clusterSleeve.Id}\n");
            }
            // ✅ NOTE: If parameter doesn't exist, that's OK - XML lookup will use MEP Element ID and category from XML
            
            // Set Filter Name based on category
            string filterName = GetFilterNameForCategory(category);
            var filterNameParam = clusterSleeve.LookupParameter("Filter Name");
            if (filterNameParam != null && !filterNameParam.IsReadOnly)
            {
                filterNameParam.Set(filterName);
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Filter Name = '{filterName}' for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Filter Name parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
            }
            
            // Set Sleeve Instance ID to -1 (indicating this is a cluster sleeve)
            var instanceIdParam = clusterSleeve.LookupParameter("Sleeve Instance ID");
            if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
            {
                instanceIdParam.Set(-1); // -1 indicates this is a cluster sleeve
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Sleeve Instance ID = -1 for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Sleeve Instance ID parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
            }
            
            // CRITICAL: Set Cluster Sleeve Instance ID parameter for XML lookup
            var clusterInstanceIdParam = clusterSleeve.LookupParameter("Cluster Sleeve Instance ID");
            if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
            {
                clusterInstanceIdParam.Set(clusterSleeve.Id.IntegerValue);
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Cluster Sleeve Instance ID = {clusterSleeve.Id.IntegerValue} for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Cluster Sleeve Instance ID parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
            }
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[SetClusterSleeveMetadata] Error setting metadata for cluster sleeve {clusterSleeve.Id}: {ex.Message}");
        }
    }

    /// <summary>
    /// ✅ NEW: Clean up individual sleeves that fall within cluster sleeve bounding boxes
    /// This is a cheap method that checks if any remaining individual sleeves are positioned
    /// within the bounding box of any cluster sleeve and deletes them
    /// </summary>
    private int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> placedClusters)
    {
        int deletedCount = 0;
        
        try
        {
            DebugLogger.Log($"[CleanupSleevesWithinClusters] 🔥 METHOD CALLED - placedClusters count: {placedClusters?.Count ?? 0}");
            DebugLogger.Info($"[CLEANUP-START] Method called with {placedClusters?.Count ?? 0} placed cluster sleeves\n");
            
            if (placedClusters == null || placedClusters.Count == 0)
            {
                DebugLogger.Log("[CleanupSleevesWithinClusters] No cluster sleeves to check against");
                DebugLogger.Info($"[CLEANUP-START] ⚠️ No cluster sleeves provided, exiting\n");
                return 0;
            }
            
            // Log cluster sleeve IDs for debugging
            if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                $"[CLEANUP-START] Cluster sleeve IDs: {string.Join(", ", placedClusters.Select(c => c.Id.IntegerValue))}\n");
            
            // Get all remaining individual sleeves (not cluster sleeves)
            // ✅ FIX: Use multiple strategies to identify sleeve families
            var allSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(s => 
                {
                    // Check if family name contains common sleeve keywords
                    bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                           s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                    
                    // Also check for specific family names used in the project
                    string familyName = s.Symbol?.FamilyName ?? "";
                    bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                        familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                    
                    return (s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                           (hasSleeveKeyword || isKnownFamily);
                })
                .ToList();
            
            // ✅ DEBUG: Log all sleeves found in Revit
            if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                $"[CLEANUP-ALL-SLEEVES-FOUND] Found {allSleeves.Count} total sleeves in Revit: {string.Join(", ", allSleeves.Select(s => s.Id.IntegerValue))}\n");
            
            var individualSleeves = allSleeves.Where(s => 
            {
                var clusterParam = s.LookupParameter("Cluster Sleeve Instance ID");
                // ✅ FIX: -1 means it's an individual sleeve, 0 or positive means it's a cluster sleeve
                // Only check sleeves with exactly -1 for cleanup
                int clusterValue = clusterParam?.AsInteger() ?? -1;
                
                // ✅ ADDITIONAL LOGIC: Also check if sleeve is a cluster by examining its Sleeve Instance ID parameter
                var sleeveInstanceIdParam = s.LookupParameter("Sleeve Instance ID");
                int sleeveInstanceId = sleeveInstanceIdParam?.AsInteger() ?? -1;
                
                // ✅ LOGIC: 
                // - If Sleeve Instance ID = -1, it's a CLUSTER sleeve (should NOT be deleted)
                // - If Cluster Sleeve Instance ID = -1 or 0, it's an INDIVIDUAL sleeve (can be deleted)
                // - We skip cluster sleeves by checking both parameters
                bool isClusterSleeve = (sleeveInstanceId == -1);
                bool isIndividual = !isClusterSleeve && (clusterValue <= 0);
                
                // ✅ DEBUG: Log ALL sleeves to find missing ones
                DebugLogger.Info($"[CLEANUP-DEBUG-ALL] Sleeve {s.Id}: Family={s.Symbol?.FamilyName}, ClusterParam={clusterValue}, SleeveInstanceId={sleeveInstanceId}, IsCluster={isClusterSleeve}, IsIndividual={isIndividual}\n");
                
                return isIndividual;
            }).ToList();
            
            DebugLogger.Log($"[CleanupSleevesWithinClusters] Found {allSleeves.Count} total sleeves, {individualSleeves.Count} individual sleeves to check against {placedClusters.Count} cluster sleeves");
            DebugLogger.Info($"[CLEANUP] Found {allSleeves.Count} total sleeves, checking {individualSleeves.Count} individual sleeves against {placedClusters.Count} cluster sleeves\n");
            
            // ✅ BATCH DELETION: Collect all sleeves to delete first, then delete in one batch
            var sleevesToDelete = new List<(ElementId sleeveId, int clusterId, Models.ClashZone clashZone)>();
            
            DebugLogger.Info($"[CLEANUP-BATCH] Scanning {individualSleeves.Count} individual sleeves for deletion candidates\n");
            
            foreach (var individualSleeve in individualSleeves)
            {
                DebugLogger.Info($"[CLEANUP-LOOP] Processing individual sleeve {individualSleeve.Id}\n");
                
                // ✅ FIX: Skip if this sleeve is actually a cluster sleeve in our placedClusters list
                if (placedClusters.Any(c => c.Id.IntegerValue == individualSleeve.Id.IntegerValue))
                {
                    DebugLogger.Info($"[CLEANUP] Skipping sleeve {individualSleeve.Id} - it's a cluster sleeve\n");
                    continue;
                }
                
                // ✅ CRITICAL FIX: Get bounding box from Revit element (works for all sleeves, even newly placed)
                int individualId = individualSleeve.Id.IntegerValue;
                var bbox = individualSleeve.get_BoundingBox(null);
                if (bbox == null)
                {
                    DebugLogger.Info($"[CLEANUP] Individual sleeve {individualId} has no bounding box, skipping\n");
                    continue;
                }
                var individualMin = bbox.Min;
                var individualMax = bbox.Max;
                
                // Try to find clash zone in cache for flag updates later
                var individualClashZone = _clashZoneCache.Values.FirstOrDefault(cz => 
                    cz.SleeveInstanceId == individualId || cz.AfterClusterSleevePlacedSleeveInstanceId == individualId);
                
                // ✅ OPTIMIZATION: Skip sleeves that were already deleted in Stage 1
                if (individualClashZone != null && individualClashZone.SleeveInstanceId == -1)
                {
                    DebugLogger.Info($"[CLEANUP-SKIP] Individual sleeve {individualId} already deleted in Stage 1 (SleeveInstanceId=-1), skipping\n");
                    continue;
                }
                
                // ✅ CRITICAL FIX: Get individual sleeve's host type to filter matching cluster sleeves
                var individualHostTypeParam = individualSleeve.LookupParameter("Host Type");
                string individualHostType = individualHostTypeParam?.AsString() ?? "";
                
                bool markedForDeletion = false;
                int deletingClusterId = -1;
                
                foreach (var clusterSleeve in placedClusters)
                {
                    // ✅ OPTIMIZED: Get cluster sleeve ID first
                    int clusterId = clusterSleeve.Id.IntegerValue;                       
                    
                    // ✅ CRITICAL FIX: Only compare against cluster sleeves with matching host type
                    var clusterHostTypeParam = clusterSleeve.LookupParameter("Host Type");
                    string clusterHostType = clusterHostTypeParam?.AsString() ?? "";
                    
                    if (!string.IsNullOrEmpty(individualHostType) && !string.IsNullOrEmpty(clusterHostType) && 
                        individualHostType != clusterHostType)
                    {
                        // Skip comparison if host types don't match (Floor vs Wall, etc.)
                        DebugLogger.Info($"[CLEANUP-SKIP] Individual sleeve {individualSleeve.Id} (hostType={individualHostType}) skipped cluster {clusterId} (hostType={clusterHostType})\n");
                        continue;
                    }
                    
                    // ✅ OPTIMIZED: Get cluster sleeve bounding box from XML cache (no expensive Revit API calls)                    
                    // Try to find cluster sleeve in cache (after XML save, this will have bounding boxes)
                    var clusterClashZone = _clashZoneCache.Values.FirstOrDefault(cz => cz.ClusterSleeveInstanceId == clusterId);
                    if (clusterClashZone == null || 
                        (clusterClashZone.ClusterSleeveBoundingBoxMinX == 0 && clusterClashZone.ClusterSleeveBoundingBoxMinY == 0))
                    {
                        DebugLogger.Info($"[CLEANUP] Cluster sleeve {clusterId} not found in cache or no bbox data (MinX={clusterClashZone?.ClusterSleeveBoundingBoxMinX ?? 0}), skipping\n");
                        continue;
                    }
                    
                    // ✅ Use bounding box from XML cache
                    var clusterMin = new XYZ(clusterClashZone.ClusterSleeveBoundingBoxMinX, clusterClashZone.ClusterSleeveBoundingBoxMinY, clusterClashZone.ClusterSleeveBoundingBoxMinZ);
                    var clusterMax = new XYZ(clusterClashZone.ClusterSleeveBoundingBoxMaxX, clusterClashZone.ClusterSleeveBoundingBoxMaxY, clusterClashZone.ClusterSleeveBoundingBoxMaxZ);
                    
                    DebugLogger.Info($"[CLEANUP] Cluster sleeve {clusterId} bbox from CACHE: Min=({clusterClashZone.ClusterSleeveBoundingBoxMinX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMinY:F6}), Max=({clusterClashZone.ClusterSleeveBoundingBoxMaxX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMaxY:F6})\n");
                    
                    // ✅ ENHANCED LOGGING: Show bounding boxes before checking (with Z coordinates for floors)
                    DebugLogger.Info($"[CLEANUP-CHECK] Individual sleeve {individualId}: Min=({individualMin.X:F3}, {individualMin.Y:F3}, {individualMin.Z:F3}), Max=({individualMax.X:F3}, {individualMax.Y:F3}, {individualMax.Z:F3})\n");
                    DebugLogger.Info($"[CLEANUP-CHECK] Cluster sleeve {clusterId}: Min=({clusterMin.X:F3}, {clusterMin.Y:F3}, {clusterMin.Z:F3}), Max=({clusterMax.X:F3}, {clusterMax.Y:F3}, {clusterMax.Z:F3})\n");
                    
                    // Check if individual sleeve is completely within cluster sleeve bounding box
                    bool withinBounds = IsWithinBounds(individualMin, individualMax, clusterMin, clusterMax);
                    
                    DebugLogger.Info($"[CLEANUP-CHECK] Individual {individualId} within Cluster {clusterId}: {withinBounds}\n");
                    
                    if (withinBounds)
                    {
                        DebugLogger.Log($"[CleanupSleevesWithinClusters] Individual sleeve {individualSleeve.Id} falls within cluster sleeve {clusterSleeve.Id} - marking for deletion");
                        DebugLogger.Info($"[CLEANUP-MARK] Individual sleeve {individualSleeve.Id} falls within cluster sleeve {clusterSleeve.Id} - marking for deletion\n");
                        
                        markedForDeletion = true;
                        deletingClusterId = clusterId;
                        break; // ✅ Found matching cluster, don't check remaining clusters
                    }
                }
                
                if (markedForDeletion)
                {
                    sleevesToDelete.Add((individualSleeve.Id, deletingClusterId, individualClashZone));
                }
            }
            
            // ✅ BATCH DELETE: Delete all marked sleeves in one transaction
            if (sleevesToDelete.Count > 0)
            {
                if (!DeploymentConfiguration.DeploymentMode) File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cluster_debug.log", 
                    $"[CLEANUP-BATCH-DELETE] Marked {sleevesToDelete.Count} sleeves for batch deletion: {string.Join(", ", sleevesToDelete.Select(s => s.sleeveId.IntegerValue))}\n");
                
                var elementIdsToDelete = sleevesToDelete.Select(s => s.sleeveId).ToList();
                
                try
                {
                    doc.Delete(elementIdsToDelete);
                    deletedCount = sleevesToDelete.Count;
                    DebugLogger.Info($"[CleanupSleevesWithinClusters] ✅ Batch deleted {deletedCount} individual sleeves");
                    
                    // ✅ CRITICAL: Update flags for all deleted sleeves (preserves IsResolved=true)
                    foreach (var (sleeveId, clusterId, clashZone) in sleevesToDelete)
                    {
                        if (clashZone != null)
                        {
                            clashZone.IsResolved = true;
                            clashZone.IsClusterResolved = true;
                            clashZone.ClusterSleeveInstanceId = clusterId;
                            clashZone.SleeveInstanceId = -1;
                        }
                        DebugLogger.Info($"[CLEANUP-FLAGS] Updated flags for {sleeveId.IntegerValue}: IsResolved=True, IsClusterResolved=True, ClusterSleeveId={clusterId}\n");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[CleanupSleevesWithinClusters] Error batch deleting sleeves: {ex.Message}");
                }
            }
            
            DebugLogger.Log($"[CleanupSleevesWithinClusters] Cleanup complete: {deletedCount} sleeves deleted");
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[CleanupSleevesWithinClusters] Error during cleanup: {ex.Message}");
        }
        
        return deletedCount;
    }
    
    /// <summary>
    /// Check if individual sleeve is more than 75% covered by cluster sleeve bounding box
    /// </summary>
    private bool IsWithinBounds(XYZ individualMin, XYZ individualMax, XYZ clusterMin, XYZ clusterMax)
    {
        // ✅ CO7ERAGE-BASED: Check if 75% or more of individual sleeve is within cluster bounds
        
        // Calculate overlap region (intersection of both bounding boxes)
        double overlapMinX = Math.Max(individualMin.X, clusterMin.X);
        double overlapMaxX = Math.Min(individualMax.X, clusterMax.X);
        double overlapMinY = Math.Max(individualMin.Y, clusterMin.Y);
        double overlapMaxY = Math.Min(individualMax.Y, clusterMax.Y);
        double overlapMinZ = Math.Max(individualMin.Z, clusterMin.Z);
        double overlapMaxZ = Math.Min(individualMax.Z, clusterMax.Z);
        
        // Calculate volumes
        double individualVolume = (individualMax.X - individualMin.X) * 
                                   (individualMax.Y - individualMin.Y) * 
                                   (individualMax.Z - individualMin.Z);
        
        // Check if there's any overlap
        if (overlapMaxX <= overlapMinX || overlapMaxY <= overlapMinY || overlapMaxZ <= overlapMinZ)
            return false;
        
        double overlapVolume = (overlapMaxX - overlapMinX) * 
                               (overlapMaxY - overlapMinY) * 
                               (overlapMaxZ - overlapMinZ);
        
        // Calculate coverage percentage
        double coverage = (individualVolume > 0) ? (overlapVolume / individualVolume) * 100.0 : 0.0;
        
        // Log coverage for debugging
        DebugLogger.Info($"[CLEANUP-COVERAGE] Individual: W={individualMax.X - individualMin.X:F3}, H={individualMax.Y - individualMin.Y:F3}, D={individualMax.Z - individualMin.Z:F3}, Vol={individualVolume:F6}\n");
        DebugLogger.Info($"[CLEANUP-COVERAGE] Overlap: W={overlapMaxX - overlapMinX:F3}, H={overlapMaxY - overlapMinY:F3}, D={overlapMaxZ - overlapMinZ:F3}, Vol={overlapVolume:F6}, Coverage={coverage:F1}%\n");
        
        // Return true if coverage is 75% or more
        return coverage >= 75.0;
    }

    /// <summary>
    /// Get XML filename for category
    /// ✅ CRITICAL: Filter name MUST come from _filterName field - no fallback allowed
    /// </summary>
        private string GetFilterNameForCategory(string category)
        {
            // ✅ CRITICAL: Filter name MUST come from _filterName field - no fallback allowed
            // This is essential for correct XML updates, clustering, and flag management
            if (string.IsNullOrEmpty(_filterName))
            {
                var errorMsg = $"Filter name is required but was not provided. Cannot determine target XML file for category '{category}'.";
                DebugLogger.Error($"[GetFilterNameForCategory] {errorMsg}");
                throw new InvalidOperationException(errorMsg);
            }
            
            return _filterName;
        }
        
        /// <summary>
        /// Save updated flags from cache to XML file after Stage 2 cleanup
        /// </summary>
        private void SaveUpdatedFlagsToXml(string xmlFilePath)
        {
            try
            {
                if (string.IsNullOrEmpty(xmlFilePath) || !File.Exists(xmlFilePath))
                {
                    DebugLogger.Warning($"[SaveUpdatedFlagsToXml] XML file not found: {xmlFilePath}");
                    return;
                }
                
                // Load existing XML
                var serializer = new XmlSerializer(typeof(Models.OpeningFilter));
                Models.OpeningFilter filter;
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                }
                
                // Update clash zones from cache (only those that were modified)
                if (filter.ClashZoneStorage == null)
                {
                    filter.ClashZoneStorage = new Models.ClashZoneStorage();
                }
                
                // Update all clash zones from cache (those modified during cleanup)
                int updatedCount = 0;
                foreach (var xmlClashZone in filter.ClashZoneStorage.ClashZones)
                {
                    var cachedClashZone = _clashZoneCache.Values
                        .FirstOrDefault(cz => cz.Id == xmlClashZone.Id);
                    
                    if (cachedClashZone != null)
                    {
                        // Update flags from cache (preserve IsResolved=true after Stage 2 cleanup)
                        xmlClashZone.IsResolved = cachedClashZone.IsResolved;
                        xmlClashZone.IsClusterResolved = cachedClashZone.IsClusterResolved;
                        xmlClashZone.SleeveInstanceId = cachedClashZone.SleeveInstanceId;
                        xmlClashZone.ClusterSleeveInstanceId = cachedClashZone.ClusterSleeveInstanceId;
                        xmlClashZone.AfterClusterSleevePlacedSleeveInstanceId = cachedClashZone.AfterClusterSleevePlacedSleeveInstanceId;
                        updatedCount++;
                    }
                }
                
                filter.LastModified = DateTime.Now;
                
                // Save back to XML
                using (var writer = new StreamWriter(xmlFilePath))
                {
                    serializer.Serialize(writer, filter);
                }
                
                DebugLogger.Info($"[SaveUpdatedFlagsToXml] ✅ Saved {updatedCount} updated flag changes to {xmlFilePath}");
                DebugLogger.Info($"[SAVE-XML] Updated {updatedCount} clash zone flags after cleanup\n");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SaveUpdatedFlagsToXml] Error: {ex.Message}");
                DebugLogger.Info($"[SAVE-XML-ERROR] {ex.Message}\n");
            }
        }
    }
}

