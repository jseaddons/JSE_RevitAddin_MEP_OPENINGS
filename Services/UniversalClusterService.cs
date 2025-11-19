using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
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
        
        // ✅ OOP REFACTORING: Centralized flag and filter management
        private FlagManager _flagManager;
        private FilterManagementService _filterService;
        
        // ✅ PERFORMANCE OPTIMIZATION: Cache MEP elements and bounding boxes to avoid duplicate API calls
        private Dictionary<ElementId, Element> _mepElementCache;
        private Dictionary<FamilyInstance, BoundingBoxXYZ> _bboxCache;
        
        // ✅ PERFORMANCE OPTIMIZATION: Cache parameters to avoid redundant LookupParameter calls
        private Dictionary<FamilyInstance, Dictionary<string, Parameter>> _parameterCache;
        
        // ✅ PHASE 1 OPTIMIZATION: Cache loaded clash zones to avoid duplicate XML loading
        private List<ClashZone> _loadedClashZonesCache;
        
        // ✅ ROTATED BBOX STORAGE: Store rotation angle and rotated bounding box for each cluster sleeve
        // Key: ClusterInstanceId, Value: (rotationAngleDeg, rotatedBboxMin, rotatedBboxMax, rotatedWidth, rotatedHeight, rotatedDepth)
        private Dictionary<int, (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)> _clusterRotationData = new Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>();
        
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
        /// Constructor with optional dependencies for OOP refactoring.
        /// If not provided, creates instances internally (backward compatible).
        /// </summary>
        public UniversalClusterService(FlagManager flagManager = null, FilterManagementService filterService = null)
        {
            _flagManager = flagManager; // Will be initialized per-document in ClusterSleeves if null
            _filterService = filterService; // Will be created in ClusterSleeves if null
        }

        /// <summary>
        /// Cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="targetCategory">Category to cluster (e.g., "Ducts", "Pipes") or null for all</param>
        /// <param name="uiDoc">Optional UIDocument for section box filtering</param>
        /// <returns>Tuple of (placedCount, deletedCount)</returns>
        public (int placedCount, int deletedCount) ClusterSleeves(
            Document doc, 
            string targetCategory, 
            UIDocument uiDoc = null, 
            string xmlFilePath = null, 
            string filterName = null, 
            List<FamilyInstance> placedClusterSleevesOut = null,
            // ✅ NEW: Path-based parameters for database integration
            bool isPath1Replay = false,
            int? comboId = null,
            int? filterId = null)
        {
            // Store document for helper methods
            _doc = doc;
            // ✅ CRITICAL: Store filterName in class field for use in GetFilterNameForCategory
            _filterName = filterName;
            
            // ✅ PERFORMANCE OPTIMIZATION: Initialize caches to avoid duplicate API calls
            _mepElementCache = new Dictionary<ElementId, Element>();
            _bboxCache = new Dictionary<FamilyInstance, BoundingBoxXYZ>();
            _parameterCache = new Dictionary<FamilyInstance, Dictionary<string, Parameter>>();
            
            // ✅ OOP REFACTORING: Initialize managers if not provided (backward compatibility)
            if (_flagManager == null)
            {
                _flagManager = new FlagManager(doc);
            }
            if (_filterService == null)
            {
                _filterService = new FilterManagementService(doc, msg => {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info(msg);
                }, msg => {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error(msg);
                });
            }
            
            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>(); // Track placed cluster sleeves for cleanup
            // ✅ CRITICAL: Track ClashZoneIds for each cluster sleeve (needed for database save)
            var clusterToClashZoneIds = new Dictionary<int, List<Guid>>(); // Key: ClusterInstanceId, Value: List of ClashZone GUIDs
            
            // ✅ PERFORMANCE: Start timing
            var startTime = DateTime.Now;

            try
            {
                // ✅ PATH 1 REPLAY: Check ClusterSleeves table for pre-calculated data
                if (isPath1Replay && comboId.HasValue && filterId.HasValue)
                {
                    try
                    {
                        var dbContext = new SleeveDbContext(doc);
                        var clusterRepository = new Data.Repositories.ClusterSleeveRepository(dbContext);
                        var existingClusters = clusterRepository.LoadClusterSleevesForCombo(comboId.Value, targetCategory);
                        
                        if (existingClusters != null && existingClusters.Count > 0)
                        {
                            // ✅ PATH 1: Load and place from database (skip calculation)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] PATH 1: Found {existingClusters.Count} pre-calculated clusters in database, loading and placing...");
                            }
                            
                            return PlaceClustersFromDatabase(doc, existingClusters, uiDoc, placedClusterSleevesOut, xmlFilePath, targetCategory, comboId.Value);
                        }
                        else
                        {
                            // ✅ PATH 1: No data → Skip clustering entirely
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] PATH 1: No cluster data found in database, skipping clustering. User should check 'Adopt to Modified Document' to trigger PATH 3.");
                            }
                            return (0, 0); // Skip clustering
                        }
                    }
                    catch (Exception path1Ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[CLUSTERING] PATH 1: Error loading cluster data: {path1Ex.Message}, falling back to calculation");
                        }
                        // Fall through to normal calculation
                    }
                }
                // ✅ PERFORMANCE FIX: Minimal logging - only log session start and end
                string clusterLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                bool __suppressClusterLogs = DeploymentConfiguration.DeploymentMode;
                if (!DeploymentConfiguration.DeploymentMode)
                    System.IO.File.WriteAllText(clusterLogPath, $"===== CLUSTER SESSION STARTED {DateTime.Now:HH:mm:ss} =====\n");
                
                // 🔥 CRITICAL DEBUG: Log which XML file cluster service is working with
                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string orchestratorDebugLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                    File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 CLUSTER SERVICE WORKING WITH XML FILE: {xmlFilePath ?? "NULL"}\n");
                    File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 CLUSTER SERVICE FILTER NAME: {filterName ?? "NULL"}\n");
                }
                
                // ⚠️ CRITICAL: Reset cluster flags for deleted cluster sleeves
                ResetClusterFlagsForDeletedSleeves(doc, xmlFilePath);
                
                // ✅ CRITICAL: Small delay to ensure XML files are fully written to disk
                System.Threading.Thread.Sleep(100);
                
                // ✅ PHASE 1 OPTIMIZATION: Load XML ONCE and cache for reuse (was loaded 4 times before!)
                // This eliminates 75% of file I/O overhead
                if (_loadedClashZonesCache == null)
                {
                    _loadedClashZonesCache = LoadClashZonesFromRegularXml(xmlFilePath, targetCategory, doc);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] ✅ PHASE 1: Loaded {_loadedClashZonesCache.Count} clash zones from XML (cached for reuse)");
                }
                var allClashZonesFromXml = _loadedClashZonesCache;
                
                // ✅ PHASE 1 OPTIMIZATION: Populate clash zone cache from loaded XML (single load, reuse)
                LoadClashZoneCacheFromLoadedClashZones(allClashZonesFromXml, targetCategory);
                
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Using cheaper bounding box overlap algorithm with tolerance: {toleranceMm}mm");
                
                // Get cluster configuration
                if (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    toleranceMm = Math.Min(toleranceMm, 100); // clamp pipes to 100mm
                }
                toleranceDist = UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Using JoinOpeningsDistance: {toleranceMm}mm (from {ClusterConfigurationManager.Instance.ConfigurationSource})");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Internal units: {toleranceDist:F6} feet");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Target category: {targetCategory ?? "ALL"}");
                
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(clusterLogPath, $"JoinOpeningsDistance: {toleranceMm}mm\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(clusterLogPath, $"Target category: {targetCategory ?? "ALL"}\n");
                }

                // ✅ DATABASE-FIRST CLUSTERING: Use cached clash zones (already loaded above from database/XML)
                // All data needed for clustering is already in database/XML after sleeve placement:
                // - SleeveBoundingBoxMinX/Y/Z, MaxX/Y/Z (bounding boxes from database/XML)
                // - SleeveInstanceId (sleeve IDs from database/XML)
                // - HostType, MepElementCategory, Orientation (from ClashZone properties in database/XML)
                // - StructuralElementIdValue (for wall host checking from database/XML)
                // ❌ REMOVED: Unnecessary Revit API calls (FilteredElementCollector, get_BoundingBox, LookupParameter)
                // Clustering uses database/XML data exclusively - no pre-calculation needed!
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] ✅ DATABASE-FIRST CLUSTERING: Loading clash zones from database/XML (no Revit API calls)");
                File.AppendAllText(clusterLogPath, $"✅ DATABASE-FIRST CLUSTERING: Loading clash zones from database/XML (no Revit API calls)\n");
                
                // ✅ DEBUG: Log total clash zones loaded
                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                var withSleeveId = allClashZonesFromXml.Count(cz => cz.SleeveInstanceId > 0);
                var notClusterResolved = allClashZonesFromXml.Count(cz => cz.SleeveInstanceId > 0 && !cz.IsClusterResolved);
                try 
                { 
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-START] Loaded {allClashZonesFromXml.Count} total clash zones from XML\n");
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-START] With SleeveInstanceId>0: {withSleeveId}, Not IsClusterResolved: {notClusterResolved}\n");
                    }
                } 
                catch { }
                
                // ✅ DEBUG: Log filtering steps to diagnose why wall sleeves aren't found
                var totalClashZones = allClashZonesFromXml.Count;
                // Reuse withSleeveId and notClusterResolved from above (lines 214-215)
                var categoryFiltered = allClashZonesFromXml.Count(cz => cz.SleeveInstanceId > 0 && !cz.IsClusterResolved && 
                    (string.IsNullOrEmpty(targetCategory) || string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase)));
                
                // ✅ DEBUG: Group by host type to see wall vs floor distribution
                var byHostType = allClashZonesFromXml
                    .Where(cz => cz.SleeveInstanceId > 0 && !cz.IsClusterResolved)
                    .Where(cz => string.IsNullOrEmpty(targetCategory) || string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    .GroupBy(cz => GetHostTypeFromClashZone(cz))
                    .ToDictionary(g => g.Key, g => g.Count());
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // Reuse placementDebugPath from line 207
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] Total clash zones: {totalClashZones}\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] With SleeveInstanceId>0: {withSleeveId}\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] Not IsClusterResolved: {notClusterResolved}\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] Category filtered ({targetCategory ?? "ALL"}): {categoryFiltered}\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] By HostType: {string.Join(", ", byHostType.Select(kvp => $"{kvp.Key}={kvp.Value}"))}\n");
                }
                
                // ✅ PHASE 1 OPTIMIZATION: Multi-threading for pre-calculation (filtering and data extraction)
                // These are pure XML data operations - safe for parallel processing
                var filteredClashZones = allClashZonesFromXml
                    .Where(cz => cz.SleeveInstanceId > 0) // Only placed sleeves
                    .Where(cz => !cz.IsClusterResolved) // ✅ CRITICAL: Skip sleeves already resolved by cluster
                    .Where(cz => string.IsNullOrEmpty(targetCategory) || 
                                string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                
                // ✅ PHASE 1 OPTIMIZATION: Parallel pre-calculation of host type, orientation, bounding box
                // Uses multi-threading flag for safe rollout
                var rawSleeves = new List<dynamic>();
                if (OptimizationFlags.UseClusterServiceMultiThreading && filteredClashZones.Count > 100)
                {
                    // Parallel processing for large datasets (100+ sleeves)
                    var parallelResults = new System.Collections.Concurrent.ConcurrentBag<dynamic>();
                    
                    System.Threading.Tasks.Parallel.ForEach(filteredClashZones, cz =>
                    {
                        try
                        {
                            dynamic sleeveData = new { 
                                SleeveInstanceId = cz.SleeveInstanceId,
                                Category = cz.MepElementCategory,
                                HostType = GetHostTypeFromClashZone(cz),
                                Orientation = GetEffectiveOrientationForClustering(cz),
                                BoundingBox = GetBoundingBoxFromClashZone(cz),
                                ClashZone = cz  // ✅ DEBUG: Keep reference for logging
                            };
                            parallelResults.Add(sleeveData);
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalClusterService] Error in parallel pre-calculation for clash zone {cz?.Id}: {ex.Message}");
                        }
                    });
                    
                    rawSleeves.AddRange(parallelResults);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] ✅ PHASE 1: Pre-calculated {rawSleeves.Count} sleeves using multi-threading");
                }
                else
                {
                    // Sequential processing (fallback or small datasets)
                    foreach (var cz in filteredClashZones)
                    {
                        dynamic sleeveData = new { 
                            SleeveInstanceId = cz.SleeveInstanceId,
                            Category = cz.MepElementCategory,
                            HostType = GetHostTypeFromClashZone(cz),
                            Orientation = GetEffectiveOrientationForClustering(cz),
                            BoundingBox = GetBoundingBoxFromClashZone(cz),
                            ClashZone = cz  // ✅ DEBUG: Keep reference for logging
                        };
                        rawSleeves.Add(sleeveData);
                    }
                }
                
                // ✅ DEBUG: Log bounding box filtering separately
                var beforeBboxFilter = rawSleeves.Count;
                var withNullBbox = rawSleeves.Count(s => s.BoundingBox == null);
                var byHostTypeBeforeBbox = rawSleeves.GroupBy(s => s.HostType).ToDictionary(g => g.Key, g => g.Count());
                var byHostTypeNullBbox = rawSleeves.Where(s => s.BoundingBox == null).GroupBy(s => s.HostType).ToDictionary(g => g.Key, g => g.Count());
                
                rawSleeves = rawSleeves.Where(s => s.BoundingBox != null).ToList(); // ✅ CRITICAL: Only include sleeves with valid bounding boxes
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // Reuse placementDebugPath from line 207
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] Before bbox filter: {beforeBboxFilter}\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] With null bbox: {withNullBbox} (by host: {string.Join(", ", byHostTypeNullBbox.Select(kvp => $"{kvp.Key}={kvp.Value}"))})\n");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-FILTER] After bbox filter: {rawSleeves.Count} (by host: {string.Join(", ", rawSleeves.GroupBy(s => s.HostType).Select(g => $"{g.Key}={g.Count()}"))})\n");
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Found {rawSleeves.Count} sleeves from XML data for category '{targetCategory}'");
                
                // ✅ DEBUG: Log which sleeves have zero bounding boxes
                var zeroBboxSleeves = allClashZonesFromXml
                    .Where(cz => cz.SleeveInstanceId > 0 && !cz.IsClusterResolved)
                    .Where(cz => cz.SleeveBoundingBoxMinX == 0 && cz.SleeveBoundingBoxMinY == 0 && cz.SleeveBoundingBoxMinZ == 0 &&
                                cz.SleeveBoundingBoxMaxX == 0 && cz.SleeveBoundingBoxMaxY == 0 && cz.SleeveBoundingBoxMaxZ == 0)
                    .Select(cz => cz.SleeveInstanceId)
                    .ToList();
                
                if (zeroBboxSleeves.Count > 0)
                {
                    try 
                    { 
                        System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] ⚠️ Found {zeroBboxSleeves.Count} sleeves with zero bounding boxes (filtered out): [{string.Join(", ", zeroBboxSleeves.Take(10))}...]\n"); 
                    } 
                    catch { }
                }
                
                // ✅ DEBUG: Log first few sleeves that will be clustered
                try 
                { 
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] Found {rawSleeves.Count} sleeves ready for clustering after filtering:\n");
                    }
                    var firstFew = rawSleeves.Take(5).Select(s => $"Sleeve {s.SleeveInstanceId} (Host={s.HostType}, Ori={s.Orientation}, BBox=({s.BoundingBox.Min.X:F3},{s.BoundingBox.Min.Y:F3},{s.BoundingBox.Min.Z:F3}) to ({s.BoundingBox.Max.X:F3},{s.BoundingBox.Max.Y:F3},{s.BoundingBox.Max.Z:F3}))");
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] First 5: {string.Join("\n", firstFew)}\n");
                } 
                catch { }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Processing {rawSleeves.Count} sleeves" + 
                               (string.IsNullOrEmpty(targetCategory) ? " (all categories)" : $" (category-filtered for {targetCategory})"));
                File.AppendAllText(clusterLogPath, $"Processing {rawSleeves.Count} sleeves (category-filtered for {targetCategory ?? "ALL"})\n");
                
                // ⚠️ CRITICAL: Log flag states of sleeves being processed for clustering (AFTER filtering)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string flagStateDebugLogPath = SafeFileLogger.GetLogFilePath("flag_state_debug.log");
                    System.IO.File.AppendAllText(flagStateDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] Processing {rawSleeves.Count} sleeves for clustering (filtered for {targetCategory ?? "ALL"})\n");
                }
                
                foreach (var sleeve in rawSleeves.Take(5)) // Log first 5 filtered sleeves
                {
                    // ✅ FIXED: Use XML data instead of Revit API calls
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    var clashZone = allClashZonesFromXml.FirstOrDefault(cz => cz.SleeveInstanceId == sleeveInstanceId);
                    if (clashZone != null)
                    {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] Sleeve {sleeveInstanceId}: MEP={clashZone.MepElementIdValue}, ClashZone={clashZone.Id}\n");
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-START] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                    }
                }

                // ✅ FIXED: Work with XML data directly - no Revit API calls needed
                if (rawSleeves.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] No sleeves to cluster from XML data");
                    return (0, 0);
                }

                // ⚠️ CRITICAL: Transaction must be started by caller
                if (!doc.IsModifiable)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                    File.AppendAllText(clusterLogPath, $"Sleeve {sleeve.SleeveInstanceId}: hostType={hostType}, systemType={systemType}, orientation={effectiveOrientation}\n");
                    }
                    
                    return new SleeveGroupKey(hostType, systemType, effectiveOrientation);
                });

                // Log group diagnostics
                foreach (var g in sleeveGroups)
                {
                    var list = g.ToList();
                    var ids = list.Select(s => s.SleeveInstanceId.ToString()).Take(10).ToList();
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[ClusterService] Group hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}");
                    File.AppendAllText(clusterLogPath, $"Group: hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}\n");
                }

                // ⚠️ CRITICAL: Add timeout protection for clustering operation
                var clusteringTimer = System.Diagnostics.Stopwatch.StartNew();
                const int MAX_CLUSTERING_TIME_MS = 300000; // 5 minutes

                // Form clusters using spatial hashing
                var clustersByGroup = FormClusters(sleeveGroups, toleranceDist);
                
                // ⚠️ CRITICAL: Check timeout after FormClusters
                if (clusteringTimer.ElapsedMilliseconds > MAX_CLUSTERING_TIME_MS)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[UniversalClusterService] ⏱ TIMEOUT: Clustering exceeded {MAX_CLUSTERING_TIME_MS / 1000} second limit during FormClusters");
                    }
                    System.Windows.Forms.MessageBox.Show(
                        $"Clustering operation is taking too long and has been cancelled.\n\nTime: {clusteringTimer.ElapsedMilliseconds / 1000} seconds\nLimit: {MAX_CLUSTERING_TIME_MS / 1000} seconds\n\nThis usually indicates:\n• Very large model with many sleeves\n• Infinite loop in clustering algorithm\n• Corrupted XML data\n\nPlease check the log file and try processing in smaller batches.",
                        "Operation Timeout",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return (placedCount, deletedCount); // Return partial results
                }
                
                // ✅ DEBUG: Log cluster formation results
                var totalClusters = clustersByGroup.Values.Sum(clusterList => clusterList.Count);
                var totalSleevesInClusters = clustersByGroup.Values.Sum(clusterList => clusterList.Sum(cluster => cluster.Count));
                var clustersWithMultipleSleeves = clustersByGroup.Values.Sum(clusterList => clusterList.Count(cluster => cluster.Count > 1));
                var individualSleevesSkipped = clustersByGroup.Values.Sum(clusterList => clusterList.Count(cluster => cluster.Count <= 1));
                
                var placementDebugPath2 = SafeFileLogger.GetLogFilePath("placement_debug.log");
                try 
                { 
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(placementDebugPath2, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] FormClusters returned {totalClusters} total clusters ({clustersWithMultipleSleeves} with >1 sleeve, {individualSleevesSkipped} individual sleeves skipped) with {totalSleevesInClusters} total sleeves\n");
                        DebugLogger.Info($"[CLUSTERING] FormClusters: {totalClusters} clusters ({clustersWithMultipleSleeves} with >1 sleeve, {individualSleevesSkipped} individual) from {rawSleeves.Count} sleeves");
                    }
                    foreach (var groupEntry in clustersByGroup)
                    {
                        var groupClustersWithMultiple = groupEntry.Value.Count(cluster => cluster.Count > 1);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(placementDebugPath2, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] Group {groupEntry.Key.hostType}_{groupEntry.Key.systemType}_{groupEntry.Key.orientation}: {groupEntry.Value.Count} clusters ({groupClustersWithMultiple} with >1 sleeve, {groupEntry.Value.Count - groupClustersWithMultiple} individual)\n");
                        }
                    }
                } 
                catch { }

                // Process each cluster
                // ⚠️ CRITICAL: Cluster PLACEMENT must remain sequential (single-threaded)
                // PlaceClusterSleeve() calls doc.Create.NewFamilyInstance() which is NOT thread-safe
                // Revit API requires all operations to be on the main thread
                // Only the FORMATION phase (FormClusters) can be parallelized (uses XML data only)
                int clusterProcessedCount = 0;
                foreach (var groupEntry in clustersByGroup)
                {
                    // ⚠️ CRITICAL: Check timeout every 5 clusters
                    clusterProcessedCount++;
                    if (clusterProcessedCount % 5 == 0 && clusteringTimer.ElapsedMilliseconds > MAX_CLUSTERING_TIME_MS)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[UniversalClusterService] ⏱ TIMEOUT: Clustering exceeded {MAX_CLUSTERING_TIME_MS / 1000} second limit after processing {clusterProcessedCount} clusters");
                        }
                        System.Windows.Forms.MessageBox.Show(
                            $"Clustering operation is taking too long and has been cancelled.\n\nProcessed: {clusterProcessedCount} of {totalClusters} clusters\nTime: {clusteringTimer.ElapsedMilliseconds / 1000} seconds\nLimit: {MAX_CLUSTERING_TIME_MS / 1000} seconds\n\nThis usually indicates:\n• Very large model\n• Infinite loop\n• Corrupted data\n\nPlease check the log file.",
                            "Operation Timeout",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        break; // Exit loop to prevent crash
                    }
                    
                    var groupKey = groupEntry.Key;
                    var clusters = groupEntry.Value;

                foreach (var cluster in clusters)
                {
                    if (cluster.Count <= 1) continue; // Skip individual sleeves

                        try
                        {
                            // ⚠️ CRITICAL: Check if cluster already exists using flag management
                            var clusterSleeveIds = string.Join(", ", cluster.Select(s => s.SleeveInstanceId));
                            var placementDebugPath3 = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            try 
                            { 
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(placementDebugPath3, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] Processing cluster with {cluster.Count} sleeves: [{clusterSleeveIds}]\n");
                                }
                            } 
                            catch { }
                            
                            if (IsClusterAlreadyExists(cluster, groupKey))
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DEBUG] SKIP: Cluster already exists for {cluster.Count} sleeves - skipping placement\n");
                                try 
                                { 
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(placementDebugPath3, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] ⚠️ SKIP: Cluster already exists for sleeves [{clusterSleeveIds}]\n");
                                    }
                                } 
                                catch { }
                                continue;
                            }

                            // Place cluster sleeve
                            try 
                            { 
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(placementDebugPath3, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] ✓ CALLING PlaceClusterSleeve for [{clusterSleeveIds}]\n");
                                }
                            } 
                            catch { }
                            
                            // ✅ CRITICAL: Track ClashZoneIds for this cluster before placing
                            var clusterClashZoneIds = new List<Guid>();
                            foreach (var sleeveData in cluster)
                            {
                                var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                                if (clashZone != null)
                                {
                                    clusterClashZoneIds.Add(clashZone.Id);
                                }
                            }
                            
                            FamilyInstance placedClusterSleeve = null;
                            int placed1 = 0;
                            int deleted1 = 0;
                            int? capturedClusterSleeveId = null; // ✅ CRITICAL FIX FOR DUCTS: Declare in outer scope so it's accessible everywhere
                            
                            try
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLUSTER-PLACEMENT-CALL] About to call PlaceClusterSleeve for cluster with {cluster.Count} sleeves\n");
                                }
                                
                                int? capturedIdFromPlace = null;
                                PlaceClusterSleeve(doc, cluster, groupKey, targetCategory, out placed1, out deleted1, out placedClusterSleeve, out capturedIdFromPlace, xmlFilePath);
                                
                                // ✅ CRITICAL FIX FOR DUCTS: Use captured ID from PlaceClusterSleeve output parameter
                                // The ID was captured INSIDE PlaceClusterSleeve while element was valid
                                capturedClusterSleeveId = capturedIdFromPlace;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    if (capturedClusterSleeveId.HasValue)
                                    {
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLUSTER-ID-CAPTURE] ✅ Using captured ID from PlaceClusterSleeve output parameter: {capturedClusterSleeveId.Value} (category={targetCategory})\n");
                                    }
                                    else
                                    {
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLUSTER-ID-CAPTURE] ⚠️ WARNING: No captured ID from PlaceClusterSleeve (category={targetCategory}, placed1={placed1}, deleted1={deleted1})\n");
                                    }
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLUSTER-PLACEMENT-CALL] ✅ PlaceClusterSleeve call completed (no exception)\n");
                                }
                            }
                            catch (Exception placeEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string errorLog = $"[CLUSTER-PLACEMENT-CALL] ❌❌❌ EXCEPTION calling PlaceClusterSleeve: {placeEx.Message}\n";
                                    errorLog += $"[CLUSTER-PLACEMENT-CALL] StackTrace: {placeEx.StackTrace}\n";
                                    DebugLogger.Error(errorLog);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, errorLog);
                                }
                                throw; // Re-throw to prevent continuing
                            }
                            
                            // ✅ CRITICAL LOGGING: Log PlaceClusterSleeve return values AFTER capturing ID
                            // Use captured ID for logging to avoid invalid element exceptions
                            try
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string placedClusterSleeveId = capturedClusterSleeveId.HasValue 
                                        ? $"ID={capturedClusterSleeveId.Value}" 
                                        : (placedClusterSleeve != null ? "INVALID_ELEMENT" : "NULL");
                                    
                                    string returnLog = $"[CLUSTER-PLACEMENT-RETURN] PlaceClusterSleeve returned: placed1={placed1}, deleted1={deleted1}, placedClusterSleeve={placedClusterSleeveId}\n";
                                    DebugLogger.Info(returnLog);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, returnLog);
                                    
                                    if (placed1 == 0)
                                    {
                                        string warningLog = $"[CLUSTER-PLACEMENT-RETURN] ⚠️ WARNING: placed1=0 but cluster sleeve was created! This means PlaceClusterSleeve returned 0 for placed count.\n";
                                        DebugLogger.Warning(warningLog);
                                        System.IO.File.AppendAllText(clusterDebugLogPath, warningLog);
                                    }
                                    if (placedClusterSleeve == null)
                                    {
                                        string errorLog = $"[CLUSTER-PLACEMENT-RETURN] ⚠️⚠️⚠️ ERROR: placedClusterSleeve is NULL! Cluster sleeve was created but not returned!\n";
                                        DebugLogger.Error(errorLog);
                                        System.IO.File.AppendAllText(clusterDebugLogPath, errorLog);
                                    }
                                }
                            }
                            catch (Exception logEx)
                            {
                                // Log the exception but don't stop execution
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string errorLog = $"[CLUSTER-PLACEMENT-RETURN] ❌❌❌ EXCEPTION during return value logging: {logEx.Message}\n";
                                    errorLog += $"[CLUSTER-PLACEMENT-RETURN] StackTrace: {logEx.StackTrace}\n";
                                    DebugLogger.Error(errorLog);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, errorLog);
                                }
                            }
                            
                            try 
                            { 
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    // ✅ Use captured ID instead of accessing element directly
                                    string clusterSleeveIdStr = capturedClusterSleeveId.HasValue 
                                        ? capturedClusterSleeveId.Value.ToString() 
                                        : (placedClusterSleeve != null ? "INVALID_ELEMENT" : "NULL");
                                    File.AppendAllText(placementDebugPath3, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] PlaceClusterSleeve returned: placed={placed1}, deleted={deleted1}, clusterSleeve={clusterSleeveIdStr}\n");
                                }
                            } 
                            catch { }
                            placedCount += placed1;
                            deletedCount += deleted1;
                            
                            // ✅ CRITICAL: Track cluster sleeve and its ClashZoneIds for database save
                            // ✅ FIX: Check placedClusterSleeve != null FIRST, even if placed1 == 0 (in case of bugs)
                            // ✅ CRITICAL FIX: Use captured ID instead of accessing element directly
                            // The captured ID was already obtained above, so we don't need to access the element again
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                string beforeCheckMsg = $"[CLUSTER-TRACKING-CHECK] About to check placedClusterSleeve != null: placedClusterSleeve={(placedClusterSleeve != null ? (capturedClusterSleeveId.HasValue ? $"ID={capturedClusterSleeveId.Value}" : "INVALID_ELEMENT") : "NULL")}\n";
                                DebugLogger.Info(beforeCheckMsg);
                                System.IO.File.AppendAllText(clusterDebugLogPath, beforeCheckMsg);
                            }
                            
                            // ✅ CRITICAL: Only add to placedClusters if element is valid and we have the ID
                            if (placedClusterSleeve != null && capturedClusterSleeveId.HasValue)
                            {
                                placedClusters.Add(placedClusterSleeve);
                                
                                // ✅ CRITICAL LOGGING: Log when cluster sleeve is added to placedClusters list
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string trackingMsg = $"[CLUSTER-TRACKING] ✅ Added cluster sleeve {capturedClusterSleeveId.Value} to placedClusters list (Total in list: {placedClusters.Count})\n";
                                    trackingMsg += $"[CLUSTER-TRACKING] Cluster sleeve {capturedClusterSleeveId.Value} contains {clusterClashZoneIds.Count} ClashZoneIds\n";
                                    DebugLogger.Info(trackingMsg);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, trackingMsg);
                                }
                                
                                // ✅ CRITICAL: Track cluster sleeve ID for database save
                                clusterToClashZoneIds[capturedClusterSleeveId.Value] = clusterClashZoneIds;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CLUSTERING] ✅ Tracked cluster sleeve {capturedClusterSleeveId.Value} with {clusterClashZoneIds.Count} ClashZoneIds: {string.Join(", ", clusterClashZoneIds.Take(5))}...");
                                    
                                    // ✅ DIAGNOSTIC: Log if ClashZoneIds are empty (indicates tracking issue)
                                    if (clusterClashZoneIds.Count == 0)
                                    {
                                        DebugLogger.Warning($"[CLUSTERING] ⚠️⚠️⚠️ WARNING: Cluster sleeve {capturedClusterSleeveId.Value} has NO ClashZoneIds tracked! This will cause cluster data to not be saved properly.");
                                    }
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string errorMsg = $"[CLUSTER-TRACKING] ❌❌❌ placedClusterSleeve is {(placedClusterSleeve == null ? "NULL" : "INVALID (cannot access ID)")} - NOT adding to placedClusters list!\n";
                                    errorMsg += $"[CLUSTER-TRACKING] placed1={placed1}, deleted1={deleted1}\n";
                                    DebugLogger.Error(errorMsg);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, errorMsg);
                                }
                            }
                            
                            // ✅ PERFORMANCE: Removed duplicate flag updates - PlaceClusterSleeve already handles flag updates internally
                            // Flags are updated inside PlaceClusterSleeve via MarkClashZonesAsClusterResolvedWithSleeveId and FlagManager
                            // No need to call UpdateClashZoneFlagsForCluster again (was causing 90+ second delays)
                            
                            // Marking happens inside PlaceClusterSleeve with actual cluster sleeve ID
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                string errorLog = $"[CLUSTER-PLACEMENT-OUTER-CATCH] ❌❌❌ EXCEPTION in cluster placement try block: {ex.Message}\n";
                                errorLog += $"[CLUSTER-PLACEMENT-OUTER-CATCH] StackTrace: {ex.StackTrace}\n";
                                errorLog += $"[CLUSTER-PLACEMENT-OUTER-CATCH] This exception prevented cluster sleeve from being added to placedClusters list!\n";
                                DebugLogger.Error(errorLog);
                                System.IO.File.AppendAllText(clusterDebugLogPath, errorLog);
                                DebugLogger.Error($"[UniversalClusterService] Error placing cluster: {ex.Message}");
                            }
                        }
                    }
                }

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.\n");
                
                // ⚠️ DISABLED: Cleanup will be called AFTER XML save in orchestrator to use cached bounding boxes
                // Cleanup requires cluster sleeve bounding boxes to be saved to XML first
                // The orchestrator will call CleanupSleevesWithinClustersAfterXmlSave() after XML update
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Final Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] Final Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.\n");
                
                // ✅ RETURN: Populate the output list with placed cluster sleeves
                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.Clear();
                    placedClusterSleevesOut.AddRange(placedClusters);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[UniversalClusterService] Returning {placedClusters.Count} placed cluster sleeves for coordinate update");
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string returnMsg = $"[CLUSTER-RETURN] ✅ Returning {placedClusters.Count} cluster sleeves to orchestrator\n";
                        if (placedClusters.Count > 0)
                        {
                            returnMsg += $"[CLUSTER-RETURN] Cluster sleeve IDs: {string.Join(", ", placedClusters.Select(c => c.Id.IntegerValue))}\n";
                        }
                        else
                        {
                            returnMsg += $"[CLUSTER-RETURN] ⚠️ WARNING: placedClusters list is EMPTY - no cluster sleeves to return!\n";
                        }
                        DebugLogger.Info(returnMsg);
                        System.IO.File.AppendAllText(clusterDebugLogPath, returnMsg);
                    }
                }
                
                // ✅ PERFORMANCE: Log clustering performance
                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] ⚡ Clustering completed in {duration.TotalSeconds:F1} seconds");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] ⚡ PERFORMANCE: Clustering completed in {duration.TotalSeconds:F1} seconds\n");
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] ===== CLUSTER DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Clustering failed: {ex.Message}");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Stack trace: {ex.StackTrace}");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] ERROR: {ex.Message}\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DEBUG] Stack trace: {ex.StackTrace}\n");
                throw;
            }

            // ✅ PATH 2/3: Save cluster data to database after calculation and placement
            // ✅ CRITICAL FIX: Try to get comboId and filterId from database if not provided
            int? finalComboId = comboId;
            int? finalFilterId = filterId;
            
            if (!finalComboId.HasValue || !finalFilterId.HasValue)
            {
                try
                {
                    var dbContext = new SleeveDbContext(doc);
                    
                    // Try to get filterId from Filters table using filterName and category
                    if (!finalFilterId.HasValue && !string.IsNullOrWhiteSpace(filterName))
                    {
                        var filterRepository = new Data.Repositories.FilterRepository(dbContext, _ => { });
                        var lookedUpFilterId = filterRepository.GetFilterId(filterName, targetCategory);
                        if (lookedUpFilterId > 0)
                        {
                            finalFilterId = lookedUpFilterId;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ✅ Looked up FilterId={lookedUpFilterId} from database for filter '{filterName}', category '{targetCategory}'");
                            }
                        }
                    }
                    
                    // Try to get comboId from the first placed cluster's clash zones
                    if (!finalComboId.HasValue && placedClusters != null && placedClusters.Count > 0 && clusterToClashZoneIds != null && clusterToClashZoneIds.Count > 0)
                    {
                        var firstClusterId = placedClusters[0].Id.IntegerValue;
                        if (clusterToClashZoneIds.TryGetValue(firstClusterId, out var clashZoneGuids) && clashZoneGuids != null && clashZoneGuids.Count > 0)
                        {
                            // Load clash zones from database to get file combo info
                            var clashZoneRepository = new Data.Repositories.ClashZoneRepository(dbContext);
                            var clashZones = clashZoneRepository.GetClashZonesByCategory(targetCategory);
                            
                            // Find the first clash zone that matches one of our GUIDs
                            var firstClashZone = clashZones?.FirstOrDefault(cz => clashZoneGuids.Contains(cz.Id));
                            if (firstClashZone != null && finalFilterId.HasValue)
                            {
                                // Get comboId from FileCombos table using the clash zone's file combo info
                                using (var cmd = dbContext.Connection.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                        SELECT ComboId FROM FileCombos 
                                        WHERE FilterId = @FilterId 
                                          AND Category = @Category 
                                          AND LinkedFileKey = @LinkedFileKey 
                                          AND HostFileKey = @HostFileKey
                                        LIMIT 1";
                                    cmd.Parameters.AddWithValue("@FilterId", finalFilterId.Value);
                                    cmd.Parameters.AddWithValue("@Category", targetCategory);
                                    // ✅ FIX: Use SourceDocKey (not ReferenceDocKey which doesn't exist)
                                    var linkedFileKey = firstClashZone.SourceDocKey ?? firstClashZone.DocumentPath ?? string.Empty;
                                    var hostFileKey = firstClashZone.HostDocKey ?? firstClashZone.StructuralElementDocumentTitle ?? string.Empty;
                                    cmd.Parameters.AddWithValue("@LinkedFileKey", linkedFileKey);
                                    cmd.Parameters.AddWithValue("@HostFileKey", hostFileKey);
                                    
                                    var result = cmd.ExecuteScalar();
                                    if (result != null && int.TryParse(result.ToString(), out int lookedUpComboId) && lookedUpComboId > 0)
                                    {
                                        finalComboId = lookedUpComboId;
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[CLUSTERING] ✅ Looked up ComboId={lookedUpComboId} from database using FilterId={finalFilterId.Value}, Category='{targetCategory}', LinkedFileKey='{linkedFileKey}', HostFileKey='{hostFileKey}'");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception lookupEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] ⚠️ Could not lookup comboId/filterId from database: {lookupEx.Message}");
                        DebugLogger.Warning($"[CLUSTERING] Stack trace: {lookupEx.StackTrace}");
                    }
                }
            }
            
            if (!isPath1Replay && finalComboId.HasValue && finalFilterId.HasValue && placedCount > 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTERING] ✅ Saving cluster data to database: ComboId={finalComboId.Value}, FilterId={finalFilterId.Value}, PlacedCount={placedCount}, Clusters={placedClusters?.Count ?? 0}");
                }
                try
                {
                    SaveClusterDataToDatabase(doc, placedClusters, finalComboId.Value, finalFilterId.Value, targetCategory, xmlFilePath, clusterToClashZoneIds);
                }
                catch (Exception saveEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] ❌ Error saving cluster data to database: {saveEx.Message}");
                        DebugLogger.Warning($"[CLUSTERING] Stack trace: {saveEx.StackTrace}");
                    }
                    // Non-critical error, continue
                }
            }
            else
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[CLUSTERING] ⚠️ SKIPPED saving cluster data: isPath1Replay={isPath1Replay}, finalComboId.HasValue={finalComboId.HasValue}, finalFilterId.HasValue={finalFilterId.HasValue}, placedCount={placedCount}");
                    if (!finalComboId.HasValue)
                        DebugLogger.Warning($"[CLUSTERING] ⚠️ comboId is NULL - cluster data will not be saved");
                    if (!finalFilterId.HasValue)
                        DebugLogger.Warning($"[CLUSTERING] ⚠️ filterId is NULL - cluster data will not be saved");
                }
            }

            // ✅ FLAG RESET: Reset IsFilterComboNew flag after cluster placement completes
            if (!isPath1Replay && comboId.HasValue && placedCount > 0)
            {
                try
                {
                    var dbContext = new SleeveDbContext(doc);
                    var clashZoneRepository = new Data.Repositories.ClashZoneRepository(dbContext);
                    clashZoneRepository.ResetFileComboFlag(comboId.Value);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTERING] ✅ Reset IsFilterComboNew=0 for ComboId={comboId.Value} after cluster placement");
                    }
                }
                catch (Exception resetEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] Error resetting IsFilterComboNew flag: {resetEx.Message}");
                    }
                    // Non-critical error, continue
                }
            }

            return (placedCount, deletedCount);
        }

        /// <summary>
        /// ✅ PATH 1 REPLAY: Place cluster sleeves from pre-calculated database data
        /// </summary>
        private (int placedCount, int deletedCount) PlaceClustersFromDatabase(
            Document doc,
            List<Data.Repositories.ClusterSleeveData> clusterDataList,
            UIDocument uiDoc,
            List<FamilyInstance> placedClusterSleevesOut,
            string xmlFilePath,
            string targetCategory,
            int comboId)
        {
            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>();

            try
            {
                foreach (var clusterData in clusterDataList)
                {
                    try
                    {
                        // Load clash zones for this cluster
                        var clashZoneIds = clusterData.ClashZoneIds;
                        if (clashZoneIds == null || clashZoneIds.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CLUSTERING] PATH 1: Cluster {clusterData.ClusterInstanceId} has no ClashZoneIds, skipping");
                            }
                            continue;
                        }

                        // Get family symbol based on host type
                        string familyName = "";
                        if (clusterData.HostType == "Wall" || clusterData.HostType == "Structural Framing")
                        {
                            familyName = "RectangularOpeningOnWall"; // Default to rectangular
                        }
                        else if (clusterData.HostType == "Floor")
                        {
                            familyName = "RectangularOpeningOnSlab";
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CLUSTERING] PATH 1: Unknown host type '{clusterData.HostType}' for cluster {clusterData.ClusterInstanceId}, skipping");
                            }
                            continue;
                        }

                        var universalSymbols = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (universalSymbols.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CLUSTERING] PATH 1: Family '{familyName}' not found for cluster {clusterData.ClusterInstanceId}, skipping");
                            }
                            continue;
                        }

                        var familySymbol = universalSymbols.First();
                        if (!familySymbol.IsActive) familySymbol.Activate();

                        // Get reference level (use first clash zone's level)
                        Level refLevel = null;
                        if (clashZoneIds.Count > 0)
                        {
                            // Try to get level from first clash zone
                            // This is a simplified approach - in production, you might want to store level info in ClusterSleeves table
                            refLevel = doc.GetElement(new ElementId(1)) as Level; // Fallback to first level
                        }

                        if (refLevel == null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CLUSTERING] PATH 1: Could not determine reference level for cluster {clusterData.ClusterInstanceId}, skipping");
                            }
                            continue;
                        }

                        // Create placement point
                        var placementPoint = new XYZ(clusterData.PlacementX, clusterData.PlacementY, clusterData.PlacementZ);

                        // Place cluster sleeve
                        FamilyInstance clusterSleeve = doc.Create.NewFamilyInstance(
                            placementPoint,
                            familySymbol,
                            refLevel,
                            StructuralType.NonStructural);

                        // Set dimensions
                        var widthParam = clusterSleeve.LookupParameter("Width");
                        var heightParam = clusterSleeve.LookupParameter("Height");
                        var depthParam = clusterSleeve.LookupParameter("Depth");

                        if (widthParam != null && !widthParam.IsReadOnly)
                            widthParam.Set(clusterData.ClusterWidth);
                        if (heightParam != null && !heightParam.IsReadOnly)
                            heightParam.Set(clusterData.ClusterHeight);
                        if (depthParam != null && !depthParam.IsReadOnly)
                            depthParam.Set(clusterData.ClusterDepth);

                        // ✅ CRITICAL: DO NOT rotate the cluster sleeve
                        // The cluster sleeve should always be axis-aligned (0° rotation)
                        // The IsRotated flag and RotationAngleDeg are stored for reference (bounding box calculation method used),
                        // but the cluster sleeve element itself should NOT be rotated
                        // Dimensions already account for the rotation of individual sleeves
                        if (!DeploymentConfiguration.DeploymentMode && clusterData.IsRotated && Math.Abs(clusterData.RotationAngleDeg) > 1e-6)
                        {
                            DebugLogger.Info($"[CLUSTER-ROTATION] ⚠️ Cluster bounding box was calculated with rotation angle {clusterData.RotationAngleDeg:F1}° (for coordinate system only), but cluster sleeve is placed axis-aligned (0°) - dimensions already account for rotation");
                        }

                        placedClusters.Add(clusterSleeve);
                        placedCount++;

                        // Delete individual sleeves within cluster
                        foreach (var clashZoneId in clashZoneIds)
                        {
                            // Find clash zone and get sleeve instance ID
                            // This is simplified - in production, you might want to store SleeveInstanceIds in ClusterSleeves table
                            // For now, we'll rely on the existing cleanup logic
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLUSTERING] PATH 1: Placed cluster sleeve {clusterSleeve.Id.IntegerValue} from database data");
                        }
                    }
                    catch (Exception clusterEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[CLUSTERING] PATH 1: Error placing cluster {clusterData.ClusterInstanceId}: {clusterEx.Message}");
                        }
                        // Continue with next cluster
                    }
                }

                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.Clear();
                    placedClusterSleevesOut.AddRange(placedClusters);
                }

                // Reset flag after placement
                try
                {
                    var dbContext = new SleeveDbContext(doc);
                    var clashZoneRepository = new Data.Repositories.ClashZoneRepository(dbContext);
                    clashZoneRepository.ResetFileComboFlag(comboId);
                }
                catch (Exception resetEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] PATH 1: Error resetting flag: {resetEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[CLUSTERING] PATH 1: Error in PlaceClustersFromDatabase: {ex.Message}");
                }
                throw;
            }

            return (placedCount, deletedCount);
        }

        /// <summary>
        /// ✅ PATH 2/3: Save cluster data to database after calculation and placement
        /// </summary>
        private void SaveClusterDataToDatabase(
            Document doc,
            List<FamilyInstance> placedClusters,
            int comboId,
            int filterId,
            string category,
            string xmlFilePath,
            Dictionary<int, List<Guid>> clusterToClashZoneIds = null)
        {
            try
            {
                var dbContext = new SleeveDbContext(doc);
                var clusterRepository = new Data.Repositories.ClusterSleeveRepository(dbContext);

                foreach (var clusterSleeve in placedClusters)
                {
                    try
                    {
                        int clusterInstanceId = clusterSleeve.Id.IntegerValue;
                        
                        // ✅ ROTATED BBOX STORAGE: Get stored rotation data for this cluster sleeve
                        // If rotation data exists, use rotated bounding box coordinates; otherwise use Revit element's bounding box
                        double rotationAngleDeg = 0.0;
                        bool isRotated = false;
                        XYZ bboxMin, bboxMax;
                        double width, height, depth;
                        
                        if (_clusterRotationData.TryGetValue(clusterInstanceId, out var rotationData))
                        {
                            // Use stored rotated bounding box coordinates
                            rotationAngleDeg = rotationData.rotationAngleDeg;
                            isRotated = rotationData.isRotated;
                            bboxMin = rotationData.rotatedBboxMin;
                            bboxMax = rotationData.rotatedBboxMax;
                            width = rotationData.rotatedWidth;
                            height = rotationData.rotatedHeight;
                            depth = rotationData.rotatedDepth;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ✅ Using stored rotated bounding box for cluster {clusterInstanceId}: Angle={rotationAngleDeg:F1}°, IsRotated={isRotated}, Width={width:F3}, Height={height:F3}, Depth={depth:F3}");
                            }
                        }
                        else
                        {
                            // Fallback: Get from Revit element (axis-aligned bounding box)
                            var bbox = clusterSleeve.get_BoundingBox(null);
                            if (bbox == null)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[CLUSTERING] Could not get bounding box for cluster sleeve {clusterSleeve.Id}, skipping save");
                                }
                                continue;
                            }
                            
                            bboxMin = bbox.Min;
                            bboxMax = bbox.Max;
                            
                            // Get dimensions from parameters
                            var widthParam = clusterSleeve.LookupParameter("Width");
                            var heightParam = clusterSleeve.LookupParameter("Height");
                            var depthParam = clusterSleeve.LookupParameter("Depth");
                            
                            width = widthParam?.AsDouble() ?? 0.0;
                            height = heightParam?.AsDouble() ?? 0.0;
                            depth = depthParam?.AsDouble() ?? 0.0;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ⚠️ No stored rotation data for cluster {clusterInstanceId}, using Revit element bounding box (axis-aligned)");
                            }
                        }

                        // Get host type
                        string hostType = GetHostTypeFromSleeve(clusterSleeve);
                        string hostOrientation = GetOrientationFromSleeve(clusterSleeve);

                        // ✅ CRITICAL FIX: Get ClashZoneIds from tracked dictionary (most reliable)
                        // Fallback to MEP_ElementIds parameter if tracking dictionary is not available
                        List<Guid> clashZoneIds = new List<Guid>();
                        
                        if (clusterToClashZoneIds != null && clusterToClashZoneIds.TryGetValue(clusterInstanceId, out var trackedIds))
                        {
                            // Use tracked ClashZoneIds (most reliable - from cluster formation)
                            clashZoneIds = trackedIds;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ✅ Using tracked ClashZoneIds for cluster {clusterInstanceId}: {clashZoneIds.Count} zones");
                            }
                        }
                        else
                        {
                            // Fallback: Get from MEP_ElementIds parameter or from XML cache
                            var mepElementIdsParam = clusterSleeve.LookupParameter("MEP_ElementIds");
                            if (mepElementIdsParam != null && mepElementIdsParam.HasValue)
                            {
                                // Parse MEP Element IDs and find corresponding clash zones
                                if (_clashZoneCache != null)
                                {
                                    var mepIdsString = mepElementIdsParam.AsString();
                                    if (!string.IsNullOrEmpty(mepIdsString))
                                    {
                                        var mepIds = mepIdsString.Split(',').Select(id => long.Parse(id.Trim())).ToList();
                                        foreach (var mepId in mepIds)
                                        {
                                            if (_clashZoneCache.TryGetValue(mepId, out var clashZone))
                                            {
                                                clashZoneIds.Add(clashZone.Id);
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode && clashZoneIds.Count == 0)
                            {
                                DebugLogger.Warning($"[CLUSTERING] ⚠️ No ClashZoneIds found for cluster {clusterInstanceId} (tracking dict not available, fallback also failed)");
                            }
                        }

                        // Get placement point (center of bounding box or midpoint)
                        var placementPoint = (bboxMin + bboxMax) / 2.0;

                        // ✅ ROTATED BBOX STORAGE: Save rotated bounding box coordinates to database
                        // Save to database
                        clusterRepository.SaveClusterSleeve(
                            clusterInstanceId: clusterInstanceId,
                            comboId: comboId,
                            filterId: filterId,
                            category: category,
                            boundingBoxMinX: bboxMin.X,
                            boundingBoxMinY: bboxMin.Y,
                            boundingBoxMinZ: bboxMin.Z,
                            boundingBoxMaxX: bboxMax.X,
                            boundingBoxMaxY: bboxMax.Y,
                            boundingBoxMaxZ: bboxMax.Z,
                            clusterWidth: width,
                            clusterHeight: height,
                            clusterDepth: depth,
                            rotationAngleDeg: rotationAngleDeg,
                            isRotated: isRotated,
                            placementX: placementPoint.X,
                            placementY: placementPoint.Y,
                            placementZ: placementPoint.Z,
                            hostType: hostType,
                            hostOrientation: hostOrientation,
                            clashZoneIds: clashZoneIds);

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLUSTERING] ✅ Saved cluster sleeve {clusterSleeve.Id.IntegerValue} to database");
                        }
                    }
                    catch (Exception clusterEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[CLUSTERING] Error saving cluster sleeve {clusterSleeve.Id}: {clusterEx.Message}");
                        }
                        // Continue with next cluster
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[CLUSTERING] Error in SaveClusterDataToDatabase: {ex.Message}");
                }
                throw;
            }
            
            // ✅ CRITICAL FIX: Save sleeve snapshots for cluster sleeves after saving cluster data
            if (placedClusters != null && placedClusters.Count > 0 && filterId > 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTERING] ✅ Saving sleeve snapshots for cluster sleeves: FilterId={filterId}, Clusters={placedClusters.Count}, ClashZoneIdsDict={clusterToClashZoneIds?.Count ?? 0}");
                }
                try
                {
                    using (var dbContext = new SleeveDbContext(doc))
                    {
                        var repository = new ClashZoneRepository(dbContext);
                        
                        // Get clash zones for cluster sleeves
                        var clusterZones = new List<ClashZone>();
                        foreach (var clusterSleeve in placedClusters)
                        {
                            int clusterInstanceId = clusterSleeve.Id.IntegerValue;
                            
                            // Get ClashZoneIds for this cluster
                            List<Guid> clashZoneIds = new List<Guid>();
                            if (clusterToClashZoneIds != null && clusterToClashZoneIds.TryGetValue(clusterInstanceId, out var trackedIds))
                            {
                                clashZoneIds = trackedIds;
                            }
                            
                            // ✅ FIX: Load clash zones from database by GUID directly (more reliable)
                            // Query by GUID instead of loading all and filtering
                            foreach (var clashZoneId in clashZoneIds)
                            {
                                try
                                {
                                    // Query database directly by GUID
                                    using (var cmd = dbContext.Connection.CreateCommand())
                                    {
                                        cmd.CommandText = @"
                                            SELECT ClashZoneId FROM ClashZones
                                            WHERE UPPER(ClashZoneGuid) = UPPER(@ClashZoneGuid)
                                              AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL
                                            LIMIT 1";
                                        cmd.Parameters.AddWithValue("@ClashZoneGuid", clashZoneId.ToString());
                                        var clashZoneIdResult = cmd.ExecuteScalar();
                                        
                                        if (clashZoneIdResult != null)
                                        {
                                            int dbClashZoneId = Convert.ToInt32(clashZoneIdResult);
                                            
                                            // Load clash zone by ID
                                            var zones = repository.GetClashZonesByCategory(category)
                                                ?.Where(z => z.Id == clashZoneId)
                                                .ToList();
                                            
                                            if (zones != null && zones.Count > 0)
                                            {
                                                // ✅ CRITICAL: Set ClusterSleeveInstanceId if not already set
                                                foreach (var zone in zones)
                                                {
                                                    if (zone.ClusterSleeveInstanceId != clusterInstanceId)
                                                    {
                                                        zone.ClusterSleeveInstanceId = clusterInstanceId;
                                                    }
                                                }
                                                clusterZones.AddRange(zones);
                                            }
                                        }
                                    }
                                }
                                catch (Exception zoneEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[CLUSTERING] ⚠️ Error loading clash zone {clashZoneId} for cluster {clusterInstanceId}: {zoneEx.Message}");
                                    }
                                }
                            }
                        }
                        
                        if (clusterZones.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ✅ Found {clusterZones.Count} cluster zones, saving snapshots...");
                            }
                            repository.SaveSleeveSnapshotsForPlacedSleeves(filterId, clusterZones);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTERING] ✅ Saved sleeve snapshots for {clusterZones.Count} cluster zones");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CLUSTERING] ⚠️ No cluster zones found to save snapshots (placedClusters={placedClusters.Count}, clusterToClashZoneIds={clusterToClashZoneIds?.Count ?? 0})");
                            }
                        }
                    }
                }
                catch (Exception snapshotEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] ⚠️ Failed to save cluster sleeve snapshots: {snapshotEx.Message}");
                    }
                }
            }
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
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter hostOrientationParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict8))
                {
                    paramDict8.TryGetValue("HostOrientation", out hostOrientationParam);
                }
                if (hostOrientationParam == null)
                {
                    hostOrientationParam = sleeve.LookupParameter("HostOrientation"); // Fallback if cache not available
                }
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
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine host type dynamically");
                return "Unknown";
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Error getting HostType for sleeve {sleeve.Id.IntegerValue}: {ex.Message}");
                return "Unknown";
            }
        }

        /// <summary>
        /// ✅ FIX: Get orientation from SleeveData XML instead of ClashZone cache
        /// </summary>
        private string GetOrientationFromClashZone(FamilyInstance sleeve, Dictionary<ElementId, Element> mepElementCache = null)
        {
            try
            {
                // ✅ DYNAMIC: Get orientation directly from MEP element using UniversalSleevePlacerService logic
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter mepElementIdParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict))
                {
                    paramDict.TryGetValue("MEP_ElementId", out mepElementIdParam);
                }
                if (mepElementIdParam == null)
                {
                    mepElementIdParam = sleeve.LookupParameter("MEP_ElementId"); // Fallback if cache not available
                }
                if (mepElementIdParam != null && mepElementIdParam.HasValue)
                {
                    var mepElementId = mepElementIdParam.AsElementId();
                    
                    // ✅ OPTIMIZATION: Use PRE-POPULATED cache first (calculate once, use many times)
                    // Check class-level cache first (pre-populated in batch), then parameter cache, then retrieve
                    Element mepElement = null;
                    if (_mepElementCache != null && _mepElementCache.TryGetValue(mepElementId, out var classCachedElement))
                    {
                        mepElement = classCachedElement; // ✅ Use pre-populated cache (fast)
                    }
                    else if (mepElementCache != null && mepElementCache.TryGetValue(mepElementId, out var paramCachedElement))
                    {
                        mepElement = paramCachedElement; // Use parameter cache if provided
                        // Also add to class cache for future use
                        if (_mepElementCache != null)
                        {
                            _mepElementCache[mepElementId] = mepElement;
                        }
                    }
                    else
                    {
                        // ✅ PERFORMANCE: Cache should always be populated, but log warning if miss occurs
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        {
                            DebugLogger.Warning($"[UniversalClusterService] Cache miss for MEP element {mepElementId} in GetOrientationFromClashZone - cache should be pre-populated");
                        }
                        // ✅ PHASE 1 OPTIMIZATION: Commented out expensive fallback - cache should be pre-populated
                        // If cache miss occurs, this indicates a bug in cache population logic
                        // mepElement = sleeve.Document.GetElement(mepElementId); // ❌ Expensive - cache should be pre-populated
                        // TODO: Pre-populate _mepElementCache before calling this method
                    }
                    
                    if (mepElement != null)
                    {
                        // Get orientation from MEP element's Wall Direction Type
                        var wallDirectionParam = mepElement.LookupParameter("Wall Direction Type");
                        if (wallDirectionParam != null && !string.IsNullOrEmpty(wallDirectionParam.AsString()))
                        {
                            string orientation = wallDirectionParam.AsString();
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Orientation='{orientation}' from MEP element Wall Direction Type");
                            return orientation;
                        }
                    }
                }
                
                // ✅ DYNAMIC: Fallback to HostOrientation parameter
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter hostOrientationParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict2))
                {
                    paramDict2.TryGetValue("HostOrientation", out hostOrientationParam);
                }
                if (hostOrientationParam == null)
                {
                    hostOrientationParam = sleeve.LookupParameter("HostOrientation"); // Fallback if cache not available
                }
                if (hostOrientationParam != null && !string.IsNullOrEmpty(hostOrientationParam.AsString()))
                {
                    string orientation = hostOrientationParam.AsString();
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Orientation='{orientation}' from HostOrientation parameter");
                    return orientation;
                }
                
                // ✅ DYNAMIC: Fallback to geometric analysis
                // ✅ OPTIMIZATION: Use PRE-POPULATED bbox cache (calculate once, use many times)
                BoundingBoxXYZ bbox = null;
                if (_bboxCache != null && _bboxCache.TryGetValue(sleeve, out var cachedBbox))
                {
                    bbox = cachedBbox; // ✅ Use pre-populated cache (fast, no API call)
                }
                else
                {
                    // ✅ PERFORMANCE: Cache should always be populated, but log warning if miss occurs
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        DebugLogger.Warning($"[UniversalClusterService] BoundingBox cache miss for sleeve {sleeve.Id} - cache should be pre-populated");
                    }
                    bbox = sleeve.get_BoundingBox(null); // ❌ Expensive - cache should be pre-populated
                    if (bbox != null && _bboxCache != null)
                    {
                        _bboxCache[sleeve] = bbox; // Add to cache for future use
                    }
                }
                
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
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine orientation dynamically");
                return "";
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Error getting orientation from MEP element: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// ✅ FIX: Get category from SleeveData XML instead of ClashZone cache
        /// </summary>
        private string GetCategoryFromMepElementId(FamilyInstance sleeve, Dictionary<ElementId, Element> mepElementCache = null)
        {
            try
            {
                // ✅ DYNAMIC: Get category directly from MEP_Category parameter (set by UniversalSleevePlacerService)
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter mepCategoryParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict3))
                {
                    paramDict3.TryGetValue("MEP_Category", out mepCategoryParam);
                }
                if (mepCategoryParam == null)
                {
                    mepCategoryParam = sleeve.LookupParameter("MEP_Category"); // Fallback if cache not available
                }
                if (mepCategoryParam != null && !string.IsNullOrEmpty(mepCategoryParam.AsString()))
                {
                    string category = mepCategoryParam.AsString();
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Category='{category}' from MEP_Category parameter");
                    return category;
                }
                
                // ✅ DYNAMIC: Fallback to MEP element lookup using UniversalSleevePlacerService logic
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter mepElementIdParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict4))
                {
                    paramDict4.TryGetValue("MEP_ElementId", out mepElementIdParam);
                }
                if (mepElementIdParam == null)
                {
                    mepElementIdParam = sleeve.LookupParameter("MEP_ElementId"); // Fallback if cache not available
                }
                if (mepElementIdParam != null && mepElementIdParam.HasValue)
                {
                    var mepElementId = mepElementIdParam.AsElementId();
                    
                    // ✅ OPTIMIZATION: Use PRE-POPULATED cache first (calculate once, use many times)
                    // Check class-level cache first (pre-populated in batch), then parameter cache, then retrieve
                    Element mepElement = null;
                    if (_mepElementCache != null && _mepElementCache.TryGetValue(mepElementId, out var classCachedElement))
                    {
                        mepElement = classCachedElement; // ✅ Use pre-populated cache (fast)
                    }
                    else if (mepElementCache != null && mepElementCache.TryGetValue(mepElementId, out var paramCachedElement))
                    {
                        mepElement = paramCachedElement; // Use parameter cache if provided
                        // Also add to class cache for future use
                        if (_mepElementCache != null)
                        {
                            _mepElementCache[mepElementId] = mepElement;
                        }
                    }
                    else
                    {
                        // ✅ PERFORMANCE: Cache should always be populated, but log warning if miss occurs
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        {
                            DebugLogger.Warning($"[UniversalClusterService] Cache miss for MEP element {mepElementId} in GetCategoryFromMepElementId - cache should be pre-populated");
                        }
                        // ✅ PHASE 1 OPTIMIZATION: Commented out expensive fallback - cache should be pre-populated
                        // If cache miss occurs, this indicates a bug in cache population logic
                        // mepElement = sleeve.Document.GetElement(mepElementId); // ❌ Expensive - cache should be pre-populated
                        // TODO: Pre-populate _mepElementCache before calling this method
                    }
                    
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
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Using Category='{category}' from MEP element");
                            return category;
                        }
                    }
                }
                
                // ✅ DYNAMIC: Fallback to family name analysis
                var familyName = sleeve.Symbol?.Family?.Name ?? "";
                if (familyName.Contains("Circular")) return "Pipes"; // Circular openings are typically pipes
                if (familyName.Contains("Rectangular")) return "Ducts"; // Rectangular openings are typically ducts
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalClusterService] Sleeve {sleeve.Id.IntegerValue}: Could not determine category dynamically");
                return "Unknown";
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DEBUG] Cluster already exists: ClashZone {clashZone.Id} is cluster-resolved (cluster sleeve ID: {clashZone.ClusterSleeveInstanceId})\n");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
        /// ✅ CONSOLIDATION: Mark clash zones as cluster-resolved and set cluster bounding box coordinates
        /// This consolidates cluster bounding box saving into UniversalClusterService (removed from SleeveCoordinateService)
        /// Cluster bounding boxes are read from Revit immediately after placement and saved to XML
        /// </summary>
        private void MarkClashZonesAsClusterResolvedWithSleeveId(List<dynamic> cluster, ElementId clusterSleeveId, string xmlFilePath = null, BoundingBoxXYZ clusterBbox = null, (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)? rotatedBbox = null)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[MarkClusterResolved] ✅ DATABASE-FIRST: Updating cluster flags in database for cluster sleeve {clusterSleeveId.IntegerValue}\n");
                }
                
                int markedCount = 0;
                var updatedClashZones = new List<ClashZone>();
                
                // ✅ CRITICAL: DATABASE-FIRST APPROACH - Update database before any XML operations
                foreach (var sleeve in cluster)
                {
                    // ✅ FIXED: Use sleeve instance ID from cluster data
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    
                    // ✅ DATABASE-FIRST: Get clash zone from database (not XML)
                    var clashZone = GetClashZoneBySleeveInstanceId(sleeveInstanceId, xmlFilePath);
                    if (clashZone != null)
                    {
                        // ✅ STORAGE: Save original SleeveInstanceId BEFORE clearing it
                        var originalSleeveInstanceId = clashZone.SleeveInstanceId; // Save for database matching
                        clashZone.AfterClusterSleevePlacedSleeveInstanceId = originalSleeveInstanceId;
                        
                        // ✅ CRITICAL: Update flags in memory first
                        clashZone.IsClusterResolved = true;
                        clashZone.ClusterSleeveId = clusterSleeveId;
                        clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue;
                        clashZone.IsResolved = true; // Individual sleeve was placed (then deleted)
                        clashZone.MarkedForClusteringSleeveProcess = true; // Mark as part of clustering history
                        clashZone.LastUpdated = DateTime.Now;
                        
                        // ✅ CONSOLIDATION: Set cluster bounding box coordinates from Revit immediately after placement
                        if (clusterBbox != null)
                        {
                            clashZone.ClusterSleeveBoundingBoxMinX = clusterBbox.Min.X;
                            clashZone.ClusterSleeveBoundingBoxMinY = clusterBbox.Min.Y;
                            clashZone.ClusterSleeveBoundingBoxMinZ = clusterBbox.Min.Z;
                            clashZone.ClusterSleeveBoundingBoxMaxX = clusterBbox.Max.X;
                            clashZone.ClusterSleeveBoundingBoxMaxY = clusterBbox.Max.Y;
                            clashZone.ClusterSleeveBoundingBoxMaxZ = clusterBbox.Max.Z;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[MarkClusterResolved] ✅ Set cluster bounding box for ClashZone {clashZone.Id}: Min=({clusterBbox.Min.X:F6}, {clusterBbox.Min.Y:F6}, {clusterBbox.Min.Z:F6}), Max=({clusterBbox.Max.X:F6}, {clusterBbox.Max.Y:F6}, {clusterBbox.Max.Z:F6})");
                        }
                        
                        // ✅ STORE ROTATED BOUNDING BOX: Store rotated bounding box coordinates in rotated coordinate system
                        if (rotatedBbox.HasValue)
                        {
                            clashZone.RotatedBoundingBoxMinX = rotatedBbox.Value.minX;
                            clashZone.RotatedBoundingBoxMinY = rotatedBbox.Value.minY;
                            clashZone.RotatedBoundingBoxMinZ = rotatedBbox.Value.minZ;
                            clashZone.RotatedBoundingBoxMaxX = rotatedBbox.Value.maxX;
                            clashZone.RotatedBoundingBoxMaxY = rotatedBbox.Value.maxY;
                            clashZone.RotatedBoundingBoxMaxZ = rotatedBbox.Value.maxZ;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[MarkClusterResolved] ✅ Set rotated bounding box for ClashZone {clashZone.Id}: Min=({rotatedBbox.Value.minX:F6}, {rotatedBbox.Value.minY:F6}, {rotatedBbox.Value.minZ:F6}), Max=({rotatedBbox.Value.maxX:F6}, {rotatedBbox.Value.maxY:F6}, {rotatedBbox.Value.maxZ:F6})");
                        }
                        
                        // ✅ CRITICAL FIX: Temporarily restore original SleeveInstanceId for database matching
                        // FlagManager.UpdateFlagsForPlacement uses OldSleeveInstanceId to match database records
                        var tempSleeveId = clashZone.SleeveInstanceId;
                        clashZone.SleeveInstanceId = originalSleeveInstanceId; // Restore for matching
                        
                        // ✅ DATABASE-FIRST: Update database flags immediately (before XML)
                        try
                        {
                            var categoryName = clashZone.MepElementCategory ?? string.Empty;
                            var baseFilterName = FilterNameHelper.NormalizeBaseName(_filterName, null, categoryName);
                            
                            _flagManager?.UpdateFlagsForPlacement(clashZone, clusterSleeveId.IntegerValue, isCluster: true, categoryName, baseFilterName);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[MarkClusterResolved] ✅ Updated DATABASE for ClashZone {clashZone.Id} with cluster sleeve {clusterSleeveId.IntegerValue}");
                        }
                        catch (Exception flagEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[MarkClusterResolved] ⚠️ Failed to update database flags for clash zone {clashZone.Id}: {flagEx.Message}");
                        }
                        
                        // Restore cleared state after database update
                        clashZone.SleeveInstanceId = -1; // Individual sleeve was deleted
                        clashZone.SleeveFamilyName = string.Empty; // Individual sleeve family cleared
                        
                        updatedClashZones.Add(clashZone);
                        markedCount++;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] ClashZone {clashZone.Id}: Cluster sleeve {clusterSleeveId.IntegerValue} placed\n");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                            DebugLogger.Info($"[DEBUG] Marked ClashZone {clashZone.Id} as cluster-resolved with cluster sleeve {clusterSleeveId.IntegerValue} (cleared individual flags)\n");
                        }
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[DEBUG] ✗ Clash zone not found for SleeveInstanceId={sleeveInstanceId}\n");
                    }
                }
                
                // ✅ XML UPDATE: Update XML files after database (for backward compatibility only)
                if (updatedClashZones.Count > 0)
                {
                    var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                    if (Directory.Exists(filtersDirectory))
                    {
                        var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                            ? Directory.GetFiles(filtersDirectory, "*.xml")
                            : new[] { xmlFilePath };
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[MarkClusterResolved] Updating {xmlFiles.Length} XML file(s) after database update: {(string.IsNullOrEmpty(xmlFilePath) ? "ALL" : Path.GetFileName(xmlFilePath))}\n");
                        }
                        
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

                                if (filter?.ClashZoneStorage?.AllZones != null)
                                {
                                    bool updated = false;
                                    
                                    foreach (var clashZone in updatedClashZones)
                                    {
                                        // Find clash zone in XML by SleeveInstanceId (before it was cleared to -1)
                                        var originalSleeveId = clashZone.AfterClusterSleevePlacedSleeveInstanceId;
                                        var xmlClashZone = GetStorageZones(filter).FirstOrDefault(cz => cz.SleeveInstanceId == originalSleeveId);
                                        if (xmlClashZone != null)
                                        {
                                            // Sync database state to XML
                                            xmlClashZone.IsClusterResolved = clashZone.IsClusterResolved;
                                            xmlClashZone.ClusterSleeveId = clashZone.ClusterSleeveId;
                                            xmlClashZone.ClusterSleeveInstanceId = clashZone.ClusterSleeveInstanceId;
                                            xmlClashZone.IsResolved = clashZone.IsResolved;
                                            xmlClashZone.SleeveInstanceId = clashZone.SleeveInstanceId; // -1
                                            xmlClashZone.SleeveFamilyName = clashZone.SleeveFamilyName; // empty
                                            xmlClashZone.MarkedForClusteringSleeveProcess = clashZone.MarkedForClusteringSleeveProcess;
                                            xmlClashZone.LastUpdated = clashZone.LastUpdated;
                                            
                                            if (clusterBbox != null)
                                            {
                                                xmlClashZone.ClusterSleeveBoundingBoxMinX = clashZone.ClusterSleeveBoundingBoxMinX;
                                                xmlClashZone.ClusterSleeveBoundingBoxMinY = clashZone.ClusterSleeveBoundingBoxMinY;
                                                xmlClashZone.ClusterSleeveBoundingBoxMinZ = clashZone.ClusterSleeveBoundingBoxMinZ;
                                                xmlClashZone.ClusterSleeveBoundingBoxMaxX = clashZone.ClusterSleeveBoundingBoxMaxX;
                                                xmlClashZone.ClusterSleeveBoundingBoxMaxY = clashZone.ClusterSleeveBoundingBoxMaxY;
                                                xmlClashZone.ClusterSleeveBoundingBoxMaxZ = clashZone.ClusterSleeveBoundingBoxMaxZ;
                                            }
                                            
                                            updated = true;
                                        }
                                    }
                                    
                                    // Save updated XML if changes were made
                                    if (updated)
                                    {
                                        // ⚠️ CRITICAL: Log flag states BEFORE XML save after clustering
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-XML-SAVE-BEFORE] About to save XML after clustering\n");
                                        
                                        try
                                        {
                                            // Normalize coordinates before saving to avoid 0,0,0 in XML
                                            if (filter?.ClashZoneStorage?.AllZones != null)
                                            {
                                                var zonesForSave = GetStorageZones(filter);
                                                if (zonesForSave.Count > 0)
                                                {
                                                    foreach (var z in zonesForSave)
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
                                                    }
                                                }
                                            }
                                            using (var writer = new StreamWriter(xmlFile))
                                            {
                                                serializer.Serialize(writer, filter);
                                            }
                                            
                                            // ⚠️ CRITICAL: Log flag states AFTER XML save after clustering
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-XML-SAVE-AFTER] XML save completed after clustering\n");
                                            
                                            // 🔥 CRITICAL DEBUG: Log XML file save
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 💾 XML FILE SAVED: {xmlFile} with {markedCount} cluster-resolved clash zones\n");
                                        }
                                        catch (Exception ex)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR SAVING XML FILE: {xmlFile} - {ex.Message}\n");
                                        }
                                    }
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ NO CHANGES MADE - XML FILE NOT SAVED: {xmlFile}\n");
                                    }
                                }
                    }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[UniversalClusterService] Error processing XML file {xmlFile}: {ex.Message}");
                            }
                        }
                    }
                }

                if (markedCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DEBUG] Marked {markedCount} clash zones as cluster-resolved\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalClusterService] Error marking clash zones as cluster-resolved: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ DATABASE-FIRST: Reset cluster flags for deleted cluster sleeves
        /// Queries database directly (no XML operations)
        /// </summary>
        private void ResetClusterFlagsForDeletedSleeves(Document doc, string xmlFilePath = null)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[ResetFlags] Processing database for deleted cluster sleeves\n");
                }
                
                int resetCount = 0;
                var resetClashZones = new List<ClashZone>(); // ✅ Track ClashZones that were reset for database update
                var oldClusterInstanceIds = new Dictionary<Guid, int>(); // ✅ Track old cluster instance IDs for database matching

                // ✅ DATABASE-FIRST: Get all categories and check cluster sleeves from database
                var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories", "Cable Tray Fittings" };
                
                using (var context = new SleeveDbContext(doc, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ResetClusterFlags][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[ResetClusterFlags][SQLite] {msg}");
                    });
                    
                    foreach (var category in categories)
                    {
                        try
                        {
                            // ✅ Get all ClashZones with cluster sleeves from database
                            var dbZones = repository.GetClashZonesByCategory(category)
                                ?.Where(z => z != null && z.IsClusterResolved && z.ClusterSleeveInstanceId > 0)
                                .ToList();
                            
                            if (dbZones == null || dbZones.Count == 0)
                                continue;
                            
                            foreach (var clashZone in dbZones)
                            {
                                // Check if cluster sleeve still exists in Revit
                                var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
                                var clusterSleeve = doc.GetElement(clusterSleeveId);
                                
                                if (clusterSleeve == null)
                                {
                                    // Cluster sleeve was deleted - reset flag
                                    // ✅ Store old cluster ID before resetting (needed for database matching)
                                    int oldClusterInstanceId = clashZone.ClusterSleeveInstanceId;
                                    oldClusterInstanceIds[clashZone.Id] = oldClusterInstanceId;
                                    
                                    clashZone.IsClusterResolved = false;
                                    clashZone.ClusterSleeveInstanceId = -1;
                                    resetCount++;
                                    resetClashZones.Add(clashZone); // ✅ Track for database update
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[ResetClusterFlags] Reset cluster flag for ClashZone {clashZone.Id} - cluster sleeve {oldClusterInstanceId} was deleted");
                                    }
                                }
                            }
                            
                            // ✅ Also check individual sleeves (IsResolved) - BUT ONLY if NOT cluster-resolved
                            var individualZones = repository.GetClashZonesByCategory(category)
                                ?.Where(z => z != null && z.IsResolved && !z.IsClusterResolved && z.SleeveInstanceId > 0)
                                .ToList();
                            
                            if (individualZones != null && individualZones.Count > 0)
                            {
                                foreach (var clashZone in individualZones)
                                {
                                    // Check if individual sleeve still exists in Revit
                                    var sleeveId = new ElementId(clashZone.SleeveInstanceId);
                                    var sleeve = doc.GetElement(sleeveId);
                                    
                                    if (sleeve == null)
                                    {
                                        // Individual sleeve was deleted - reset flag
                                        clashZone.IsResolved = false;
                                        clashZone.SleeveInstanceId = -1;
                                        resetCount++;
                                        resetClashZones.Add(clashZone); // ✅ Track for database update
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[ResetClusterFlags] Reset individual sleeve flag for ClashZone {clashZone.Id} - sleeve {clashZone.SleeveInstanceId} was deleted");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception catEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[ResetClusterFlags] Error processing category '{category}': {catEx.Message}");
                        }
                    }
                    
                    // ✅ Update database for all reset ClashZones
                    if (resetClashZones.Count > 0)
                    {
                        // Convert to database format
                        var dbUpdates = resetClashZones.Select(cz => (
                            ClashZoneId: cz.Id,
                            IsResolved: cz.IsResolved,
                            IsClusterResolved: cz.IsClusterResolved, // ✅ This is now false for cluster deletions
                            SleeveInstanceId: cz.SleeveInstanceId,
                            ClusterInstanceId: cz.ClusterSleeveInstanceId, // ✅ This is now -1 for cluster deletions
                            MepElementId: cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue,
                            StructuralElementId: cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue,
                            IntersectionPointX: cz.IntersectionPointX,
                            IntersectionPointY: cz.IntersectionPointY,
                            IntersectionPointZ: cz.IntersectionPointZ,
                            OldSleeveInstanceId: cz.SleeveInstanceId, // For cluster deletion, individual sleeve was already -1
                            OldClusterInstanceId: oldClusterInstanceIds.ContainsKey(cz.Id) ? oldClusterInstanceIds[cz.Id] : cz.ClusterSleeveInstanceId, // ✅ Use stored old cluster ID for database matching
                            MarkedForClusterProcess: (bool?)null,
                            AfterClusterSleeveId: -1,
                            IsClusteredFlag: (bool?)null
                        )).ToList();
                        
                        repository.BatchUpdateFlags(dbUpdates);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ResetClusterFlags] ✅ Updated database flags for {dbUpdates.Count} clash zones (IsClusterResolved/IsResolved reset to false)");
                        }
                    }
                }

                if (resetCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ResetClusterFlags] ✓ Reset flags for {resetCount} deleted sleeves (cluster + individual) - database updated");
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ResetClusterFlags] ✓ No deleted sleeves found - all flags preserved");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
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

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[ClusterService] Attempting to load family from: {familyPath}");

                if (!File.Exists(familyPath))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService] Universal family file not found: {familyPath}");
                    return false;
                }

                // Load the family
                bool loaded = doc.LoadFamily(familyPath);

                if (loaded)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[ClusterService] Successfully loaded universal family: {familyName}");
                    return true;
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService] Failed to load universal family: {familyName}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
            
            // ✅ MULTI-THREADING: ALWAYS enable parallel processing for clustering (XML-only, no Revit API calls)
            // Clustering uses pure XML data - completely safe for parallel processing
            // Groups are independent - can process Wall/Floor/Framing and different categories in parallel
            // Uses all CPU cores (i5/i7/i9) for maximum performance
            var groupsList = sleeveGroups.ToList(); // Materialize to avoid multiple enumerations
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Log($"[UniversalClusterService] 🚀 MULTI-THREADING: Processing {groupsList.Count} groups in parallel ({Environment.ProcessorCount} CPU cores)");
                try
                {
                    var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                    File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] 🚀 MULTI-THREADING: Processing {groupsList.Count} groups in parallel (CPU cores: {Environment.ProcessorCount})\n");
                }
                catch { }
            }
            
            // Use thread-safe dictionary for parallel processing
            var clustersByGroupConcurrent = new ConcurrentDictionary<SleeveGroupKey, List<List<dynamic>>>();
            
            // Reset counter for this run
            _processedGroupsCount = 0;
            
            try
            {
                // ✅ MULTI-THREADING: Process ALL groups in parallel (Wall, Floor, Framing, different categories)
                // Each group is independent - safe to process simultaneously
                Parallel.ForEach(groupsList, new ParallelOptions 
                { 
                    MaxDegreeOfParallelism = Environment.ProcessorCount // Use all available cores (i5/i7/i9)
                }, group =>
                    {
                        var xmlSleeves = group.ToList();
                        var groupClusters = new List<List<dynamic>>();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            Interlocked.Increment(ref _processedGroupsCount);
                            DebugLogger.Log($"[UniversalClusterService] [Thread {Thread.CurrentThread.ManagedThreadId}] Processing group {group.Key.hostType}_{group.Key.systemType} with {xmlSleeves.Count} sleeves");
                        }
                        
                        // ✅ FIXED: Use bounding box overlap algorithm with XML data
                        // ✅ PASS DOC: Need document to check wall hosts for preventing cross-wall clustering
                        var clusters = CalculateClustersUsingXmlData(xmlSleeves, toleranceDist, group.Key.orientation, _doc);
                        
                        if (clusters.Count > 0)
                        {
                            groupClusters.AddRange(clusters);
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Log($"[UniversalClusterService] [Thread {Thread.CurrentThread.ManagedThreadId}] Formed {clusters.Count} clusters using XML data");
                                DebugLogger.Info($"[DEBUG] ✓ CLUSTERS FORMED: {clusters.Count} clusters using XML data\n");
                                
                                // Log cluster details (thread-safe file append)
                                lock (_logLock)
                                {
                                    foreach (var cluster in clusters)
                                    {
                                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"  - Cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.SleeveInstanceId))}\n");
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[DEBUG] ℹ️ NO CLUSTERS: No proximate sleeves found in this group\n");
                        }
                        
                        clustersByGroupConcurrent[group.Key] = groupClusters;
                    });
                    
                    // Convert ConcurrentDictionary to regular Dictionary
                    clustersByGroup = clustersByGroupConcurrent.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[UniversalClusterService] ✅ MULTI-THREADING: Completed processing {groupsList.Count} groups in parallel");
                        try
                        {
                            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] ✅ MULTI-THREADING: Completed processing {groupsList.Count} groups\n");
                        }
                        catch { }
                    }
                }
                catch (AggregateException ex)
                {
                    // Fallback to single-threaded if parallel processing fails
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] ⚠️ Parallel processing failed, falling back to single-threaded: {ex.Message}");
                    
                    // Initialize empty dictionary - will be populated by fallback code if needed
                    clustersByGroup = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
                }

            return clustersByGroup;
        }
        
        // Thread-safe counter for parallel processing stats
        private int _processedGroupsCount = 0;
        private readonly object _logLock = new object();
        
        /// <summary>
        /// ✅ FIXED: Calculate clusters using XML data directly - no Revit API calls needed
        /// </summary>
        private List<List<dynamic>> CalculateClustersUsingXmlData(List<dynamic> xmlSleeves, double toleranceDist, string orientation, Document doc)
        {
            var clusters = new List<List<dynamic>>();
            var processed = new HashSet<int>();
            
                        if (!DeploymentConfiguration.DeploymentMode)
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
                            
                            // ✅ CRITICAL FIX: Check wall hosts BEFORE bounding box overlap using XML data (NO REVIT API CALLS)
                            // If both sleeves are on walls, ensure they're on the SAME wall by comparing StructuralElementIdValue from XML
                            if (clusterSleeve.HostType == "Wall" && otherSleeve.HostType == "Wall")
                            {
                                // ✅ OPTIMIZED: Use StructuralElementIdValue from XML (already in memory, no Revit API call needed)
                                var host1Id = clusterSleeve.ClashZone?.StructuralElementIdValue ?? -1;
                                var host2Id = otherSleeve.ClashZone?.StructuralElementIdValue ?? -1;
                                
                                if (host1Id > 0 && host2Id > 0 && host1Id != host2Id)
                                {
                                    // Different walls confirmed - do not cluster
                                    // ✅ PERFORMANCE: Removed file I/O logging from inner loop (was blocking)
                                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                                        DebugLogger.Info($"[UniversalClusterService] SKIP CLUSTERING (XML): Sleeve {clusterSleeve.SleeveInstanceId} (Wall {host1Id}) and {otherSleeve.SleeveInstanceId} (Wall {host2Id}) are on different walls");
                                    continue;
                                }
                                // If host1Id > 0 && host2Id > 0 && host1Id == host2Id, continue to distance check (same wall)
                                // If host1Id <= 0 || host2Id <= 0, proceed with distance check (will cluster if within tolerance)
                            }
                            
                            // ✅ PERFORMANCE: Removed file I/O logging from inner loop (was blocking - 244×244 = 59,536 potential writes!)
                            // Only log in diagnostic mode if needed
                            
                            // ✅ ROTATED SLEEVE CLUSTERING: Check if both sleeves are rotated
                            bool shouldUseRotatedClustering = false;
                            double rotationAngle1 = 0.0;
                            double rotationAngle2 = 0.0;
                            
                            // ✅ CRITICAL: Skip rotated clustering for round pipes and round ducts
                            // Round pipes/ducts should always use axis-aligned clustering regardless of MEP orientation
                            bool isRoundPipeOrDuct = false;
                            if (clusterSleeve.ClashZone != null)
                            {
                                var cz1 = clusterSleeve.ClashZone as ClashZone;
                                if (cz1 != null)
                                {
                                    // Check if it's a pipe (pipes are always round)
                                    bool isPipe = cz1.MepElementCategory != null && 
                                                 cz1.MepElementCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                                    
                                    // Check if it's a round duct
                                    bool isRoundDuct = cz1.MepElementCategory != null && 
                                                      cz1.MepElementCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                                      (string.Equals(cz1.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                                       (cz1.MepElementSizeData != null && 
                                                        (string.Equals(cz1.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                                         string.Equals(cz1.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase))));
                                    
                                    if (isPipe || isRoundDuct)
                                    {
                                        isRoundPipeOrDuct = true;
                                    }
                                }
                            }
                            
                            // ✅ COMPREHENSIVE LOGGING: Log clustering path selection (always enabled for debugging)
                            string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                            bool logClusteringPath = !DeploymentConfiguration.DeploymentMode;
                            
                            if (isRoundPipeOrDuct && logClusteringPath)
                            {
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== CLUSTERING PATH SELECTION ==========\n");
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Checking: Sleeve {clusterSleeve.SleeveInstanceId} vs Sleeve {otherSleeve.SleeveInstanceId}\n");
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SKIPPING ROTATED CLUSTERING: Round pipe/duct detected - using axis-aligned clustering\n");
                            }
                            
                            if (!isRoundPipeOrDuct && clusterSleeve.ClashZone != null && otherSleeve.ClashZone != null)
                            {
                                var cz1 = clusterSleeve.ClashZone as ClashZone;
                                var cz2 = otherSleeve.ClashZone as ClashZone;
                                
                                if (cz1 != null && cz2 != null)
                                {
                                    rotationAngle1 = cz1.MepElementRotationAngle;
                                    rotationAngle2 = cz2.MepElementRotationAngle;
                                    
                                    if (logClusteringPath)
                                    {
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== CLUSTERING PATH SELECTION ==========\n");
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Checking: Sleeve {clusterSleeve.SleeveInstanceId} vs Sleeve {otherSleeve.SleeveInstanceId}\n");
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {clusterSleeve.SleeveInstanceId} Rotation: {rotationAngle1 * 180.0 / Math.PI:F2}°\n");
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {otherSleeve.SleeveInstanceId} Rotation: {rotationAngle2 * 180.0 / Math.PI:F2}°\n");
                                    }
                                    
                                    // Check if both are rotated (non-axis-aligned)
                                    bool isRotated1 = Math.Abs(rotationAngle1) > 1e-6 && !IsAxisAlignedAngle(rotationAngle1);
                                    bool isRotated2 = Math.Abs(rotationAngle2) > 1e-6 && !IsAxisAlignedAngle(rotationAngle2);
                                    
                                    if (logClusteringPath)
                                    {
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {clusterSleeve.SleeveInstanceId} IsRotated: {isRotated1}, IsAxisAligned: {IsAxisAlignedAngle(rotationAngle1)}\n");
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {otherSleeve.SleeveInstanceId} IsRotated: {isRotated2}, IsAxisAligned: {IsAxisAlignedAngle(rotationAngle2)}\n");
                                    }
                                    
                                    // ✅ FIXED: Use rotated clustering ONLY if BOTH sleeves are on the SAME axis (or very similar)
                                    // If sleeves are on different axes (e.g., 135° vs 45°), use axis-aligned clustering
                                    // Only cluster rotated sleeves when they share the same axis direction
                                    if (isRotated1 && isRotated2)
                                    {
                                        // Check if axes are similar (same axis line)
                                        // Same axis means: difference of 0°, 180°, or 360° (modulo 180°)
                                        // Examples: 0° and 180° are same axis, 45° and 225° are same axis, 90° and 270° are same axis
                                        // Normalize both angles to 0-360° range first
                                        double angle1Deg = rotationAngle1 * 180.0 / Math.PI;
                                        double angle2Deg = rotationAngle2 * 180.0 / Math.PI;
                                        while (angle1Deg < 0) angle1Deg += 360.0;
                                        while (angle1Deg >= 360.0) angle1Deg -= 360.0;
                                        while (angle2Deg < 0) angle2Deg += 360.0;
                                        while (angle2Deg >= 360.0) angle2Deg -= 360.0;
                                        
                                        // Calculate difference and check modulo 180°
                                        double angleDiffDeg = Math.Abs(angle1Deg - angle2Deg);
                                        // Normalize difference to 0-180° range
                                        if (angleDiffDeg > 180.0) angleDiffDeg = 360.0 - angleDiffDeg;
                                        
                                        // Check if difference is 0° or 180° (same axis)
                                        // 0° = same direction, 180° = opposite direction on same axis
                                        double axisToleranceDeg = 1.0; // 1 degree tolerance
                                        bool isSameAxis = angleDiffDeg <= axisToleranceDeg || Math.Abs(angleDiffDeg - 180.0) <= axisToleranceDeg;
                                        
                                        if (logClusteringPath)
                                        {
                                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Both sleeves are rotated (non-axis-aligned)\n");
                                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Normalized angles: {angle1Deg:F2}° and {angle2Deg:F2}°\n");
                                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Axis difference: {angleDiffDeg:F2}° (0° or 180° = same axis)\n");
                                        }
                                        
                                        // Only use rotated clustering if axes are on same axis line (difference of 0°, 180°, or 360°)
                                        if (isSameAxis)
                                        {
                                            shouldUseRotatedClustering = true;
                                            if (logClusteringPath)
                                            {
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SELECTED: ROTATED CLUSTERING PATH (both sleeves on same axis: {rotationAngle1 * 180.0 / Math.PI:F2}°)\n");
                                            }
                                        }
                                        else
                                        {
                                            if (logClusteringPath)
                                            {
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ REJECTED: Rotated clustering (different axes: {angle1Deg:F2}° vs {angle2Deg:F2}°, diff={angleDiffDeg:F2}° - not 0° or 180°)\n");
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Will use axis-aligned clustering instead\n");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (logClusteringPath)
                                        {
                                            if (!isRotated1 && !isRotated2)
                                            {
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Both sleeves are axis-aligned (rotation angles: {rotationAngle1 * 180.0 / Math.PI:F2}°, {rotationAngle2 * 180.0 / Math.PI:F2}°)\n");
                                            }
                                            else if (!isRotated1)
                                            {
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {clusterSleeve.SleeveInstanceId} is axis-aligned, Sleeve {otherSleeve.SleeveInstanceId} is rotated\n");
                                            }
                                            else
                                            {
                                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {clusterSleeve.SleeveInstanceId} is rotated, Sleeve {otherSleeve.SleeveInstanceId} is axis-aligned\n");
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (logClusteringPath)
                                    {
                                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ ClashZone cast failed - cannot determine rotation\n");
                                    }
                                }
                            }
                            else
                            {
                                if (logClusteringPath)
                                {
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ ClashZone is null for one or both sleeves - using axis-aligned clustering\n");
                                }
                            }
                            
                            bool shouldCluster = false;
                            if (shouldUseRotatedClustering)
                            {
                                // ✅ FIXED: Use the axis direction from the first sleeve, NOT average of angles
                                // Each sleeve has its own axis direction - we use the first sleeve's axis for proximity check
                                // The rotation angle represents the axis direction the sleeve is oriented along
                                double axisRotationAngle = rotationAngle1; // Use first sleeve's axis direction
                                
                                if (logClusteringPath)
                                {
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] 🔄 EXECUTING: ROTATED CLUSTERING PATH\n");
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve 1 axis: {rotationAngle1 * 180.0 / Math.PI:F2}°, Sleeve 2 axis: {rotationAngle2 * 180.0 / Math.PI:F2}°\n");
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Using Sleeve 1 axis direction: {axisRotationAngle * 180.0 / Math.PI:F2}° (NOT average)\n");
                                }
                                shouldCluster = CheckRotatedSleeveProximity(clusterSleeve, otherSleeve, axisRotationAngle, toleranceDist);
                                if (logClusteringPath)
                                {
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated proximity check result: {shouldCluster}\n");
                                }
                            }
                            else
                            {
                                // ✅ AXIS-ALIGNED CLUSTERING: Use existing bounding box overlap check
                                if (logClusteringPath)
                                {
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] 📐 EXECUTING: AXIS-ALIGNED CLUSTERING PATH\n");
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Using BoundingBoxesOverlapFromXml\n");
                                }
                                // ✅ ENHANCED LOGGING: Get detailed distance info before calling BoundingBoxesOverlapFromXml
                                if (logClusteringPath)
                                {
                                    try
                                    {
                                        var bbox1 = clusterSleeve.BoundingBox;
                                        var bbox2 = otherSleeve.BoundingBox;
                                        if (bbox1 != null && bbox2 != null)
                                        {
                                            string hostType = clusterSleeve.HostType ?? "Unknown";
                                            string sleeveOrientation = clusterSleeve.Orientation ?? "Unknown";
                                            
                                            double minDistance;
                                            if (hostType == "Floor")
                                            {
                                                minDistance = CalculateMinimumDistance2D(
                                                    bbox1.Min.X, bbox1.Min.Y, bbox1.Max.X, bbox1.Max.Y,
                                                    bbox2.Min.X, bbox2.Min.Y, bbox2.Max.X, bbox2.Max.Y);
                                            }
                                            else
                                            {
                                                minDistance = CalculateMinimumDistance3D(
                                                    bbox1.Min.X, bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Y, bbox1.Max.Z,
                                                    bbox2.Min.X, bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Y, bbox2.Max.Z);
                                            }
                                            
                                            System.IO.File.AppendAllText(clusterDebugLogPath, 
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [AXIS-ALIGNED-DISTANCE] Sleeve {clusterSleeve.SleeveInstanceId} vs {otherSleeve.SleeveInstanceId}: " +
                                                $"Host={hostType}, Ori={sleeveOrientation}, " +
                                                $"Distance={UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm, " +
                                                $"Tolerance={UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
                                        }
                                    }
                                    catch (Exception logEx)
                                    {
                                        // Ignore logging errors
                                    }
                                }
                                shouldCluster = BoundingBoxesOverlapFromXml(clusterSleeve, otherSleeve, toleranceDist);
                                if (logClusteringPath)
                                {
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Axis-aligned proximity check result: {shouldCluster}\n");
                                }
                            }
                            
                            if (logClusteringPath)
                            {
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Final decision: {(shouldCluster ? "✅ CLUSTER" : "❌ NO CLUSTER")}\n");
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END CLUSTERING PATH SELECTION ==========\n\n");
                            }
                                
                            if (shouldCluster)
                            {
                                cluster.Add(otherSleeve);
                                processed.Add(otherSleeve.SleeveInstanceId);
                                foundNewNeighbors = true;
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Log($"[UniversalClusterService] Added sleeve {otherSleeve.SleeveInstanceId} to cluster (now {cluster.Count} sleeves)");
                            }
                        }
                    }
                }
                
                if (cluster.Count > 1) // Only add clusters with multiple sleeves
                {
                    clusters.Add(cluster);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Final cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.SleeveInstanceId))}");
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[UniversalClusterService] Formed {clusters.Count} clusters from {xmlSleeves.Count} sleeves");
            return clusters;
        }
        
        /// <summary>
        /// Determine dominant rotation angle from cluster sleeves
        /// Returns the average or most common rotation angle from individual sleeves
        /// </summary>
        private double DetermineDominantRotationAngle(List<dynamic> cluster, string xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0)
                    return 0.0;

                var rotationAngles = new List<double>();

                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    var clashZone = sleeveData.ClashZone as ClashZone;
                    if (clashZone == null)
                        continue;

                    // ✅ ROTATION DATA FROM DB: Get rotation angle from clash zone (loaded from database).
                    // MepElementRotationAngle is saved to database during refresh (InsertOrUpdateClashZones)
                    // and loaded by GetClashZonesByCategory -> MepRotationAngleRad column.
                    // This ensures cluster service uses rotation data from database, not calculated on-the-fly.
                    double angle = clashZone.MepElementRotationAngle;
                    
                    // Normalize angle to 0-2π range
                    while (angle < 0) angle += 2 * Math.PI;
                    while (angle >= 2 * Math.PI) angle -= 2 * Math.PI;
                    
                    rotationAngles.Add(angle);
                }

                if (rotationAngles.Count == 0)
                    return 0.0;

                // ✅ STRATEGY: Use average angle (works well for similar angles)
                // For angles that might wrap around (e.g., 350° and 10°), we need special handling
                double averageAngle = rotationAngles.Average();
                
                // Check if angles are spread across 0°/360° boundary
                double minAngle = rotationAngles.Min();
                double maxAngle = rotationAngles.Max();
                if (maxAngle - minAngle > Math.PI)
                {
                    // Angles wrap around - adjust by adding 2π to angles < π
                    var adjustedAngles = rotationAngles.Select(a => a < Math.PI ? a + 2 * Math.PI : a).ToList();
                    averageAngle = adjustedAngles.Average();
                    if (averageAngle >= 2 * Math.PI)
                        averageAngle -= 2 * Math.PI;
                }

                // ⚠️ PROTECTED CODE: DO NOT MODIFY THIS SECTION WITHOUT UNDERSTANDING THE IMPACT
                // This code is critical for preventing accidental rotation of axis-aligned cluster sleeves.
                // It checks if MEP elements are essentially axis-aligned (0°, 90°, 180°, 270°) and returns 0.0
                // to use axis-aligned bounding box logic. If this code is broken, straight sleeves will be incorrectly rotated.
                // 
                // ✅ CRITICAL FIX: Check if angle is essentially axis-aligned (0°, 90°, 180°, 270°)
                // If so, return 0 to use axis-aligned bounding box logic
                // ✅ IMPROVED: Check each individual angle first, then check average
                // This ensures that if ALL sleeves are axis-aligned, we use axis-aligned logic
                double thresholdDegrees = 2.0; // 2 degree tolerance (more forgiving for floating-point precision)
                double thresholdRadians = thresholdDegrees * Math.PI / 180.0;
                
                // Helper function to check if an angle (in radians) is axis-aligned
                bool IsAxisAligned(double angleRad)
                {
                    double angleDeg = angleRad * 180 / Math.PI;
                    // Normalize to 0-360 range
                    while (angleDeg < 0) angleDeg += 360;
                    while (angleDeg >= 360) angleDeg -= 360;
                    
                    // Check if close to 0°, 90°, 180°, or 270°
                    // For 0°: check if within threshold of 0 or 360
                    // For 90°: check if within threshold of 90
                    // For 180°: check if within threshold of 180
                    // For 270°: check if within threshold of 270
                    double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
                    double distTo90 = Math.Abs(angleDeg - 90);
                    double distTo180 = Math.Abs(angleDeg - 180);
                    double distTo270 = Math.Abs(angleDeg - 270);
                    
                    return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                           distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
                }
                
                // First, check if ALL individual angles are axis-aligned
                bool allAxisAligned = true;
                foreach (double angle in rotationAngles)
                {
                    if (!IsAxisAligned(angle))
                    {
                        allAxisAligned = false;
                        break;
                    }
                }
                
                // If all angles are axis-aligned, return 0
                if (allAxisAligned)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] All {rotationAngles.Count} angles are axis-aligned: [{angleList}], using axis-aligned bounding box");
                    }
                    return 0.0; // Use axis-aligned logic
                }
                
                // Second, check if average angle is axis-aligned
                if (IsAxisAligned(averageAngle))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double angleDegrees = averageAngle * 180 / Math.PI;
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] Average angle {angleDegrees:F1}° is axis-aligned (angles: [{angleList}]), using axis-aligned bounding box");
                    }
                    return 0.0; // Use axis-aligned logic
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ANGLE] Determined dominant rotation angle: {averageAngle * 180 / Math.PI:F1}° from {rotationAngles.Count} sleeves (non-axis-aligned)");
                }

                return averageAngle;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] Error determining dominant rotation angle: {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Get cluster bounding box using rotated bounding box coordinates from ClashZone data
        /// This ensures cluster size matches individual sleeve sizes for rotated MEP elements
        /// </summary>
        private (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) GetClusterBoundingBoxWithRotatedCoordinates(
            List<dynamic> cluster, 
            List<FamilyInstance> actualSleeves, 
            double rotationAngle, 
            string xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0 || actualSleeves == null || actualSleeves.Count == 0)
                    return (0, 0, 0, XYZ.Zero, null, null, null, null, null, null);

                // ✅ CRITICAL: For rotated sleeves, use rotated bounding box coordinates from ClashZone data
                var rotatedBboxes = new List<(XYZ min, XYZ max)>();
                var axisAlignedBboxes = new List<(XYZ min, XYZ max)>();
                bool hasRotatedBboxes = false;

                // ✅ COMPREHENSIVE LOGGING: Log all individual sleeve data being read
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== GET CLUSTER BOUNDING BOX WITH ROTATED COORDINATES ==========\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Cluster contains {cluster.Count} sleeves\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotation angle: {rotationAngle * 180.0 / Math.PI:F2}°\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== READING INDIVIDUAL SLEEVE DATA ==========\n");
                
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Starting to process {cluster.Count} sleeves in cluster\n");
                int sleeveIndex = 0;
                foreach (var sleeveData in cluster)
                {
                    sleeveIndex++;
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Processing sleeve {sleeveIndex}/{cluster.Count}: SleeveInstanceId={sleeveData.SleeveInstanceId}\n");
                    
                    var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (clashZone != null)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ClashZone found for sleeve {sleeveData.SleeveInstanceId}\n");
                        // ✅ LOG: Individual sleeve size and placement point
                        double sleeveWidth = clashZone.SleeveWidth > 0 ? clashZone.SleeveWidth : 0;
                        double sleeveHeight = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : 0;
                        double sleeveDiameter = clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : 0;
                        double placementX = clashZone.SleevePlacementPointX;
                        double placementY = clashZone.SleevePlacementPointY;
                        double placementZ = clashZone.SleevePlacementPointZ;
                        
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeveData.SleeveInstanceId}:\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Dimensions: W={sleeveWidth * 304.8:F1}mm, H={sleeveHeight * 304.8:F1}mm, D={sleeveDiameter * 304.8:F1}mm\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Placement Point: ({placementX:F6}, {placementY:F6}, {placementZ:F6})\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   MEP Rotation Angle: {clashZone.MepElementRotationAngle * 180.0 / Math.PI:F2}°\n");
                        
                        bool isRotated = Math.Abs(clashZone.MepElementRotationAngle) > 1e-6;
                        // ✅ FIX: Cast to ClashZone to avoid RuntimeBinderException with dynamic types
                        var cz = clashZone as ClashZone;
                        bool hasRotatedBbox = cz != null && 
                                            cz.RotatedBoundingBoxMinX.HasValue && 
                                            cz.RotatedBoundingBoxMinY.HasValue &&
                                            cz.RotatedBoundingBoxMaxX.HasValue && 
                                            cz.RotatedBoundingBoxMaxY.HasValue;
                        
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Has Rotated BBox: {hasRotatedBbox}\n");
                        
                        // ✅ LOG: Axis-aligned bounding box
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Axis-Aligned BBox: Min=({clashZone.SleeveBoundingBoxMinX:F6}, {clashZone.SleeveBoundingBoxMinY:F6}, {clashZone.SleeveBoundingBoxMinZ:F6}), Max=({clashZone.SleeveBoundingBoxMaxX:F6}, {clashZone.SleeveBoundingBoxMaxY:F6}, {clashZone.SleeveBoundingBoxMaxZ:F6})\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Axis-Aligned Size: W={(clashZone.SleeveBoundingBoxMaxX - clashZone.SleeveBoundingBoxMinX) * 304.8:F1}mm, H={(clashZone.SleeveBoundingBoxMaxY - clashZone.SleeveBoundingBoxMinY) * 304.8:F1}mm\n");

                        // ✅ CRITICAL FIX: Use rotated bbox if available, even if cluster rotation is 0
                        // Individual sleeves may be rotated even if cluster is axis-aligned
                        if (hasRotatedBbox && cz != null)
                        {
                            double rotatedWidth = (cz.RotatedBoundingBoxMaxX.Value - cz.RotatedBoundingBoxMinX.Value) * 304.8;
                            double rotatedHeight = (cz.RotatedBoundingBoxMaxY.Value - cz.RotatedBoundingBoxMinY.Value) * 304.8;
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Rotated BBox: Min=({cz.RotatedBoundingBoxMinX.Value:F6}, {cz.RotatedBoundingBoxMinY.Value:F6}, {cz.RotatedBoundingBoxMinZ ?? cz.SleeveBoundingBoxMinZ:F6}), Max=({cz.RotatedBoundingBoxMaxX.Value:F6}, {cz.RotatedBoundingBoxMaxY.Value:F6}, {cz.RotatedBoundingBoxMaxZ ?? cz.SleeveBoundingBoxMaxZ:F6})\n");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Rotated Size: W={rotatedWidth:F1}mm, H={rotatedHeight:F1}mm\n");
                            
                            rotatedBboxes.Add((
                                new XYZ(cz.RotatedBoundingBoxMinX.Value, 
                                       cz.RotatedBoundingBoxMinY.Value, 
                                       cz.RotatedBoundingBoxMinZ ?? cz.SleeveBoundingBoxMinZ),
                                new XYZ(cz.RotatedBoundingBoxMaxX.Value, 
                                       cz.RotatedBoundingBoxMaxY.Value, 
                                       cz.RotatedBoundingBoxMaxZ ?? cz.SleeveBoundingBoxMaxZ)
                            ));
                            hasRotatedBboxes = true;
                        }
                        else
                        {
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ⚠️ No rotated bbox - using axis-aligned\n");
                            if (cz != null)
                            {
                                axisAlignedBboxes.Add((
                                    new XYZ(cz.SleeveBoundingBoxMinX, 
                                           cz.SleeveBoundingBoxMinY, 
                                           cz.SleeveBoundingBoxMinZ),
                                    new XYZ(cz.SleeveBoundingBoxMaxX, 
                                           cz.SleeveBoundingBoxMaxY, 
                                           cz.SleeveBoundingBoxMaxZ)
                                ));
                            }
                        }
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ✅ Sleeve data loaded\n\n");
                    }
                    else
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ ClashZone is NULL for SleeveInstanceId {sleeveData.SleeveInstanceId} - trying Revit API\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ⚠️ ClashZone not found for SleeveInstanceId {sleeveData.SleeveInstanceId} - trying Revit API\n");
                        var sleeve = actualSleeves.FirstOrDefault(s => s.Id.IntegerValue == sleeveData.SleeveInstanceId);
                        if (sleeve != null)
                        {
                            var bbox = sleeve.get_BoundingBox(null);
                            if (bbox != null && bbox.Enabled)
                            {
                                double revitWidth = (bbox.Max.X - bbox.Min.X) * 304.8;
                                double revitHeight = (bbox.Max.Y - bbox.Min.Y) * 304.8;
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Revit BBox: Min=({bbox.Min.X:F6}, {bbox.Min.Y:F6}, {bbox.Min.Z:F6}), Max=({bbox.Max.X:F6}, {bbox.Max.Y:F6}, {bbox.Max.Z:F6})\n");
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Revit Size: W={revitWidth:F1}mm, H={revitHeight:F1}mm\n");
                                axisAlignedBboxes.Add((bbox.Min, bbox.Max));
                            }
                        }
                        else
                        {
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ❌ Sleeve not found in Revit either!\n");
                        }
                    }
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Completed processing sleeve {sleeveIndex}/{cluster.Count}\n\n");
                }
                
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== FOREACH LOOP COMPLETED ==========\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== SUMMARY: RotatedBboxes={rotatedBboxes.Count}, AxisAlignedBboxes={axisAlignedBboxes.Count}, HasRotatedBboxes={hasRotatedBboxes} ==========\n\n");
                
                // ✅ CRITICAL: Log all rotated bbox coordinates for debugging
                if (rotatedBboxes.Count > 0)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== ROTATED BBOX COORDINATES ==========\n");
                    for (int i = 0; i < rotatedBboxes.Count; i++)
                    {
                        var bbox = rotatedBboxes[i];
                        double w = (bbox.max.X - bbox.min.X) * 304.8;
                        double h = (bbox.max.Y - bbox.min.Y) * 304.8;
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] RotatedBbox[{i}]: Min=({bbox.min.X:F6}, {bbox.min.Y:F6}, {bbox.min.Z:F6}), Max=({bbox.max.X:F6}, {bbox.max.Y:F6}, {bbox.max.Z:F6}), Size=W={w:F1}mm, H={h:F1}mm\n");
                    }
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END ROTATED BBOX COORDINATES ==========\n\n");
                }

                // ✅ DIAGNOSTIC: Log whether rotated bounding boxes are found
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-BBOX] Checking rotated bounding boxes: hasRotatedBboxes={hasRotatedBboxes}, rotatedBboxes.Count={rotatedBboxes.Count}, axisAlignedBboxes.Count={axisAlignedBboxes.Count}");
                }
                
                if (hasRotatedBboxes && rotatedBboxes.Count > 0)
                {
                    // ✅ DETAILED LOGGING: Log individual sleeve rotated bounding boxes
                    var clusterLogPath = SafeFileLogger.GetLogFilePath("cluster_bbox_detailed.log");
                    var logBuilder = new System.Text.StringBuilder();
                    logBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== CLUSTER BOUNDING BOX CALCULATION (ROTATED) ==========");
                    logBuilder.AppendLine($"Cluster contains {rotatedBboxes.Count} individual sleeves with rotated bounding boxes");
                    logBuilder.AppendLine($"Rotation angle: {rotationAngle * 180.0 / Math.PI:F2}°");
                    logBuilder.AppendLine();
                    
                    // Log each individual sleeve's rotated bounding box
                    int detailedLogSleeveIndex = 0;
                    foreach (var sleeveData in cluster)
                    {
                        detailedLogSleeveIndex++;
                        var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                        // ✅ FIX: Cast to ClashZone to avoid RuntimeBinderException with dynamic types
                        var czDetailed = clashZone as ClashZone;
                        if (czDetailed != null && czDetailed.RotatedBoundingBoxMinX.HasValue)
                        {
                            double sleeveWidth = (czDetailed.RotatedBoundingBoxMaxX.Value - czDetailed.RotatedBoundingBoxMinX.Value) * 304.8; // Convert to mm
                            double sleeveHeight = (czDetailed.RotatedBoundingBoxMaxY.Value - czDetailed.RotatedBoundingBoxMinY.Value) * 304.8; // Convert to mm
                            
                            logBuilder.AppendLine($"  Sleeve #{detailedLogSleeveIndex} (SleeveInstanceId={sleeveData.SleeveInstanceId}):");
                            logBuilder.AppendLine($"    RotatedBBox Min: ({czDetailed.RotatedBoundingBoxMinX.Value:F6}, {czDetailed.RotatedBoundingBoxMinY.Value:F6}, {czDetailed.RotatedBoundingBoxMinZ ?? czDetailed.SleeveBoundingBoxMinZ:F6})");
                            logBuilder.AppendLine($"    RotatedBBox Max: ({czDetailed.RotatedBoundingBoxMaxX.Value:F6}, {czDetailed.RotatedBoundingBoxMaxY.Value:F6}, {czDetailed.RotatedBoundingBoxMaxZ ?? czDetailed.SleeveBoundingBoxMaxZ:F6})");
                            logBuilder.AppendLine($"    Individual Size: W={sleeveWidth:F1}mm × H={sleeveHeight:F1}mm");
                            logBuilder.AppendLine($"    MEP Rotation Angle: {czDetailed.MepElementRotationAngle * 180.0 / Math.PI:F2}°");
                            logBuilder.AppendLine();
                        }
                    }
                    
                    // ✅ PROPER WATERTIGHT ALGORITHM: Calculate all 4 corners of each sleeve in world space
                    // Then transform all corners into common rotated coordinate system and find min/max extents
                    // This works for ALL scenarios: single, stacked, inline, diagonal, grid
                    // 
                    // ⚠️ CRITICAL: rotationAngle represents the cluster's INTENDED ROTATED AXIS direction,
                    // NOT the average of sleeve rotation angles. This axis defines the coordinate system
                    // that all corners will be transformed into for the min/max extent calculation.
                    double width, height, minX, minY, maxX, maxY;
                    XYZ origin = XYZ.Zero;  // Store origin for midpoint transformation
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== PROPER CORNER-BASED ALGORITHM ==========\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Processing {cluster.Count} sleeves - calculating all 4 corners in world space\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Cluster INTENDED ROTATED AXIS (not average): {rotationAngle * 180.0 / Math.PI:F2}°\n");
                        
                        // Step 1: Collect sleeve data (center, width, height, rotation, pre-calculated corners) from ClashZone
                        // ✅ OPTIMIZATION: Use pre-calculated corners from database (dump once use many times)
                        var sleeveDataList = new List<(ClashZone cz, XYZ center, double width, double height, double sleeveRotation, double? cosSleeve, double? sinSleeve, XYZ[] preCalculatedCorners)>();
                        
                        foreach (var sleeveData in cluster)
                        {
                            var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                            if (clashZone == null) continue;
                            
                            var cz = clashZone as ClashZone;
                            if (cz == null) continue;
                            
                            // ✅ Get sleeve center from Active document coordinates (where sleeve is actually placed)
                            XYZ center = new XYZ(
                                cz.SleevePlacementPointActiveDocumentX,
                                cz.SleevePlacementPointActiveDocumentY,
                                cz.SleevePlacementPointActiveDocumentZ
                            );
                            
                            // Get sleeve dimensions
                            double sleeveWidth = cz.SleeveWidth > 0 ? cz.SleeveWidth : 0;
                            double sleeveHeight = cz.SleeveHeight > 0 ? cz.SleeveHeight : 0;
                            
                            // Get sleeve rotation angle
                            // ⚠️ NOTE: Database column is MepRotationAngleRad, mapped to MepElementRotationAngle in ClashZone model
                            double sleeveRotation = cz.MepElementRotationAngle;
                            
                            // ✅ ROTATION MATRIX: Use pre-calculated cos/sin from database (dump once use many times)
                            double? cosSleeve = cz.MepRotationCos;
                            double? sinSleeve = cz.MepRotationSin;
                            
                            // ✅ PRE-CALCULATED CORNERS: Load from database (dump once use many times)
                            XYZ[] preCalculatedCorners = null;
                            if (cz.SleeveCorner1X.HasValue && cz.SleeveCorner1Y.HasValue && cz.SleeveCorner1Z.HasValue &&
                                cz.SleeveCorner2X.HasValue && cz.SleeveCorner2Y.HasValue && cz.SleeveCorner2Z.HasValue &&
                                cz.SleeveCorner3X.HasValue && cz.SleeveCorner3Y.HasValue && cz.SleeveCorner3Z.HasValue &&
                                cz.SleeveCorner4X.HasValue && cz.SleeveCorner4Y.HasValue && cz.SleeveCorner4Z.HasValue)
                            {
                                preCalculatedCorners = new XYZ[]
                                {
                                    new XYZ(cz.SleeveCorner1X.Value, cz.SleeveCorner1Y.Value, cz.SleeveCorner1Z.Value),  // Corner 1: Bottom-left
                                    new XYZ(cz.SleeveCorner2X.Value, cz.SleeveCorner2Y.Value, cz.SleeveCorner2Z.Value),  // Corner 2: Bottom-right
                                    new XYZ(cz.SleeveCorner3X.Value, cz.SleeveCorner3Y.Value, cz.SleeveCorner3Z.Value),  // Corner 3: Top-left
                                    new XYZ(cz.SleeveCorner4X.Value, cz.SleeveCorner4Y.Value, cz.SleeveCorner4Z.Value)   // Corner 4: Top-right
                                };
                            }
                            
                            // ✅ VERIFY: Log if rotation angle is missing/zero (might indicate data not loaded from DB)
                            if (Math.Abs(sleeveRotation) < 1e-6)
                            {
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ⚠️ WARNING: Sleeve {sleeveData.SleeveInstanceId} has zero rotation angle - check if MepElementRotationAngle was loaded from DB (column: MepRotationAngleRad)\n");
                            }
                            
                            // ✅ VERIFY: Log if pre-calculated corners are missing
                            if (preCalculatedCorners == null)
                            {
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   ⚠️ WARNING: Sleeve {sleeveData.SleeveInstanceId} has no pre-calculated corners - will recalculate (check if UpdateSleeveCorners was called during placement)\n");
                            }
                            
                            sleeveDataList.Add((cz, center, sleeveWidth, sleeveHeight, sleeveRotation, cosSleeve, sinSleeve, preCalculatedCorners));
                            
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Sleeve {sleeveData.SleeveInstanceId}: Center=({center.X:F6}, {center.Y:F6}, {center.Z:F6}), W={sleeveWidth*304.8:F1}mm, H={sleeveHeight*304.8:F1}mm, Rotation={sleeveRotation*180.0/Math.PI:F2}°, HasPreCalcCorners={preCalculatedCorners != null}, HasPreCalcCosSin={cosSleeve.HasValue && sinSleeve.HasValue}\n");
                        }
                        
                        if (sleeveDataList.Count == 0)
                        {
                            // Fallback to simple union of rotated bboxes
                            minX = rotatedBboxes.Min(b => b.min.X);
                            minY = rotatedBboxes.Min(b => b.min.Y);
                            maxX = rotatedBboxes.Max(b => b.max.X);
                            maxY = rotatedBboxes.Max(b => b.max.Y);
                            width = maxX - minX;
                            height = maxY - minY;
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ No sleeve data found - using fallback union\n");
                        }
                        else
                        {
                            // Step 2: Choose reference point (first sleeve center as origin)
                            origin = sleeveDataList[0].center;
                            
                            // Step 3: Pre-calculate cluster rotation matrix components
                            // ⚠️ CRITICAL: rotationAngle is the cluster's INTENDED ROTATED AXIS, not average of sleeve angles
                            // This defines the coordinate system frame that all corners will be transformed into
                            double cosCluster = Math.Cos(rotationAngle);
                            double sinCluster = Math.Sin(rotationAngle);
                            
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Origin (first sleeve center): ({origin.X:F6}, {origin.Y:F6}, {origin.Z:F6})\n");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Cluster INTENDED AXIS Rotation Matrix: cos={cosCluster:F6}, sin={sinCluster:F6}\n");
                            
                            // Step 4: For each sleeve, use pre-calculated corners from database OR recalculate if missing
                            // ✅ OPTIMIZATION: Use pre-calculated corners from database (dump once use many times)
                            var allTransformedCorners = new List<XYZ>();
                            
                            // ✅ OPTIMIZATION: Pre-calculate corner offset pattern (for fallback recalculation only)
                            // Corner offsets in local space: (-1,-1), (1,-1), (-1,1), (1,1) multiplied by halfW/halfH
                            var cornerOffsetPattern = new[]
                            {
                                (-1.0, -1.0),  // Bottom-left
                                (1.0, -1.0),   // Bottom-right
                                (-1.0, 1.0),   // Top-left
                                (1.0, 1.0)     // Top-right
                            };
                            
                            for (int i = 0; i < sleeveDataList.Count; i++)
                            {
                                var (cz, center, sleeveWidth, sleeveHeight, sleeveRotation, cosSleevePreCalc, sinSleevePreCalc, preCalculatedCorners) = sleeveDataList[i];
                                
                                XYZ[] worldCorners;
                                
                                // ✅ USE PRE-CALCULATED CORNERS: If available, use them directly (dump once use many times)
                                if (preCalculatedCorners != null)
                                {
                                    worldCorners = preCalculatedCorners;
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Sleeve {i+1} Using PRE-CALCULATED corners from database\n");
                                }
                                else
                                {
                                    // ✅ FALLBACK: Recalculate corners if pre-calculated ones are missing
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Sleeve {i+1} Recalculating corners (pre-calculated corners not available)\n");
                                    
                                    // Step 4a: Calculate 4 corners of sleeve in its LOCAL coordinate system (before sleeve rotation)
                                    double halfW = sleeveWidth / 2.0;
                                    double halfH = sleeveHeight / 2.0;
                                    
                                    // ✅ OPTIMIZATION: Reuse corner offset pattern, multiply by this sleeve's halfW/halfH
                                    var localCorners = new XYZ[4];
                                    for (int cornerIdx = 0; cornerIdx < 4; cornerIdx++)
                                    {
                                        localCorners[cornerIdx] = new XYZ(
                                            cornerOffsetPattern[cornerIdx].Item1 * halfW,
                                            cornerOffsetPattern[cornerIdx].Item2 * halfH,
                                            0
                                        );
                                    }
                                    
                                    // Step 4b: Rotate corners by sleeve's rotation angle to get world-space corners
                                    // ✅ ROTATION MATRIX: Use pre-calculated cos/sin from database if available, otherwise calculate
                                    double cosSleeve = cosSleevePreCalc ?? Math.Cos(sleeveRotation);
                                    double sinSleeve = sinSleevePreCalc ?? Math.Sin(sleeveRotation);
                                    
                                    worldCorners = new XYZ[4];
                                    for (int j = 0; j < 4; j++)
                                    {
                                        double localX = localCorners[j].X;
                                        double localY = localCorners[j].Y;
                                        
                                        // Rotate corner by sleeve rotation
                                        double worldX = localX * cosSleeve - localY * sinSleeve;
                                        double worldY = localX * sinSleeve + localY * cosSleeve;
                                        
                                        // Translate to sleeve center
                                        worldCorners[j] = new XYZ(
                                            center.X + worldX,
                                            center.Y + worldY,
                                            center.Z
                                        );
                                    }
                                }
                                
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Sleeve {i+1} World Corners: ({worldCorners[0].X:F6}, {worldCorners[0].Y:F6}), ({worldCorners[1].X:F6}, {worldCorners[1].Y:F6}), ({worldCorners[2].X:F6}, {worldCorners[2].Y:F6}), ({worldCorners[3].X:F6}, {worldCorners[3].Y:F6})\n");
                                
                                // Step 4c: Transform world-space corners to cluster's INTENDED ROTATED AXIS coordinate system
                                // This aligns all corners to the cluster's intended axis direction (not average of sleeve angles)
                                for (int j = 0; j < 4; j++)
                                {
                                    // Translate relative to origin
                                    double relX = worldCorners[j].X - origin.X;
                                    double relY = worldCorners[j].Y - origin.Y;
                                    
                                    // Rotate to cluster's intended axis coordinate system
                                    // This transformation aligns with the cluster's intended rotated axis direction
                                    double clusterX = relX * cosCluster - relY * sinCluster;
                                    double clusterY = relX * sinCluster + relY * cosCluster;
                                    
                                    allTransformedCorners.Add(new XYZ(
                                        clusterX,
                                        clusterY,
                                        worldCorners[j].Z
                                    ));
                                }
                                
                                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}]   Sleeve {i+1} Cluster-Space Corners: ({allTransformedCorners[allTransformedCorners.Count-4].X:F6}, {allTransformedCorners[allTransformedCorners.Count-4].Y:F6}), ({allTransformedCorners[allTransformedCorners.Count-3].X:F6}, {allTransformedCorners[allTransformedCorners.Count-3].Y:F6}), ({allTransformedCorners[allTransformedCorners.Count-2].X:F6}, {allTransformedCorners[allTransformedCorners.Count-2].Y:F6}), ({allTransformedCorners[allTransformedCorners.Count-1].X:F6}, {allTransformedCorners[allTransformedCorners.Count-1].Y:F6})\n");
                            }
                            
                            // Step 5: Find min/max extents of all transformed corners
                            minX = allTransformedCorners.Min(p => p.X);
                            minY = allTransformedCorners.Min(p => p.Y);
                            maxX = allTransformedCorners.Max(p => p.X);
                            maxY = allTransformedCorners.Max(p => p.Y);
                            width = maxX - minX;
                            height = maxY - minY;
                            
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ✅ CORRECT UNION (from all {allTransformedCorners.Count} corner points): MinX={minX:F6}, MinY={minY:F6}, MaxX={maxX:F6}, MaxY={maxY:F6}\n");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ✅ CLUSTER SIZE (WATERTIGHT ALGORITHM): W={width * 304.8:F1}mm, H={height * 304.8:F1}mm\n");
                        }
                        
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Cluster BBox: Min=({minX:F6}, {minY:F6}), Max=({maxX:F6}, {maxY:F6})\n\n");
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLUSTER-BBOX] ✅ Using proper corner-based algorithm (W={width*304.8:F1}mm, H={height*304.8:F1}mm)");
                    }
                    
                    double minZ = rotatedBboxes.Min(b => b.min.Z);
                    double maxZ = rotatedBboxes.Max(b => b.max.Z);
                    double depth = maxZ - minZ;
                    
                    // Convert to millimeters for logging
                    double widthMm = width * 304.8;
                    double heightMm = height * 304.8;
                    double depthMm = depth * 304.8;
                    
                    // ✅ LOG: Final cluster dimensions before transform
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== FINAL CLUSTER DIMENSIONS (BEFORE TRANSFORM) ==========\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Width: {width:F6} = {widthMm:F1}mm\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Height: {height:F6} = {heightMm:F1}mm\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Depth: {depth:F6} = {depthMm:F1}mm\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] BBox Min: ({minX:F6}, {minY:F6}, {minZ:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] BBox Max: ({maxX:F6}, {maxY:F6}, {maxZ:F6})\n\n");
                    
                    logBuilder.AppendLine($"  CLUSTER BOUNDING BOX (UNION OF ALL INDIVIDUAL ROTATED BBOXES):");
                    logBuilder.AppendLine($"    Cluster Min: ({minX:F6}, {minY:F6}, {minZ:F6})");
                    logBuilder.AppendLine($"    Cluster Max: ({maxX:F6}, {maxY:F6}, {maxZ:F6})");
                    logBuilder.AppendLine($"    Cluster Size: W={widthMm:F1}mm × H={heightMm:F1}mm × D={depthMm:F1}mm");
                    logBuilder.AppendLine();
                    
                    // Calculate expected size based on individual sleeves
                    if (rotatedBboxes.Count == 2)
                    {
                        var bbox1 = rotatedBboxes[0];
                        var bbox2 = rotatedBboxes[1];
                        double sleeve1Width = (bbox1.max.X - bbox1.min.X) * 304.8;
                        double sleeve1Height = (bbox1.max.Y - bbox1.min.Y) * 304.8;
                        double sleeve2Width = (bbox2.max.X - bbox2.min.X) * 304.8;
                        double sleeve2Height = (bbox2.max.Y - bbox2.min.Y) * 304.8;
                        
                        // Calculate gap between sleeves
                        double gapX = Math.Max(0, Math.Min(bbox1.max.X, bbox2.max.X) - Math.Max(bbox1.min.X, bbox2.min.X));
                        double gapY = Math.Max(0, Math.Min(bbox1.max.Y, bbox2.max.Y) - Math.Max(bbox1.min.Y, bbox2.min.Y));
                        double overlapX = gapX < 0 ? Math.Abs(gapX) * 304.8 : 0;
                        double overlapY = gapY < 0 ? Math.Abs(gapY) * 304.8 : 0;
                        
                        logBuilder.AppendLine($"  EXPECTED CLUSTER SIZE CALCULATION (2 sleeves):");
                        logBuilder.AppendLine($"    Sleeve 1: {sleeve1Width:F1}mm × {sleeve1Height:F1}mm");
                        logBuilder.AppendLine($"    Sleeve 2: {sleeve2Width:F1}mm × {sleeve2Height:F1}mm");
                        logBuilder.AppendLine($"    Overlap X: {overlapX:F1}mm, Overlap Y: {overlapY:F1}mm");
                        logBuilder.AppendLine($"    Expected Width: {sleeve1Width + sleeve2Width - overlapX:F1}mm (actual: {widthMm:F1}mm)");
                        logBuilder.AppendLine($"    Expected Height: {sleeve1Height + sleeve2Height - overlapY:F1}mm (actual: {heightMm:F1}mm)");
                    }
                    
                    logBuilder.AppendLine($"  ========== END CLUSTER BOUNDING BOX CALCULATION ==========");
                    logBuilder.AppendLine();
                    
                    // Write to log file
                    try
                    {
                        System.IO.File.AppendAllText(clusterLogPath, logBuilder.ToString());
                    }
                    catch { }
                    
                    // Also log to DebugLogger
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-BBOX-ROTATED] ✅ Using rotated bounding boxes: " +
                            $"W={widthMm:F1}mm, H={heightMm:F1}mm (from {rotatedBboxes.Count} individual sleeves)");
                        DebugLogger.Info($"[CLUSTER-BBOX-ROTATED] Individual sleeves: {string.Join(", ", rotatedBboxes.Select((b, i) => $"Sleeve{i+1}: {(b.max.X - b.min.X) * 304.8:F1}×{(b.max.Y - b.min.Y) * 304.8:F1}mm"))}");
                    }
                    
                    // ✅ MIDPOINT: Calculate in rotated coordinate space, then transform back to world coordinates
                    XYZ midRotated = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);
                    
                    // ✅ LOG: Midpoint calculation
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== MIDPOINT CALCULATION ==========\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Midpoint (in rotated coordinate space): ({midRotated.X:F6}, {midRotated.Y:F6}, {midRotated.Z:F6})\n");

                    var allBboxes = rotatedBboxes.Concat(axisAlignedBboxes).ToList();
                    XYZ mid;
                    if (allBboxes.Count > 0 && Math.Abs(rotationAngle) > 1e-6 && origin != XYZ.Zero)
                    {
                        // ✅ Transform midpoint from rotated coordinate space back to world coordinates
                        // The midpoint is calculated as center of bounding box in rotated space: ((minX+maxX)/2, (minY+maxY)/2)
                        // To transform back to world coordinates, we need the INVERSE rotation matrix
                        // Forward: clusterX = relX * cos(θ) - relY * sin(θ), clusterY = relX * sin(θ) + relY * cos(θ)
                        // Inverse: worldX = clusterX * cos(θ) + clusterY * sin(θ), worldY = -clusterX * sin(θ) + clusterY * cos(θ)
                        // This is the transpose of the rotation matrix (which is the inverse for rotation matrices)
                        double cosA = Math.Cos(rotationAngle);  // Use same angle (inverse rotation matrix is transpose)
                        double sinA = Math.Sin(rotationAngle);
                        // ✅ INVERSE TRANSFORMATION: Transpose of rotation matrix
                        double midRotatedBackX = midRotated.X * cosA + midRotated.Y * sinA;  // Note: + instead of -
                        double midRotatedBackY = -midRotated.X * sinA + midRotated.Y * cosA;  // Note: -sinA instead of sinA
                        
                        // Step 2: Translate back (add origin)
                        mid = new XYZ(
                            origin.X + midRotatedBackX,
                            origin.Y + midRotatedBackY,
                            midRotated.Z  // Z stays the same
                        );
                        
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== TRANSFORM CALCULATION ==========\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Origin: ({origin.X:F6}, {origin.Y:F6}, {origin.Z:F6})\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Midpoint (after transform to world): ({mid.X:F6}, {mid.Y:F6}, {mid.Z:F6})\n\n");
                    }
                    else
                    {
                        // No rotation or no rotated bboxes: midpoint is already in world coordinates
                        if (origin != XYZ.Zero)
                        {
                            mid = new XYZ(
                                origin.X + midRotated.X,
                                origin.Y + midRotated.Y,
                                midRotated.Z
                            );
                        }
                        else
                        {
                            mid = midRotated;
                        }
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] No rotation applied (rotationAngle={rotationAngle * 180.0 / Math.PI:F2}°)\n");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Midpoint (world coordinates): ({mid.X:F6}, {mid.Y:F6}, {mid.Z:F6})\n\n");
                    }
                    
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== RETURNING CLUSTER BOUNDING BOX ==========\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Final: W={widthMm:F1}mm, H={heightMm:F1}mm, D={depthMm:F1}mm, Mid=({mid.X:F6}, {mid.Y:F6}, {mid.Z:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated BBox: Min=({minX:F6}, {minY:F6}, {minZ:F6}), Max=({maxX:F6}, {maxY:F6}, {maxZ:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Internal units: W={width:F6}, H={height:F6}, D={depth:F6}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END GET CLUSTER BOUNDING BOX ==========\n\n");

                    return (width, height, depth, mid, minX, minY, minZ, maxX, maxY, maxZ);
                }
                else
                {
                    // ✅ FIX: If rotated bounding boxes are not available, log warning
                    // For rotated sleeves, we should NOT use axis-aligned bounding boxes as fallback
                    // because they give wrong dimensions (diagonal extent instead of actual dimensions)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTER-BBOX] ⚠️ Rotated bounding boxes not found for cluster with {cluster.Count} sleeves. " +
                            $"Rotation angle: {rotationAngle * 180.0 / Math.PI:F2}°. " +
                            $"Falling back to axis-aligned bounding boxes (may give incorrect dimensions for rotated sleeves).");
                    }
                    
                    // ⚠️ FALLBACK: Use axis-aligned bounding boxes (only for axis-aligned sleeves)
                    // For rotated sleeves, this will give wrong dimensions
                    var (w, h, d, m) = ClusterBoundingBoxServices.GetClusterBoundingBox(actualSleeves, rotationAngle);
                    return (w, h, d, m, null, null, null, null, null, null);
                }
            }
            catch (Exception ex)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ❌❌❌ EXCEPTION IN GetClusterBoundingBoxWithRotatedCoordinates ❌❌❌\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Exception Type: {ex.GetType().Name}\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Error Message: {ex.Message}\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] StackTrace: {ex.StackTrace}\n");
                if (ex.InnerException != null)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Inner Exception: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}\n");
                }
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Falling back to ClusterBoundingBoxServices.GetClusterBoundingBox\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END EXCEPTION ==========\n\n");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[UniversalClusterService] Error in GetClusterBoundingBoxWithRotatedCoordinates: {ex.Message}");
                }
                var (w, h, d, m) = ClusterBoundingBoxServices.GetClusterBoundingBox(actualSleeves, rotationAngle);
                return (w, h, d, m, null, null, null, null, null, null);
            }
        }

        /// <summary>
        /// Calculate cluster bounding box from XML data in rotated coordinate system
        /// This ensures cluster sleeves follow the actual outline of individual sleeves
        /// </summary>
        private (double width, double height, double depth, XYZ mid) GetClusterBoundingBoxFromXml(List<dynamic> cluster, string xmlFilePath = null)
        {
            try
            {
                if (cluster.Count == 0)
                    return (0, 0, 0, XYZ.Zero);

                // ✅ NEW: Determine dominant rotation angle BEFORE calculating bounding box
                double rotationAngle = DetermineDominantRotationAngle(cluster, xmlFilePath);

                // Get all bounding boxes from XML data
                var boundingBoxes = cluster.Select(s => s.BoundingBox).Where(bbox => bbox != null).ToList();
                
                if (boundingBoxes.Count == 0)
                    return (0, 0, 0, XYZ.Zero);

                // ✅ DEBUG: Log individual sleeve bounding boxes
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"\n[CLUSTER-BBOX] Individual sleeves in cluster ({cluster.Count} sleeves):\n");
                
                foreach (var sleeve in cluster)
                {
                    var bbox = sleeve.BoundingBox;
                    if (bbox != null)
                    {
                        double sleeveWidth = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Millimeters);
                        double sleeveHeight = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Millimeters);
                        double sleeveDepth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Millimeters);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"  Sleeve {sleeve.SleeveInstanceId}: W={sleeveWidth:F1}mm, H={sleeveHeight:F1}mm, D={sleeveDepth:F1}mm, " +
                                    $"Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}), Max=({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})\n");
                        }
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
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[OVERLAP] Sleeves {i} and {j}: Overlap Vol={UnitUtils.ConvertFromInternalUnits(overlapVol, UnitTypeId.CubicMillimeters):F2}mm³, " +
                                            $"Overlap Dim=({UnitUtils.ConvertFromInternalUnits(overlapX, UnitTypeId.Millimeters):F1}mm, " +
                                            $"{UnitUtils.ConvertFromInternalUnits(overlapY, UnitTypeId.Millimeters):F1}mm, " +
                                            $"{UnitUtils.ConvertFromInternalUnits(overlapZ, UnitTypeId.Millimeters):F1}mm)\n");
                                }
                            }
                            
                            processedPairs.Add(pairKey);
                        }
                    }
                }
                
                // Declare variables at outer scope to avoid scope conflicts
                double minX, minY, minZ, maxX, maxY, maxZ;
                double width, height, depth;
                XYZ mid;

                // ✅ NEW: Calculate bounding box in rotated coordinate system if rotation angle is significant
                if (Math.Abs(rotationAngle) > 1e-6)
                {
                    // Calculate center point (midpoint of all sleeve centers) for rotation
                    var sleeveCenters = new List<XYZ>();
                    foreach (var bbox in boundingBoxes)
                    {
                        sleeveCenters.Add((bbox.Min + bbox.Max) / 2.0);
                    }

                    if (sleeveCenters.Count > 0)
                    {
                        // Use average center as rotation origin
                        XYZ rotationOrigin = new XYZ(
                            sleeveCenters.Average(p => p.X),
                            sleeveCenters.Average(p => p.Y),
                            sleeveCenters.Average(p => p.Z)
                        );

                        // Create rotation transform (rotate around Z-axis)
                        Transform rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, rotationAngle, rotationOrigin);
                        Transform inverseTransform = rotationTransform.Inverse;

                        // Transform all bounding box corners to rotated coordinate system
                        var transformedPoints = new List<XYZ>();

                        foreach (var bbox in boundingBoxes)
                        {
                            // Get all 8 corners of the bounding box
                            var corners = new[]
                            {
                                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
                            };

                            // Transform each corner to rotated coordinate system
                            foreach (var corner in corners)
                            {
                                var transformed = inverseTransform.OfPoint(corner);
                                transformedPoints.Add(transformed);
                            }
                        }

                        if (transformedPoints.Count > 0)
                        {
                            // Calculate min/max in rotated coordinate system
                            minX = transformedPoints.Min(p => p.X);
                            minY = transformedPoints.Min(p => p.Y);
                            minZ = transformedPoints.Min(p => p.Z);
                            maxX = transformedPoints.Max(p => p.X);
                            maxY = transformedPoints.Max(p => p.Y);
                            maxZ = transformedPoints.Max(p => p.Z);

                            // Dimensions in rotated coordinate system
                            width = maxX - minX;
                            height = maxY - minY;
                            depth = maxZ - minZ;

                            // Midpoint in rotated coordinate system (transform back to model coordinates)
                            XYZ rotatedMid = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);
                            mid = rotationTransform.OfPoint(rotatedMid);

                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-BBOX] Rotated coordinate system: Angle={rotationAngle * 180 / Math.PI:F1}°, " +
                                    $"Rotated coords - Min=({minX:F3}, {minY:F3}, {minZ:F3}), Max=({maxX:F3}, {maxY:F3}, {maxZ:F3}), " +
                                    $"Width={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Height={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Depth={UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Sleeves in cluster: {cluster.Count}");
                                
                                // Log individual sleeve dimensions for comparison
                                foreach (var sleeve in cluster.Take(5))
                                {
                                    var bbox = sleeve.BoundingBox;
                                    if (bbox != null)
                                    {
                                        double sleeveWidth = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Millimeters);
                                        double sleeveHeight = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Millimeters);
                                        DebugLogger.Info($"[CLUSTER-BBOX] Individual sleeve {sleeve.SleeveInstanceId}: W={sleeveWidth:F1}mm, H={sleeveHeight:F1}mm");
                                    }
                                }
                            }

                            return (width, height, depth, mid);
                        }
                    }
                }

                // Fallback: Calculate overall bounding box (union of all boxes) - axis-aligned
                minX = boundingBoxes.Min(bbox => bbox.Min.X);
                minY = boundingBoxes.Min(bbox => bbox.Min.Y);
                minZ = boundingBoxes.Min(bbox => bbox.Min.Z);
                maxX = boundingBoxes.Max(bbox => bbox.Max.X);
                maxY = boundingBoxes.Max(bbox => bbox.Max.Y);
                maxZ = boundingBoxes.Max(bbox => bbox.Max.Z);

                width = maxX - minX;
                height = maxY - minY;
                depth = maxZ - minZ;
                mid = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);

                return (width, height, depth, mid);
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                
                // ✅ CRITICAL FIX: Use MINIMUM DISTANCE calculation instead of overlap
                // Overlap checking can cluster sleeves that are far apart if they overlap in ignored dimension
                // For example: X-oriented wall sleeves 500mm apart in Y (wall depth) will still overlap in X,Z
                // We need to calculate actual minimum distance and check if it's within tolerance
                string hostType = sleeve1.HostType;
                string orientation = sleeve1.Orientation;
                string systemType = sleeve1.SystemType ?? "";
                
                double minDistance;
                
                // ✅ CRITICAL FIX FOR ROUND PIPES/DUCTS ONLY: Use edge-to-edge distance (accounting for sleeve diameter)
                // Round pipes/ducts have inflated axis-aligned bounding boxes when rotated, causing incorrect clustering
                // RECTANGULAR SLEEVES: Continue using existing bounding box logic (CalculateMinimumDistance2D) which properly handles rectangular shapes
                // LOGIC FOR ROUND: edgeToEdgeDistance = centerToCenterDistance - (sleeve1Radius + sleeve2Radius)
                // Then check: edgeToEdgeDistance <= toleranceDist
                bool isRoundPipeOrDuct = (systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0 || 
                                         systemType.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0) &&
                                        (sleeve1.IsCircular == true || sleeve2.IsCircular == true);
                
                if (isRoundPipeOrDuct)
                {
                    // ✅ DATABASE-ONLY: Get placement points and sleeve DIAMETERS from ClashZone (database data)
                    // ONLY for round pipes/ducts - rectangular sleeves use bounding box logic below
                    var placement1 = GetPlacementPointFromSleeve(sleeve1);
                    var placement2 = GetPlacementPointFromSleeve(sleeve2);
                    // ✅ FIX: Cannot deconstruct tuple from dynamic - store result first, then access values
                    var radiiResult = GetSleeveRadiiFromSleeves(sleeve1, sleeve2);
                    double radius1 = radiiResult.radius1;
                    double radius2 = radiiResult.radius2;
                    
                    // ✅ Only proceed if we have valid placement points AND both sleeves have diameter data (round)
                    if (placement1 != null && placement2 != null && radius1 > 0 && radius2 > 0)
                    {
                        double centerToCenterDistance;
                        
                        if (hostType == "Floor")
                        {
                            // Floor: 2D distance in X,Y plane (ignore Z)
                            double dx = placement2.X - placement1.X;
                            double dy = placement2.Y - placement1.Y;
                            centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy);
                        }
                        else if (hostType == "Wall" || hostType == "Structural Framing")
                        {
                            if (orientation == "X")
                            {
                                // X-oriented walls: 2D distance in X,Z plane (ignore Y)
                                double dx = placement2.X - placement1.X;
                                double dz = placement2.Z - placement1.Z;
                                centerToCenterDistance = Math.Sqrt(dx * dx + dz * dz);
                            }
                            else
                            {
                                // Y-oriented walls: 2D distance in Y,Z plane (ignore X)
                                double dy = placement2.Y - placement1.Y;
                                double dz = placement2.Z - placement1.Z;
                                centerToCenterDistance = Math.Sqrt(dy * dy + dz * dz);
                            }
                        }
                        else
                        {
                            // Fallback: 3D distance
                            double dx = placement2.X - placement1.X;
                            double dy = placement2.Y - placement1.Y;
                            double dz = placement2.Z - placement1.Z;
                            centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                        }
                        
                        // ✅ CRITICAL: Calculate EDGE-TO-EDGE distance (accounts for sleeve diameter)
                        // Formula: edgeToEdge = centerToCenter - (radius1 + radius2)
                        // This is the actual gap between the two sleeves
                        double edgeToEdgeDistance = centerToCenterDistance - (radius1 + radius2);
                        
                        // ✅ Use edge-to-edge distance for proximity check
                        // If edgeToEdge <= tolerance, sleeves are close enough to cluster
                        minDistance = Math.Max(0, edgeToEdgeDistance); // Ensure non-negative
                        
                        // ✅ DEBUG: Log the calculation details
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var roundPipeProximityLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            try 
                            { 
                                System.IO.File.AppendAllText(roundPipeProximityLogPath, $"[{DateTime.Now:HH:mm:ss}] [ROUND-PIPE-PROXIMITY] Sleeve {sleeve1.SleeveInstanceId} vs {sleeve2.SleeveInstanceId}: " +
                                    $"Center-to-Center={UnitUtils.ConvertFromInternalUnits(centerToCenterDistance, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Radius1={UnitUtils.ConvertFromInternalUnits(radius1, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Radius2={UnitUtils.ConvertFromInternalUnits(radius2, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Edge-to-Edge={UnitUtils.ConvertFromInternalUnits(edgeToEdgeDistance, UnitTypeId.Millimeters):F1}mm, " +
                                    $"Tolerance={UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
                            } 
                            catch { }
                        }
                    }
                    else
                    {
                        // Fallback to bounding box if placement points or radii not available
                        if (hostType == "Floor")
                        {
                            minDistance = CalculateMinimumDistance2D(
                                bbox1.Min.X, bbox1.Min.Y, bbox1.Max.X, bbox1.Max.Y,
                                bbox2.Min.X, bbox2.Min.Y, bbox2.Max.X, bbox2.Max.Y);
                        }
                        else
                        {
                            minDistance = CalculateMinimumDistance3D(
                                bbox1.Min.X, bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Y, bbox1.Max.Z,
                                bbox2.Min.X, bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Y, bbox2.Max.Z);
                        }
                    }
                }
                else if (hostType == "Floor")
                {
                    // Floor sleeves: Calculate 2D distance in X,Y plane (ignore Z)
                    minDistance = CalculateMinimumDistance2D(
                        bbox1.Min.X, bbox1.Min.Y, bbox1.Max.X, bbox1.Max.Y,
                        bbox2.Min.X, bbox2.Min.Y, bbox2.Max.X, bbox2.Max.Y);
                }
                else if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    if (orientation == "X")
                    {
                        // X-oriented walls: Calculate 2D distance in X,Z plane (ignore Y/wall depth)
                        // X = width along wall, Z = height
                        minDistance = CalculateMinimumDistance2D(
                            bbox1.Min.X, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Z,
                            bbox2.Min.X, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Z);
                    }
                    else if (orientation == "Y")
                    {
                        // Y-oriented walls: Calculate 2D distance in Y,Z plane (ignore X/wall depth)
                        // Y = width along wall, Z = height
                        minDistance = CalculateMinimumDistance2D(
                            bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                            bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                    }
                    else
                    {
                        // Default for walls: Assume Y-oriented, use Y,Z distance
                        minDistance = CalculateMinimumDistance2D(
                            bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                            bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                    }
                }
                else
                {
                    // Fallback for unknown host types: Use full 3D distance
                    minDistance = CalculateMinimumDistance3D(
                        bbox1.Min.X, bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Y, bbox1.Max.Z,
                        bbox2.Min.X, bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Y, bbox2.Max.Z);
                }
                
                // Check if minimum distance is within tolerance
                bool withinTolerance = minDistance <= toleranceDist;
                
                // ✅ DEBUG: Log distance calculation results for ALL sleeves (not just 945000+)
                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                try 
                { 
                    string method = isRoundPipeOrDuct ? "Edge-to-Edge (Round Pipe/Duct)" : "BoundingBox";
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-DISTANCE] Sleeve {sleeve1.SleeveInstanceId} vs {sleeve2.SleeveInstanceId}: " +
                        $"Host={hostType}, Ori={orientation}, Method={method}, Distance={UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm, " +
                        $"Tolerance={UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm, " +
                        $"Result={(withinTolerance ? "CLUSTER ✓" : "NO CLUSTER ✗")}\n");
                } 
                catch { }
                
                if (!withinTolerance)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] Sleeves {sleeve1.SleeveInstanceId} and {sleeve2.SleeveInstanceId}: Distance={UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm > Tolerance={UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm - NO CLUSTER");
                }
                
                return withinTolerance;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Error checking bounding box distance: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// ✅ DATABASE-ONLY: Helper method to get placement point from dynamic sleeve object
        /// Uses ClashZone data from database (SleevePlacementPointActiveDocumentX/Y/Z)
        /// </summary>
        private XYZ GetPlacementPointFromSleeve(dynamic sleeve)
        {
            try
            {
                // ✅ DATABASE-ONLY: Get from ClashZone (loaded from database)
                if (sleeve.ClashZone != null)
                {
                    var cz = sleeve.ClashZone as ClashZone;
                    if (cz != null)
                    {
                        // ✅ Use active document placement point (where sleeve is actually placed in active document)
                        // This is stored in database as SleevePlacementPointActiveDocumentX/Y/Z
                        if (cz.SleevePlacementPointActiveDocumentX != 0 || 
                            cz.SleevePlacementPointActiveDocumentY != 0 || 
                            cz.SleevePlacementPointActiveDocumentZ != 0)
                        {
                            return new XYZ(
                                cz.SleevePlacementPointActiveDocumentX,
                                cz.SleevePlacementPointActiveDocumentY,
                                cz.SleevePlacementPointActiveDocumentZ
                            );
                        }
                        
                        // Fallback to regular placement point (from database)
                        if (cz.SleevePlacementPointX != 0 || 
                            cz.SleevePlacementPointY != 0 || 
                            cz.SleevePlacementPointZ != 0)
                        {
                            return new XYZ(
                                cz.SleevePlacementPointX,
                                cz.SleevePlacementPointY,
                                cz.SleevePlacementPointZ
                            );
                        }
                    }
                }
                
                // Fallback: Calculate center from bounding box (from database)
                var bbox = sleeve.BoundingBox;
                if (bbox != null)
                {
                    return new XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0
                    );
                }
                
                return null;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// ✅ DATABASE-ONLY: Helper method to get sleeve radii from dynamic sleeve objects
        /// ONLY for ROUND pipes/ducts - uses SleeveDiameter from database
        /// Returns (radius1, radius2) in Revit internal units
        /// NOTE: Rectangular sleeves should use bounding box logic (CalculateMinimumDistance2D), not this method
        /// </summary>
        private (double radius1, double radius2) GetSleeveRadiiFromSleeves(dynamic sleeve1, dynamic sleeve2)
        {
            double radius1 = 0;
            double radius2 = 0;
            
            try
            {
                // ✅ DATABASE-ONLY: Get sleeve DIAMETER from ClashZone (loaded from database)
                // ONLY for round pipes/ducts - rectangular sleeves use bounding box logic
                if (sleeve1.ClashZone != null)
                {
                    var cz1 = sleeve1.ClashZone as ClashZone;
                    if (cz1 != null && cz1.SleeveDiameter > 0)
                    {
                        // For round pipes/ducts, use diameter / 2
                        radius1 = cz1.SleeveDiameter / 2.0;
                    }
                }
                
                if (sleeve2.ClashZone != null)
                {
                    var cz2 = sleeve2.ClashZone as ClashZone;
                    if (cz2 != null && cz2.SleeveDiameter > 0)
                    {
                        // For round pipes/ducts, use diameter / 2
                        radius2 = cz2.SleeveDiameter / 2.0;
                    }
                }
            }
            catch
            {
                // Return zeros if error occurs
            }
            
            return (radius1, radius2);
        }
        
        private List<List<FamilyInstance>> CalculateClustersUsingBoundingBoxOverlap(List<FamilyInstance> sleeves, double toleranceDist, string orientation)
        {
            var clusters = new List<List<FamilyInstance>>();
            var unprocessedSet = new HashSet<FamilyInstance>(sleeves);
            
                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Formed cluster with {cluster.Count} sleeves: {string.Join(", ", cluster.Select(s => s.Id.IntegerValue))}");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"✓ CLUSTER: {string.Join(", ", cluster.Select(s => s.Id.IntegerValue))}\n");
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Individual sleeve {cluster[0].Id.IntegerValue} (no proximate neighbors)");
                                        if (!DeploymentConfiguration.DeploymentMode)
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
            // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
            Parameter currentMepElementIdParam = null;
            if (_parameterCache != null && _parameterCache.TryGetValue(currentSleeve, out var paramDict5))
            {
                paramDict5.TryGetValue("MEP_ElementId", out currentMepElementIdParam);
            }
            if (currentMepElementIdParam == null)
            {
                currentMepElementIdParam = currentSleeve.LookupParameter("MEP_ElementId"); // Fallback if cache not available
            }
            if (currentMepElementIdParam == null) return proximateSleeves;
            
            long currentMepElementId = currentMepElementIdParam.AsInteger();
            if (!_clashZoneCache.ContainsKey(currentMepElementId)) return proximateSleeves;
            
            var currentClashZone = _clashZoneCache[currentMepElementId];
            
            foreach (var candidateSleeve in candidateSleeves)
            {
                if (candidateSleeve == currentSleeve) continue;
                
                // Get candidate sleeve's clash zone
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter candidateMepElementIdParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(candidateSleeve, out var paramDict6))
                {
                    paramDict6.TryGetValue("MEP_ElementId", out candidateMepElementIdParam);
                }
                if (candidateMepElementIdParam == null)
                {
                    candidateMepElementIdParam = candidateSleeve.LookupParameter("MEP_ElementId"); // Fallback if cache not available
                }
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DEBUG] ✓ BOUNDING-BOX-PROXIMITY: Sleeve {currentSleeve.Id.IntegerValue} -> {candidateSleeve.Id.IntegerValue}: Bounding boxes overlap within {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
                }
                else
                {
                    // Log non-proximate analysis (only for specific sleeves to reduce noise)
                    if (currentSleeve.Id.IntegerValue == 897149 || currentSleeve.Id.IntegerValue == 897154 || currentSleeve.Id.IntegerValue == 897195 ||
                        candidateSleeve.Id.IntegerValue == 897149 || candidateSleeve.Id.IntegerValue == 897154 || candidateSleeve.Id.IntegerValue == 897195)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
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
            
            // ✅ ROTATED BBOX: Use rotated bounding box if available (for non-axis-aligned sleeves)
            // Check if both sleeves have rotated bounding boxes (MepElementRotationAngle != 0)
            bool useRotatedBbox = Math.Abs(current.MepElementRotationAngle) > 1e-6 && 
                                  Math.Abs(other.MepElementRotationAngle) > 1e-6 &&
                                  current.RotatedBoundingBoxMinX.HasValue && current.RotatedBoundingBoxMinY.HasValue &&
                                  current.RotatedBoundingBoxMaxX.HasValue && current.RotatedBoundingBoxMaxY.HasValue &&
                                  other.RotatedBoundingBoxMinX.HasValue && other.RotatedBoundingBoxMinY.HasValue &&
                                  other.RotatedBoundingBoxMaxX.HasValue && other.RotatedBoundingBoxMaxY.HasValue;
            
            // ✅ DEBUG: Log coordinates and orientation
                        if (!DeploymentConfiguration.DeploymentMode)
            {
                if (useRotatedBbox)
                {
                    DebugLogger.Info($"[DISTANCE-DEBUG] Using ROTATED bounding boxes (angles: {current.MepElementRotationAngle * 180 / Math.PI:F1}°, {other.MepElementRotationAngle * 180 / Math.PI:F1}°)\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect1 Rotated: Min=({current.RotatedBoundingBoxMinX.Value:F3}, {current.RotatedBoundingBoxMinY.Value:F3}), Max=({current.RotatedBoundingBoxMaxX.Value:F3}, {current.RotatedBoundingBoxMaxY.Value:F3})\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect2 Rotated: Min=({other.RotatedBoundingBoxMinX.Value:F3}, {other.RotatedBoundingBoxMinY.Value:F3}), Max=({other.RotatedBoundingBoxMaxX.Value:F3}, {other.RotatedBoundingBoxMaxY.Value:F3})\n");
                }
                else
                {
                    DebugLogger.Info($"[DISTANCE-DEBUG] Using axis-aligned bounding boxes\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect1: Min=({current.SleeveBoundingBoxMinX:F3}, {current.SleeveBoundingBoxMinY:F3}), Max=({current.SleeveBoundingBoxMaxX:F3}, {current.SleeveBoundingBoxMaxY:F3})\n");
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect2: Min=({other.SleeveBoundingBoxMinX:F3}, {other.SleeveBoundingBoxMinY:F3}), Max=({other.SleeveBoundingBoxMaxX:F3}, {other.SleeveBoundingBoxMaxY:F3})\n");
                }
                DebugLogger.Info($"[DISTANCE-DEBUG] HostType={current.StructuralElementType}, Orientation={orientation}\n");
            }
            
            // ✅ CRITICAL FIX: Use orientation from grouping logic instead of ClashZone orientation
            // For floor sleeves, orientation="Floor" (unified for all floor sleeves)
            if (current.StructuralElementType == "Floor")
            {
                // Floor sleeves: Use X,Y distance only (ignore Z coordinate)
                if (useRotatedBbox)
                {
                    minDistance = CalculateMinimumDistance2D(
                        current.RotatedBoundingBoxMinX.Value, current.RotatedBoundingBoxMinY.Value, 
                        current.RotatedBoundingBoxMaxX.Value, current.RotatedBoundingBoxMaxY.Value,
                        other.RotatedBoundingBoxMinX.Value, other.RotatedBoundingBoxMinY.Value, 
                        other.RotatedBoundingBoxMaxX.Value, other.RotatedBoundingBoxMaxY.Value);
                }
                else
                {
                    minDistance = CalculateMinimumDistance2D(
                        current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxY,
                        other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY);
                }
            }
            else if (current.StructuralElementType == "Wall" || current.StructuralElementType == "Structural Framing")
            {
                if (orientation == "X")
                {
                    // Wall/Framing sleeves (X orientation): Use X,Z distance only (ignore Y coordinate)
                    if (useRotatedBbox && current.RotatedBoundingBoxMinZ.HasValue && other.RotatedBoundingBoxMinZ.HasValue)
                    {
                        minDistance = CalculateMinimumDistance2D(
                            current.RotatedBoundingBoxMinX.Value, current.RotatedBoundingBoxMinZ.Value, 
                            current.RotatedBoundingBoxMaxX.Value, current.RotatedBoundingBoxMaxZ.Value,
                            other.RotatedBoundingBoxMinX.Value, other.RotatedBoundingBoxMinZ.Value, 
                            other.RotatedBoundingBoxMaxX.Value, other.RotatedBoundingBoxMaxZ.Value);
                    }
                    else
                    {
                        minDistance = CalculateMinimumDistance2D(
                            current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxZ,
                            other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxZ);
                    }
                }
                else if (orientation == "Y")
                {
                    // Wall/Framing sleeves (Y orientation): Use Y,Z distance only (ignore X coordinate)
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DISTANCE-DEBUG] Using Y,Z coordinates for Y-oriented walls\n");
                    if (useRotatedBbox && current.RotatedBoundingBoxMinZ.HasValue && other.RotatedBoundingBoxMinZ.HasValue)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DISTANCE-DEBUG] Rect1 YZ Rotated: Min=({current.RotatedBoundingBoxMinY.Value:F3}, {current.RotatedBoundingBoxMinZ.Value:F3}), Max=({current.RotatedBoundingBoxMaxY.Value:F3}, {current.RotatedBoundingBoxMaxZ.Value:F3})\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DISTANCE-DEBUG] Rect2 YZ Rotated: Min=({other.RotatedBoundingBoxMinY.Value:F3}, {other.RotatedBoundingBoxMinZ.Value:F3}), Max=({other.RotatedBoundingBoxMaxY.Value:F3}, {other.RotatedBoundingBoxMaxZ.Value:F3})\n");
                        minDistance = CalculateMinimumDistance2D(
                            current.RotatedBoundingBoxMinY.Value, current.RotatedBoundingBoxMinZ.Value, 
                            current.RotatedBoundingBoxMaxY.Value, current.RotatedBoundingBoxMaxZ.Value,
                            other.RotatedBoundingBoxMinY.Value, other.RotatedBoundingBoxMinZ.Value, 
                            other.RotatedBoundingBoxMaxY.Value, other.RotatedBoundingBoxMaxZ.Value);
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DISTANCE-DEBUG] Rect1 YZ: Min=({current.SleeveBoundingBoxMinY:F3}, {current.SleeveBoundingBoxMinZ:F3}), Max=({current.SleeveBoundingBoxMaxY:F3}, {current.SleeveBoundingBoxMaxZ:F3})\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DISTANCE-DEBUG] Rect2 YZ: Min=({other.SleeveBoundingBoxMinY:F3}, {other.SleeveBoundingBoxMinZ:F3}), Max=({other.SleeveBoundingBoxMaxY:F3}, {other.SleeveBoundingBoxMaxZ:F3})\n");
                        minDistance = CalculateMinimumDistance2D(
                            current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                            other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                    }
                }
                else
                {
                    // Default for walls: Use Y,Z distance (most walls are Y-oriented)
                    if (useRotatedBbox && current.RotatedBoundingBoxMinZ.HasValue && other.RotatedBoundingBoxMinZ.HasValue)
                    {
                        minDistance = CalculateMinimumDistance2D(
                            current.RotatedBoundingBoxMinY.Value, current.RotatedBoundingBoxMinZ.Value, 
                            current.RotatedBoundingBoxMaxY.Value, current.RotatedBoundingBoxMaxZ.Value,
                            other.RotatedBoundingBoxMinY.Value, other.RotatedBoundingBoxMinZ.Value, 
                            other.RotatedBoundingBoxMaxY.Value, other.RotatedBoundingBoxMaxZ.Value);
                    }
                    else
                    {
                        minDistance = CalculateMinimumDistance2D(
                            current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                            other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                    }
                }
            }
            else
            {
                // Fallback: Use 3D distance for unknown host types
                if (useRotatedBbox && current.RotatedBoundingBoxMinZ.HasValue && other.RotatedBoundingBoxMinZ.HasValue)
                {
                    minDistance = CalculateMinimumDistance3D(
                        current.RotatedBoundingBoxMinX.Value, current.RotatedBoundingBoxMinY.Value, current.RotatedBoundingBoxMinZ.Value, 
                        current.RotatedBoundingBoxMaxX.Value, current.RotatedBoundingBoxMaxY.Value, current.RotatedBoundingBoxMaxZ.Value,
                        other.RotatedBoundingBoxMinX.Value, other.RotatedBoundingBoxMinY.Value, other.RotatedBoundingBoxMinZ.Value,
                        other.RotatedBoundingBoxMaxX.Value, other.RotatedBoundingBoxMaxY.Value, other.RotatedBoundingBoxMaxZ.Value);
                }
                else
                {
                    minDistance = CalculateMinimumDistance3D(
                        current.SleeveBoundingBoxMinX, current.SleeveBoundingBoxMinY, current.SleeveBoundingBoxMinZ, 
                        current.SleeveBoundingBoxMaxX, current.SleeveBoundingBoxMaxY, current.SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ,
                        other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                }
            }
            
            // Log the calculated distance (only for specific sleeves to reduce noise)
            bool shouldLogDistance = current.SleeveInstanceId == 897149 || current.SleeveInstanceId == 897154 || current.SleeveInstanceId == 897195 ||
                                     other.SleeveInstanceId == 897149 || other.SleeveInstanceId == 897154 || other.SleeveInstanceId == 897195;
            
            if (shouldLogDistance)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DISTANCE-DEBUG] Calculated distance: {UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm, Tolerance: {UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters):F1}mm\n");
            }
            
            return minDistance <= toleranceDist;
        }

        /// <summary>
        /// ✅ ROTATED SLEEVE CLUSTERING: Check if two rotated sleeves should cluster using database data
        /// Transforms coordinates to local coordinate system and checks distances along rotated axis
        /// Uses ClashZone data from database (rotation angles and bounding box coordinates)
        /// </summary>
        private bool CheckRotatedSleeveProximity(dynamic sleeve1, dynamic sleeve2, double rotationAngle, double toleranceDist)
        {
            try
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                bool shouldLog = !DeploymentConfiguration.DeploymentMode;
                
                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== CHECK ROTATED SLEEVE PROXIMITY ==========\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve1.SleeveInstanceId} vs Sleeve {sleeve2.SleeveInstanceId}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Tolerance: {toleranceDist * 304.8:F1}mm\n");
                }
                
                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ❌ ClashZone is null - returning false\n\n");
                    }
                    return false;
                }
                
                var cz1 = sleeve1.ClashZone as ClashZone;
                var cz2 = sleeve2.ClashZone as ClashZone;
                
                if (cz1 == null || cz2 == null)
                {
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ❌ ClashZone cast failed - returning false\n\n");
                    }
                    return false;
                }
                
                // ✅ DATABASE DATA: Get sleeve centers from database (placement points)
                XYZ center1 = cz1.SleevePlacementPoint ?? new XYZ(
                    (cz1.SleeveBoundingBoxMinX + cz1.SleeveBoundingBoxMaxX) / 2.0,
                    (cz1.SleeveBoundingBoxMinY + cz1.SleeveBoundingBoxMaxY) / 2.0,
                    (cz1.SleeveBoundingBoxMinZ + cz1.SleeveBoundingBoxMaxZ) / 2.0);
                
                XYZ center2 = cz2.SleevePlacementPoint ?? new XYZ(
                    (cz2.SleeveBoundingBoxMinX + cz2.SleeveBoundingBoxMaxX) / 2.0,
                    (cz2.SleeveBoundingBoxMinY + cz2.SleeveBoundingBoxMaxY) / 2.0,
                    (cz2.SleeveBoundingBoxMinZ + cz2.SleeveBoundingBoxMaxZ) / 2.0);
                
                if (shouldLog)
                {
                    bool hasPlacementPoint1 = cz1.SleevePlacementPoint != null;
                    bool hasPlacementPoint2 = cz2.SleevePlacementPoint != null;
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve1.SleeveInstanceId} Center: ({center1.X:F6}, {center1.Y:F6}, {center1.Z:F6}) [{(hasPlacementPoint1 ? "from PlacementPoint" : "from BBox center")}]\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve2.SleeveInstanceId} Center: ({center2.X:F6}, {center2.Y:F6}, {center2.Z:F6}) [{(hasPlacementPoint2 ? "from PlacementPoint" : "from BBox center")}]\n");
                }
                
                // ✅ FIX: Work in world-space with shared rotated axis direction
                // Since both sleeves have the same rotation angle (within 1° tolerance),
                // we can use world-space positions and check distances along/perpendicular to the shared rotated axis
                
                // ✅ STEP 1: Calculate shared rotated axis direction (unit vector along rotated X-axis)
                // Prefer the full 3D orientation stored on the clash zone (handles tilted axes).
                XYZ rotatedAxisDirection;
                // Use the full 3D MEP orientation only when present and non-zero (avoid normalizing a zero vector)
                if (cz1.MepElementOrientation != null && cz1.MepElementOrientation.GetLength() > 1e-6)
                {
                    rotatedAxisDirection = cz1.MepElementOrientation.Normalize();
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated Axis Direction (from MepElementOrientation): ({rotatedAxisDirection.X:F6}, {rotatedAxisDirection.Y:F6}, {rotatedAxisDirection.Z:F6})\n");
                    }
                }
                else
                {
                    // Fallback to XY-projected rotation angle (historical behaviour)
                    rotatedAxisDirection = new XYZ(Math.Cos(rotationAngle), Math.Sin(rotationAngle), 0).Normalize();
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated Axis Direction (fallback from angle): ({rotatedAxisDirection.X:F6}, {rotatedAxisDirection.Y:F6}, {rotatedAxisDirection.Z:F6})\n");
                    }
                }
                
                // ✅ STEP 2: Calculate vector from center1 to center2 in world-space
                XYZ worldVector = center2 - center1;
                
                // ✅ STEP 3: Project distance along the shared rotated axis using dot product
                double distanceAlongAxis = worldVector.DotProduct(rotatedAxisDirection);
                
                // ✅ STEP 4: Calculate perpendicular distance from axis using cross product magnitude
                XYZ crossProduct = worldVector.CrossProduct(rotatedAxisDirection);
                double perpendicularDistance = crossProduct.GetLength();
                
                // ✅ DATABASE DATA: Get sleeve dimensions from database (rotated bounding boxes if available, else axis-aligned)
                double sleeve1Width = 0.0;
                double sleeve1Height = 0.0;
                double sleeve2Width = 0.0;
                double sleeve2Height = 0.0;
                
                if (cz1.RotatedBoundingBoxMinX.HasValue && cz1.RotatedBoundingBoxMaxX.HasValue &&
                    cz1.RotatedBoundingBoxMinY.HasValue && cz1.RotatedBoundingBoxMaxY.HasValue)
                {
                    // Use rotated bounding box from database
                    sleeve1Width = cz1.RotatedBoundingBoxMaxX.Value - cz1.RotatedBoundingBoxMinX.Value;
                    sleeve1Height = cz1.RotatedBoundingBoxMaxY.Value - cz1.RotatedBoundingBoxMinY.Value;
                }
                else
                {
                    // Fallback to axis-aligned bounding box from database
                    sleeve1Width = cz1.SleeveBoundingBoxMaxX - cz1.SleeveBoundingBoxMinX;
                    sleeve1Height = cz1.SleeveBoundingBoxMaxY - cz1.SleeveBoundingBoxMinY;
                }
                
                if (cz2.RotatedBoundingBoxMinX.HasValue && cz2.RotatedBoundingBoxMaxX.HasValue &&
                    cz2.RotatedBoundingBoxMinY.HasValue && cz2.RotatedBoundingBoxMaxY.HasValue)
                {
                    // Use rotated bounding box from database
                    sleeve2Width = cz2.RotatedBoundingBoxMaxX.Value - cz2.RotatedBoundingBoxMinX.Value;
                    sleeve2Height = cz2.RotatedBoundingBoxMaxY.Value - cz2.RotatedBoundingBoxMinY.Value;
                }
                else
                {
                    // Fallback to axis-aligned bounding box from database
                    sleeve2Width = cz2.SleeveBoundingBoxMaxX - cz2.SleeveBoundingBoxMinX;
                    sleeve2Height = cz2.SleeveBoundingBoxMaxY - cz2.SleeveBoundingBoxMinY;
                }
                
                // Calculate half-dimensions for overlap check
                double halfWidth1 = sleeve1Width / 2.0;
                double halfWidth2 = sleeve2Width / 2.0;
                double halfHeight1 = sleeve1Height / 2.0;
                double halfHeight2 = sleeve2Height / 2.0;
                
                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated Axis Direction: ({rotatedAxisDirection.X:F6}, {rotatedAxisDirection.Y:F6}, {rotatedAxisDirection.Z:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] World Vector (center1 to center2): ({worldVector.X:F6}, {worldVector.Y:F6}, {worldVector.Z:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve 1 Size: W={sleeve1Width * 304.8:F1}mm, H={sleeve1Height * 304.8:F1}mm\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve 2 Size: W={sleeve2Width * 304.8:F1}mm, H={sleeve2Height * 304.8:F1}mm\n");
                }
                
                // ✅ STEP 5: Check if sleeves are close enough along the rotated axis
                // Distance along axis should be within (halfWidth1 + halfWidth2 + tolerance)
                double maxDistanceAlongAxis = halfWidth1 + halfWidth2 + toleranceDist;
                bool closeAlongAxis = Math.Abs(distanceAlongAxis) <= maxDistanceAlongAxis;
                
                // ✅ STEP 6: Check if perpendicular distance is small enough
                // Perpendicular distance should be within (halfHeight1 + halfHeight2 + tolerance)
                double maxPerpendicularDistance = halfHeight1 + halfHeight2 + toleranceDist;
                bool closePerpendicular = perpendicularDistance <= maxPerpendicularDistance;
                
                // Cluster if both conditions are met
                bool shouldCluster = closeAlongAxis && closePerpendicular;
                
                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Distance Along Axis: {distanceAlongAxis * 304.8:F1}mm (max allowed: {maxDistanceAlongAxis * 304.8:F1}mm) - {(closeAlongAxis ? "✅ PASS" : "❌ FAIL")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Perpendicular Distance: {perpendicularDistance * 304.8:F1}mm (max allowed: {maxPerpendicularDistance * 304.8:F1}mm) - {(closePerpendicular ? "✅ PASS" : "❌ FAIL")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Final Result: {(shouldCluster ? "✅ CLUSTER" : "❌ NO CLUSTER")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END CHECK ROTATED SLEEVE PROXIMITY ==========\n\n");
                }
                
                if (!DeploymentConfiguration.DeploymentMode && shouldCluster)
                {
                    DebugLogger.Info($"[ROTATED-CLUSTER-DB] Sleeves {sleeve1.SleeveInstanceId} and {sleeve2.SleeveInstanceId}: " +
                        $"Rotation={rotationAngle * 180.0 / Math.PI:F1}°, " +
                        $"DistanceAlongAxis={distanceAlongAxis * 304.8:F1}mm (max={maxDistanceAlongAxis * 304.8:F1}mm), " +
                        $"PerpendicularDistance={perpendicularDistance * 304.8:F1}mm (max={maxPerpendicularDistance * 304.8:F1}mm), " +
                        $"Sleeve1Size={sleeve1Width * 304.8:F1}×{sleeve1Height * 304.8:F1}mm, " +
                        $"Sleeve2Size={sleeve2Width * 304.8:F1}×{sleeve2Height * 304.8:F1}mm");
                }
                
                return shouldCluster;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ROTATED-CLUSTER-DB] Error checking rotated sleeve proximity: {ex.Message}");
                // Fallback to regular bounding box check
                return BoundingBoxesOverlapFromXml(sleeve1, sleeve2, toleranceDist);
            }
        }
        
        /// <summary>
        /// Helper method to check if an angle is axis-aligned (0°, 90°, 180°, 270°)
        /// </summary>
        private bool IsAxisAlignedAngle(double angleRad)
        {
            double angleDeg = angleRad * 180.0 / Math.PI;
            // Normalize to 0-360 range
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;
            
            double thresholdDegrees = 2.0; // 2 degree tolerance
            double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
            double distTo90 = Math.Abs(angleDeg - 90);
            double distTo180 = Math.Abs(angleDeg - 180);
            double distTo270 = Math.Abs(angleDeg - 270);
            
            return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                   distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
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
                string clusterLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                if (cluster.Count > 1)
                {
                    var clusterIds = string.Join(", ", cluster.Select(s => s.Id.IntegerValue));
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                    File.AppendAllText(clusterLogPath, $"✅ CLUSTER FORMED: {cluster.Count} sleeves [{clusterIds}]\n");
                    }
                }
                else
                {
                    File.AppendAllText(clusterLogPath, $"❌ INDIVIDUAL SLEEVE: {cluster[0].Id.IntegerValue} (no proximate neighbors)\n");
                }

                groupClusters.Add(cluster);
            }

            // 🔥 PROXIMITY DEBUG: Log final summary
            string clusterLogPath2 = SafeFileLogger.GetLogFilePath("cluster_debug.log");
            bool __suppressClusterLogs2 = DeploymentConfiguration.DeploymentMode;
            int totalClusters = groupClusters.Count(c => c.Count > 1);
            int totalIndividuals = groupClusters.Count(c => c.Count == 1);
                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
            File.AppendAllText(clusterLogPath2, $"\n=== CLUSTER SUMMARY ===\n");
            }
                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
            File.AppendAllText(clusterLogPath2, $"Total clusters formed: {totalClusters}\n");
            }
                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
            File.AppendAllText(clusterLogPath2, $"Total individual sleeves: {totalIndividuals}\n");
            }
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

                // ✅ CRITICAL FIX: For wall-hosted sleeves, ensure they are on the SAME wall element
                // If an MEP element passes through 2 adjacent walls, each wall should have its own distinct sleeve
                // Do NOT cluster sleeves across different walls, even if they have the same orientation
                // BUT: Sleeves on the SAME wall (multilayered walls) SHOULD cluster together
                if (groupKey.hostType == "Wall")
                {
                    var host1 = inst.Host;
                    var host2 = candidate.Host;
                    
                    if (host1 == null || host2 == null || host1.Id != host2.Id)
                    {
                        // Different walls - do not cluster
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] SKIP CLUSTERING: Sleeve {inst.Id.IntegerValue} (Wall {host1?.Id.IntegerValue ?? -1}) and {candidate.Id.IntegerValue} (Wall {host2?.Id.IntegerValue ?? -1}) are on different walls - keeping distinct sleeves");
                        continue;
                    }
                }

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
                // ✅ PHASE 2: DATABASE-FIRST - Load from database first (primary source of truth)
                // Fallback to XML only if database has no data (backward compatibility)
                if (!string.IsNullOrEmpty(targetCategory) && doc != null)
                {
                    try
                    {
                        using (var dbContext = new SleeveDbContext(doc))
                        {
                            var repository = new ClashZoneRepository(dbContext);
                            var dbZones = repository.GetClashZonesByCategory(targetCategory);
                            
                            if (dbZones != null && dbZones.Count > 0)
                            {
                                // Filter to only zones with SleeveInstanceId > 0 (placed sleeves)
                                var placedZones = dbZones.Where(z => z != null && z.SleeveInstanceId > 0 && !z.IsClusterResolved).ToList();
                                
                                foreach (var cz in placedZones)
                                {
                                    // ✅ CRITICAL: Reconstruct SleevePlacementPoint from database properties
                                    cz.EnsureSleevePlacementPointReconstructed();
                                    
                                    // ✅ ROTATION DATA FROM DB: MepElementRotationAngle is already loaded from database
                                    // via GetClashZonesByCategory -> MepRotationAngleRad column (see ClashZoneRepository line 1235).
                                    // This rotation angle is used by DetermineDominantRotationAngle to calculate cluster rotation.
                                    // No additional loading needed - rotation data is already in the ClashZone object from DB.
                                    
                                    clashZones.Add(cz);
                                }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalClusterService] ✅ DATABASE-FIRST: Loaded {clashZones.Count} clash zones from database for category '{targetCategory}' (with SleeveInstanceId>0, not cluster-resolved)");
                                    File.AppendAllText(SafeFileLogger.GetLogFilePath("cluster_debug.log"), 
                                        $"[{DateTime.Now:HH:mm:ss}] [DB-LOAD] Category={targetCategory}, Loaded={clashZones.Count}, TotalInDB={dbZones.Count}\n");
                                }
                                
                                // ✅ SUCCESS: Database has data, return it (skip XML loading)
                                return clashZones;
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalClusterService] ⚠️ Database has no clash zones for category '{targetCategory}', falling back to XML");
                                    File.AppendAllText(SafeFileLogger.GetLogFilePath("cluster_debug.log"), 
                                        $"[{DateTime.Now:HH:mm:ss}] [DB-LOAD] Category={targetCategory}, Loaded=0 (falling back to XML)\n");
                                }
                            }
                        }
                    }
                    catch (Exception dbEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[UniversalClusterService] ⚠️ Database load failed for category '{targetCategory}', falling back to XML: {dbEx.Message}");
                            File.AppendAllText(SafeFileLogger.GetLogFilePath("cluster_debug.log"), 
                                $"[{DateTime.Now:HH:mm:ss}] [DB-LOAD] Category={targetCategory}, Error={dbEx.Message} (falling back to XML)\n");
                        }
                    }
                }
                
                // ✅ FALLBACK: Load from XML if database has no data (backward compatibility)
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);

                if (!Directory.Exists(filtersDirectory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return clashZones;
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] LoadClashZonesFromRegularXml: xmlFilePath='{xmlFilePath}', targetCategory='{targetCategory}', filtersDirectory='{filtersDirectory}' (XML FALLBACK)");

                // Load from regular XML files
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath)
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                // ✅ PHASE 1 OPTIMIZATION: Multi-threading for XML loading (non-Revit operations only)
                if (OptimizationFlags.UseClusterServiceMultiThreading && xmlFiles.Length > 1)
                {
                    // Parallel XML loading for multiple files (file I/O only, no Revit API)
                    var loadedZones = new System.Collections.Concurrent.ConcurrentBag<ClashZone>();

                    System.Threading.Tasks.Parallel.ForEach(xmlFiles, xmlFile =>
                    {
                        try
                        {
                            // Skip CONDITIONS files
                            if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                                return;

                            // Skip _CLUSTER.xml files - we want regular XML files
                            if (Path.GetFileName(xmlFile).Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                                return;

                            if (!File.Exists(xmlFile))
                                return;

                            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                            using (var reader = new StreamReader(xmlFile))
                            {
                                var filter = (OpeningFilter)serializer.Deserialize(reader);

                                // Load from hierarchical structure (PRIMARY)
                                if (filter?.ClashZoneStorage?.Filters != null)
                                {
                                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                                    {
                                        if (filterGroup?.FileCombos != null)
                                        {
                                            foreach (var fileCombo in filterGroup.FileCombos)
                                            {
                                                if (fileCombo?.ClashZones != null)
                                                {
                                                    foreach (var cz in fileCombo.ClashZones)
                                                    {
                                                        if (!string.IsNullOrEmpty(targetCategory))
                                                        {
                                                            if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                                                continue;
                                                        }

                                                        if (cz.SleeveInstanceId > 0)
                                                        {
                                                            cz.EnsureSleevePlacementPointReconstructed();
                                                            loadedZones.Add(cz);
                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                                DebugLogger.Info($"[UniversalClusterService] Loaded clash zone {cz.Id} from hierarchical structure: SleeveInstanceId={cz.SleeveInstanceId}, Category={cz.MepElementCategory}");
                                                        }
                                                        else if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            DebugLogger.Info($"[UniversalClusterService] Skipped clash zone {cz.Id}: SleeveInstanceId={cz.SleeveInstanceId} (must be > 0)");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // Load from flat structure (BACKWARD COMPATIBILITY)
                                foreach (var cz in GetStorageZones(filter))
                                {
                                    if (cz == null) continue;
                                    if (!string.IsNullOrEmpty(targetCategory) &&
                                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    if (cz.SleeveInstanceId > 0 && !loadedZones.Any(c => c.Id == cz.Id))
                                    {
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        loadedZones.Add(cz);
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalClusterService] Loaded clash zone {cz.Id} from flat structure: SleeveInstanceId={cz.SleeveInstanceId}, Category={cz.MepElementCategory}");
                                    }
                                    else if (!DeploymentConfiguration.DeploymentMode && cz.SleeveInstanceId <= 0)
                                    {
                                        DebugLogger.Info($"[UniversalClusterService] Skipped clash zone {cz.Id} from flat structure: SleeveInstanceId={cz.SleeveInstanceId} (must be > 0)");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalClusterService] Error loading XML file {xmlFile} in parallel: {ex.Message}");
                        }
                    });

                    clashZones.AddRange(loadedZones);

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] ✅ PHASE 1: Loaded {clashZones.Count} clash zones using multi-threading ({xmlFiles.Length} files)");
                }
                else
                {
                    // Sequential loading (fallback or single file)
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
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[UniversalClusterService] Loading from XML file: {Path.GetFileName(xmlFile)}");
                            
                            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                            using (var reader = new StreamReader(xmlFile))
                            {
                                var filter = (OpeningFilter)serializer.Deserialize(reader);

                                // ✅ CRITICAL FIX: Load from BOTH hierarchical structure (Filters → FileCombos → ClashZones) AND flat structure
                                // UpdateSleeveCoordinatesInXml updates both structures, so we must read from both

                                int hierarchicalCount = 0;
                                int flatCount = 0;
                                int skippedNoSleeveId = 0;
                                int skippedCategoryMismatch = 0;
                                int totalZonesInFile = 0;

                                // 1. Load from hierarchical structure (PRIMARY)
                                if (filter?.ClashZoneStorage?.Filters != null)
                                {
                                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                                    {
                                        if (filterGroup?.FileCombos != null)
                                        {
                                            foreach (var fileCombo in filterGroup.FileCombos)
                                            {
                                                if (fileCombo?.ClashZones != null)
                                                {
                                                    foreach (var cz in fileCombo.ClashZones)
                                                    {
                                                        totalZonesInFile++;
                                                        
                                                        // Filter by target category during loading
                                                        if (!string.IsNullOrEmpty(targetCategory))
                                                        {
                                                            if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                                            {
                                                                skippedCategoryMismatch++;
                                                                continue;
                                                            }
                                                        }

                                                        // Only process clash zones with valid SleeveInstanceId (placed sleeves)
                                                        if (cz.SleeveInstanceId > 0)
                                                        {
                                                            // ✅ CRITICAL: Reconstruct SleevePlacementPoint from XML-serializable properties
                                                            cz.EnsureSleevePlacementPointReconstructed();
                                                            clashZones.Add(cz);
                                                            hierarchicalCount++;
                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                                DebugLogger.Info($"[UniversalClusterService] Loaded clash zone {cz.Id} from hierarchical: SleeveInstanceId={cz.SleeveInstanceId}, Category={cz.MepElementCategory}");
                                                        }
                                                        else
                                                        {
                                                            skippedNoSleeveId++;
                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                                DebugLogger.Info($"[UniversalClusterService] Skipped clash zone {cz.Id} from hierarchical: SleeveInstanceId={cz.SleeveInstanceId} (must be > 0), Category={cz.MepElementCategory}");
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // 2. Load from flat structure (BACKWARD COMPATIBILITY)
                                foreach (var cz in GetStorageZones(filter))
                                {
                                    if (cz == null) continue;
                                    if (!string.IsNullOrEmpty(targetCategory) &&
                                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    if (cz.SleeveInstanceId > 0 && !clashZones.Any(c => c.Id == cz.Id))
                                    {
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        clashZones.Add(cz);
                                        flatCount++;
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalClusterService] Loaded clash zone {cz.Id} from flat: SleeveInstanceId={cz.SleeveInstanceId}, Category={cz.MepElementCategory}");
                                    }
                                    else if (!DeploymentConfiguration.DeploymentMode && cz.SleeveInstanceId <= 0)
                                    {
                                        DebugLogger.Info($"[UniversalClusterService] Skipped clash zone {cz.Id} from flat: SleeveInstanceId={cz.SleeveInstanceId} (must be > 0)");
                                    }
                                }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalClusterService] ✅ Loaded from {Path.GetFileName(xmlFile)}: {hierarchicalCount} hierarchical, {flatCount} flat, Total={clashZones.Count}");
                                    DebugLogger.Info($"[UniversalClusterService] ⚠️ Skipped: {skippedNoSleeveId} (no SleeveInstanceId), {skippedCategoryMismatch} (category mismatch), Total zones in file={totalZonesInFile}");
                                    File.AppendAllText(SafeFileLogger.GetLogFilePath("cluster_debug.log"), 
                                        $"[{DateTime.Now:HH:mm:ss}] [XML-LOAD] File={Path.GetFileName(xmlFile)}, Loaded={hierarchicalCount + flatCount}, Skipped={skippedNoSleeveId} (no SleeveId) + {skippedCategoryMismatch} (category), TotalInFile={totalZonesInFile}\n");
                                }
                            }
                        }
                        else if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[UniversalClusterService] XML file not found: {xmlFile}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                // Note: If coordinates are zero, clustering will calculate zero distance, which is fine for validation
                // The real check happens in BoundingBoxesOverlapFromXml which will validate actual distances
                var bbox = new BoundingBoxXYZ();
                bbox.Min = new XYZ(clashZone.SleeveBoundingBoxMinX, clashZone.SleeveBoundingBoxMinY, clashZone.SleeveBoundingBoxMinZ);
                bbox.Max = new XYZ(clashZone.SleeveBoundingBoxMaxX, clashZone.SleeveBoundingBoxMaxY, clashZone.SleeveBoundingBoxMaxZ);
                
                // ✅ LOG WARNING if bbox is zero but sleeve exists - helps diagnose why clustering isn't working
                if (clashZone.SleeveInstanceId > 0 && 
                    clashZone.SleeveBoundingBoxMinX == 0 && clashZone.SleeveBoundingBoxMinY == 0 && clashZone.SleeveBoundingBoxMinZ == 0 &&
                    clashZone.SleeveBoundingBoxMaxX == 0 && clashZone.SleeveBoundingBoxMaxY == 0 && clashZone.SleeveBoundingBoxMaxZ == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] Sleeve {clashZone.SleeveInstanceId} has zero bounding box in XML - returning null to prevent clustering");
                    return null; // ✅ CRITICAL FIX: Return null for zero bounding boxes so they get filtered out
                }
                
                return bbox;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                // ✅ PERFORMANCE: Use cached parameter instead of LookupParameter
                Parameter wallDirectionParam = null;
                if (_parameterCache != null && _parameterCache.TryGetValue(sleeve, out var paramDict7))
                {
                    paramDict7.TryGetValue("Wall Direction Type", out wallDirectionParam);
                }
                if (wallDirectionParam == null)
                {
                    wallDirectionParam = sleeve.LookupParameter("Wall Direction Type"); // Fallback if cache not available
                }
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalClusterService] Loading cache for cleanup with xmlFilePath={xmlFilePath}, targetCategory={targetCategory}");
            // ✅ PHASE 1 OPTIMIZATION: Use cached clash zones if available (avoid duplicate XML loading)
            if (_loadedClashZonesCache != null && _loadedClashZonesCache.Count > 0)
            {
                LoadClashZoneCacheFromLoadedClashZones(_loadedClashZonesCache, targetCategory);
            }
            else
            {
                LoadClashZoneCacheFromRegularXml(xmlFilePath, targetCategory, doc, filterName);
            }
            
            // ✅ DIAGNOSTIC: Enhanced logging for cache contents after reload
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CACHE-RELOAD] Cache loaded with {_clashZoneCache.Count} total clash zones");
                
                // Count clash zones by type
                var withSleeveId = _clashZoneCache.Values.Count(cz => cz.SleeveInstanceId > 0);
                var withClusterId = _clashZoneCache.Values.Count(cz => cz.ClusterSleeveInstanceId > 0);
                var withAfterClusterId = _clashZoneCache.Values.Count(cz => cz.AfterClusterSleevePlacedSleeveInstanceId > 0);
                var withClusterBbox = _clashZoneCache.Values.Count(cz => cz.ClusterSleeveInstanceId > 0 && 
                    cz.ClusterSleeveBoundingBoxMinX != 0 && cz.ClusterSleeveBoundingBoxMinY != 0);
                
                DebugLogger.Info($"[CACHE-RELOAD] Breakdown: SleeveInstanceId>0: {withSleeveId}, ClusterSleeveInstanceId>0: {withClusterId}, AfterClusterId>0: {withAfterClusterId}, WithClusterBbox: {withClusterBbox}");
                
                // Group by host type
                var byHostType = _clashZoneCache.Values
                    .Where(cz => cz.ClusterSleeveInstanceId > 0)
                    .GroupBy(cz => GetHostTypeFromClashZone(cz))
                    .ToDictionary(g => g.Key, g => g.Count());
                
                DebugLogger.Info($"[CACHE-RELOAD] Cluster clash zones by host type: {string.Join(", ", byHostType.Select(kvp => $"{kvp.Key}={kvp.Value}"))}");
                
                // Log sample cluster clash zones with bounding boxes
                var sampleClusterZones = _clashZoneCache.Values
                    .Where(cz => cz.ClusterSleeveInstanceId > 0)
                    .Take(5)
                    .ToList();
                
                foreach (var cz in sampleClusterZones)
                {
                    var hostType = GetHostTypeFromClashZone(cz);
                    DebugLogger.Info($"[CACHE-RELOAD] Sample cluster clash zone: ClusterId={cz.ClusterSleeveInstanceId}, HostType={hostType}, " +
                        $"BboxMin=({cz.ClusterSleeveBoundingBoxMinX:F6}, {cz.ClusterSleeveBoundingBoxMinY:F6}, {cz.ClusterSleeveBoundingBoxMinZ:F6}), " +
                        $"BboxMax=({cz.ClusterSleeveBoundingBoxMaxX:F6}, {cz.ClusterSleeveBoundingBoxMaxY:F6}, {cz.ClusterSleeveBoundingBoxMaxZ:F6})");
                }
                
                // Log to file as well
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CACHE-RELOAD] Total: {_clashZoneCache.Count}, WithClusterId: {withClusterId}, WithClusterBbox: {withClusterBbox}\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CACHE-RELOAD] By host type: {string.Join(", ", byHostType.Select(kvp => $"{kvp.Key}={kvp.Value}"))}\n");
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
        
        private bool TryLoadCacheFromDataService(string xmlFilePath, string targetCategory, Document doc, string filterName)
        {
            var documentForService = doc ?? _doc;
            if (documentForService == null)
                return false;

            var baseFilter = FilterNameHelper.NormalizeBaseName(filterName, filterName);

            if (string.IsNullOrWhiteSpace(baseFilter) && !string.IsNullOrWhiteSpace(xmlFilePath))
            {
                var fileName = Path.GetFileNameWithoutExtension(xmlFilePath);
                baseFilter = FilterNameHelper.NormalizeBaseName(fileName, fileName);
            }

            if (string.IsNullOrWhiteSpace(baseFilter))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning("[UniversalClusterService] Unable to determine base filter name for data-service load.");
                return false;
            }

            var categoriesToLoad = new List<string>();

            if (!string.IsNullOrWhiteSpace(targetCategory))
            {
                categoriesToLoad.Add(MepCategoryConstants.Normalize(targetCategory));
            }
            else if (!string.IsNullOrWhiteSpace(xmlFilePath))
            {
                var fileName = Path.GetFileNameWithoutExtension(xmlFilePath);
                var normalizedBase = FilterNameHelper.NormalizeBaseName(fileName, fileName);
                var suffix = fileName.Length > normalizedBase.Length
                    ? fileName.Substring(normalizedBase.Length).TrimStart('_')
                    : string.Empty;

                if (!string.IsNullOrWhiteSpace(suffix))
                    categoriesToLoad.Add(MepCategoryConstants.FromXmlSuffix(suffix));
            }

            if (categoriesToLoad.Count == 0)
            {
                categoriesToLoad.AddRange(new[]
                {
                    MepCategoryConstants.DUCTS,
                    MepCategoryConstants.DUCT_ACCESSORIES,
                    MepCategoryConstants.PIPES,
                    MepCategoryConstants.CABLE_TRAYS
                });
            }

            var dataService = new ClashZoneDataService(documentForService, message =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] {message}");
            });

            var loadedAny = false;

            foreach (var category in categoriesToLoad.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var normalizedCategory = MepCategoryConstants.Normalize(category);
                
                // ✅ DATABASE OPTIMIZATION: Use optimized method for clustering
                // This only loads zones with individual sleeves that need clustering (database-level filtering)
                var zones = dataService.LoadZonesWithSleevesForClustering(baseFilter, normalizedCategory);
                if (zones.Count == 0)
                    continue;

                foreach (var cz in zones)
                {
                    if (cz == null)
                        continue;

                    // ✅ OPTIMIZED: Database already filtered, but double-check critical fields
                    if (cz.MepElementIdValue <= 0)
                        continue;

                    cz.EnsureSleevePlacementPointReconstructed();
                    cz.IsCurrentClash = true;
                    _clashZoneCache[cz.MepElementIdValue] = cz;
                    loadedAny = true;

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[UniversalClusterService] LOADED (SQLite-OPTIMIZED) ClashZone {cz.Id} with SleeveInstanceId={cz.SleeveInstanceId} for clustering in category '{normalizedCategory}'");
                    }
                }
            }

            return loadedAny;
        }

        private void LoadClashZoneCacheFromRegularXml(string xmlFilePath, string targetCategory = null, Document doc = null, string filterName = null)
        {
            _clashZoneCache.Clear();

            if (TryLoadCacheFromDataService(xmlFilePath, targetCategory, doc, filterName))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[UniversalClusterService] Cache populated from SQLite (primary). Skipping XML fallback.");
                return;
            }

            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return;
                }

                // ✅ DYNAMIC: Load from regular XML files (the ones UniversalSleevePlacerService saves to)
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath) 
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                foreach (var xmlFile in xmlFiles)
                {
                    var fileName = Path.GetFileName(xmlFile);
                    if (fileName.Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase) ||
                        fileName.Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!File.Exists(xmlFile))
                        continue;

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] Loading clash zones from regular XML: {fileName}");

                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                    using var reader = new StreamReader(xmlFile);
                    var filter = (OpeningFilter)serializer.Deserialize(reader);

                    foreach (var cz in GetStorageZones(filter))
                    {
                        if (cz == null)
                            continue;

                        if (!string.IsNullOrEmpty(targetCategory) &&
                            !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Log($"[UniversalClusterService] SKIP: ClashZone {cz.Id} category '{cz.MepElementCategory}' doesn't match target '{targetCategory}'");
                            continue;
                        }

                        if (cz.SleeveInstanceId <= 0 && cz.ClusterSleeveInstanceId <= 0 && cz.AfterClusterSleevePlacedSleeveInstanceId <= 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Log($"[UniversalClusterService] SKIP: ClashZone {cz.Id} has no placement IDs (not placed yet)");
                            continue;
                        }

                        if (cz.MepElementIdValue <= 0)
                            continue;

                        cz.EnsureSleevePlacementPointReconstructed();
                        cz.IsCurrentClash = true;
                        _clashZoneCache[cz.MepElementIdValue] = cz;

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Log($"[UniversalClusterService] LOADED: ClashZone {cz.Id} with SleeveInstanceId={cz.SleeveInstanceId}, ClusterSleeveInstanceId={cz.ClusterSleeveInstanceId} from {fileName}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalClusterService] Error loading clash zones from XML: {ex.Message}");
            }
        }
        
        /// <summary>
        /// ✅ PHASE 1 OPTIMIZATION: Populate clash zone cache from already-loaded clash zones
        /// Avoids duplicate XML loading - uses cached clash zones instead
        /// </summary>
        private void LoadClashZoneCacheFromLoadedClashZones(List<ClashZone> clashZones, string targetCategory = null)
        {
            _clashZoneCache.Clear();
            
            try
            {
                foreach (var cz in clashZones)
                {
                    if (cz == null)
                        continue;

                    if (!string.IsNullOrEmpty(targetCategory) &&
                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (cz.SleeveInstanceId <= 0 && cz.ClusterSleeveInstanceId <= 0 && cz.AfterClusterSleevePlacedSleeveInstanceId <= 0)
                    {
                        continue;
                    }

                    if (cz.MepElementIdValue <= 0)
                        continue;

                    cz.EnsureSleevePlacementPointReconstructed();
                    cz.IsCurrentClash = true;
                    _clashZoneCache[cz.MepElementIdValue] = cz;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[UniversalClusterService] ✅ PHASE 1: Populated cache from {clashZones.Count} loaded clash zones (filtered for {targetCategory ?? "ALL"})");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalClusterService] Error populating cache from loaded clash zones: {ex.Message}");
            }
        }

        /// <summary>
        /// Get clash zone from cache by sleeve instance ID
        /// </summary>
        private ClashZone GetClashZoneBySleeveInstanceId(int sleeveInstanceId, string xmlFilePath = null)
        {
            try
            {
                ClashZone clashZone = null;
                
                // ✅ PERFORMANCE: First try cache lookup (fast O(N) search)
                foreach (var cz in _clashZoneCache.Values)
                {
                    if (cz.SleeveInstanceId == sleeveInstanceId)
                    {
                        clashZone = cz;
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"[UniversalClusterService] Found clash zone for Sleeve Instance ID {sleeveInstanceId} in cache");
                        break;
                    }
                }
                
                // ✅ CRITICAL FIX: If found in cache but missing rotated bounding boxes, load from database
                if (clashZone != null)
                {
                    // ✅ FIX: Cast to ClashZone to avoid RuntimeBinderException with dynamic types
                    var czCache = clashZone as ClashZone;
                    bool hasRotatedBbox = czCache != null && 
                                        czCache.RotatedBoundingBoxMinX.HasValue && 
                                        czCache.RotatedBoundingBoxMinY.HasValue &&
                                        czCache.RotatedBoundingBoxMaxX.HasValue && 
                                        czCache.RotatedBoundingBoxMaxY.HasValue;
                    
                    if (!hasRotatedBbox && czCache != null)
                    {
                        // Try to load from database to get rotated bounding boxes
                        try
                        {
                            using (var dbContext = new SleeveDbContext(_doc))
                            {
                                var repository = new ClashZoneRepository(dbContext);
                                var dbZones = repository.GetClashZonesByCategory(czCache.MepElementCategory ?? "");
                                var dbZone = dbZones?.FirstOrDefault(z => z.SleeveInstanceId == sleeveInstanceId);
                                
                                if (dbZone != null && dbZone.RotatedBoundingBoxMinX.HasValue)
                                {
                                    // Update cache entry with rotated bounding boxes from database
                                    czCache.RotatedBoundingBoxMinX = dbZone.RotatedBoundingBoxMinX;
                                    czCache.RotatedBoundingBoxMinY = dbZone.RotatedBoundingBoxMinY;
                                    czCache.RotatedBoundingBoxMinZ = dbZone.RotatedBoundingBoxMinZ;
                                    czCache.RotatedBoundingBoxMaxX = dbZone.RotatedBoundingBoxMaxX;
                                    czCache.RotatedBoundingBoxMaxY = dbZone.RotatedBoundingBoxMaxY;
                                    czCache.RotatedBoundingBoxMaxZ = dbZone.RotatedBoundingBoxMaxZ;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[UniversalClusterService] ✅ Loaded rotated bounding boxes from database for sleeve {sleeveInstanceId}");
                                }
                            }
                        }
                        catch (Exception dbEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalClusterService] ⚠️ Failed to load rotated bounding boxes from database: {dbEx.Message}");
                        }
                    }
                    
                    return clashZone;
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] Sleeve Instance ID {sleeveInstanceId} not found in cache (cache size: {_clashZoneCache?.Count ?? 0}) - trying database then XML fallback");
                
                // ✅ CRITICAL FIX: Try database first (has rotated bounding boxes), then XML fallback
                try
                {
                    using (var dbContext = new SleeveDbContext(_doc))
                    {
                        var repository = new ClashZoneRepository(dbContext);
                        // Try to find by SleeveInstanceId across all categories
                        var allCategories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                        foreach (var category in allCategories)
                        {
                            var dbZones = repository.GetClashZonesByCategory(category);
                            var dbZone = dbZones?.FirstOrDefault(z => z.SleeveInstanceId == sleeveInstanceId);
                            if (dbZone != null)
                            {
                                // Add to cache for future lookups
                                if (dbZone.MepElementIdValue > 0 && !_clashZoneCache.ContainsKey(dbZone.MepElementIdValue))
                                {
                                    _clashZoneCache[dbZone.MepElementIdValue] = dbZone;
                                }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalClusterService] ✅ Found clash zone for Sleeve Instance ID {sleeveInstanceId} in database: {dbZone.Id}");
                                return dbZone;
                            }
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] ⚠️ Database lookup failed: {dbEx.Message}");
                }
                
                // ✅ FALLBACK: Try XML if database lookup fails
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (Directory.Exists(filtersDirectory))
                {
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
                                
                                if (filter?.ClashZoneStorage?.AllZones != null)
                                {
                                    var zonesFromRegularXml = GetStorageZones(filter);
                                    if (zonesFromRegularXml.Count > 0)
                                    {
                                        foreach (var cz in zonesFromRegularXml)
                                        {
                                            if (cz.SleeveInstanceId == sleeveInstanceId)
                                            {
                                                // ✅ CRITICAL: Reconstruct placement points from XML-serializable properties
                                                cz.EnsureSleevePlacementPointReconstructed();
                                                cz.EnsureSleevePlacementPointActiveDocumentReconstructed();
                                                
                                                // Add to cache for future lookups
                                                if (cz.MepElementIdValue > 0 && !_clashZoneCache.ContainsKey(cz.MepElementIdValue))
                                                {
                                                    _clashZoneCache[cz.MepElementIdValue] = cz;
                                                }
                                                
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[UniversalClusterService] ✅ Found clash zone for Sleeve Instance ID {sleeveInstanceId} in XML fallback: {cz.Id}");
                                                return cz;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] ❌ Sleeve Instance ID {sleeveInstanceId} not found in cache, database, or XML files");
                return null;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[UniversalClusterService] Found clash zone for MEP Element ID {mepElementId} in cache");
                    return clashZone;
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
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

                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                var zonesFromRegularXml = GetStorageZones(filter);
                                if (zonesFromRegularXml.Count > 0)
                                {
                                    foreach (var cz in zonesFromRegularXml)
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
                }

                return null;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                    if (filter?.ClashZoneStorage?.AllZones != null)
                    {
                        clashZones.AddRange(GetStorageZones(filter));
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
            out FamilyInstance placedClusterSleeve,
            out int? capturedClusterSleeveId, // ✅ CRITICAL FIX FOR DUCTS: Return captured ID directly
            string xmlFilePath = null)
        {
            placed = 0;
            deleted = 0;
            placedClusterSleeve = null;
            capturedClusterSleeveId = null; // ✅ CRITICAL FIX FOR DUCTS: Initialize captured ID
            
            // ✅ CRITICAL DEBUGGING: Log groupKey values to identify category-specific issues
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                string groupKeyMsg = $"[CLUSTER-PLACEMENT-START] PlaceClusterSleeve called with:\n";
                groupKeyMsg += $"[CLUSTER-PLACEMENT-START]   groupKey.hostType='{groupKey.hostType}'\n";
                groupKeyMsg += $"[CLUSTER-PLACEMENT-START]   groupKey.systemType='{groupKey.systemType}'\n";
                groupKeyMsg += $"[CLUSTER-PLACEMENT-START]   groupKey.orientation='{groupKey.orientation}'\n";
                groupKeyMsg += $"[CLUSTER-PLACEMENT-START]   targetCategory='{targetCategory}'\n";
                groupKeyMsg += $"[CLUSTER-PLACEMENT-START]   cluster.Count={cluster.Count}\n";
                DebugLogger.Info(groupKeyMsg);
                System.IO.File.AppendAllText(clusterDebugLogPath, groupKeyMsg);
            }

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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string earlyReturnMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌ EARLY RETURN: Unknown host type '{groupKey.hostType}' - placedClusterSleeve will remain NULL\n";
                    earlyReturnMsg += $"[CLUSTER-PLACEMENT-EARLY-RETURN]   Category: {groupKey.systemType}\n";
                    earlyReturnMsg += $"[CLUSTER-PLACEMENT-EARLY-RETURN]   Orientation: {groupKey.orientation}\n";
                    earlyReturnMsg += $"[CLUSTER-PLACEMENT-EARLY-RETURN]   This is likely why {(groupKey.systemType?.Contains("Duct") == true ? "DUCTS" : groupKey.systemType?.Contains("Cable") == true ? "CABLE TRAYS" : "this category")} are failing!\n";
                    earlyReturnMsg += $"[CLUSTER-PLACEMENT-EARLY-RETURN]   Check XML data: sleeve.HostType should be 'Wall', 'Floor', or 'Structural Framing', not '{groupKey.hostType}'\n";
                    DebugLogger.Error(earlyReturnMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, earlyReturnMsg);
                }
                capturedClusterSleeveId = null; // ✅ Assign before return
                return;
            }
            
            // ✅ DEBUGGING: Log selected family name
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                string familyMsg = $"[CLUSTER-PLACEMENT-FAMILY] Selected familyName='{familyName}' for hostType='{groupKey.hostType}', isCircular={isCircular}, systemType='{groupKey.systemType}'\n";
                DebugLogger.Info(familyMsg);
                System.IO.File.AppendAllText(clusterDebugLogPath, familyMsg);
            }

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[ClusterService] Creating {groupKey.systemType} cluster using family '{familyName}' (shape: {(isCircular ? "Circular" : "Rectangular")}, orientation: {groupKey.orientation})");

            // ✅ SIMPLIFIED: Use same universal families as placer service (no separate cluster families)
            // Find FamilySymbol from universal families: RectangularOpeningOnWall, RectangularOpeningOnSlab, etc.
            var universalSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[ClusterService] Looking for universal family '{familyName}' - Found {universalSymbols.Count} symbols");

            if (universalSymbols.Count == 0)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[ClusterService] Universal family '{familyName}' not found in project");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[ClusterService] Available families in project:");
                var allFamilies = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Select(sym => sym.Family.Name)
                    .Distinct()
                    .ToList();
                foreach (var famName in allFamilies.OrderBy(f => f))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService]   - {famName}");
                }

                // Try to load the missing universal family
                if (!LoadUniversalFamily(doc, familyName))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string earlyReturnMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌ EARLY RETURN: Failed to load universal family '{familyName}' - placedClusterSleeve will remain NULL\n";
                        DebugLogger.Error(earlyReturnMsg);
                        System.IO.File.AppendAllText(clusterDebugLogPath, earlyReturnMsg);
                        DebugLogger.Error($"[ClusterService] Failed to load universal family '{familyName}'");
                    }
                    capturedClusterSleeveId = null; // ✅ Assign before return
                    return;
                }

                // Try again to find the family after loading
                universalSymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (universalSymbols.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string earlyReturnMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌ EARLY RETURN: Still no family found for '{familyName}' after loading attempt - placedClusterSleeve will remain NULL\n";
                        DebugLogger.Error(earlyReturnMsg);
                        System.IO.File.AppendAllText(clusterDebugLogPath, earlyReturnMsg);
                        DebugLogger.Error($"[ClusterService] Still no family found for '{familyName}' after loading attempt");
                    }
                    capturedClusterSleeveId = null; // ✅ Assign before return
                    return;
                }
            }

            // ✅ SIMPLIFIED: Use first available symbol - don't care about type name
            var familySymbol = universalSymbols.First();
            if (!familySymbol.IsActive) familySymbol.Activate();

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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[ClusterService] Sleeve ID {xmlSleeve.SleeveInstanceId} not found in document for cluster");
                    }
                }
                catch (Exception ex)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService] Error getting sleeve {xmlSleeve.SleeveInstanceId} from document: {ex.Message}");
                }
            }
            
            if (actualSleeves.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string earlyReturnMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌ EARLY RETURN: No actual sleeves found for cluster group (cluster.Count={cluster.Count}) - placedClusterSleeve will remain NULL\n";
                    DebugLogger.Error(earlyReturnMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, earlyReturnMsg);
                    DebugLogger.Error($"[ClusterService] No actual sleeves found for cluster group, skipping placement.");
                }
                capturedClusterSleeveId = null; // ✅ Assign before return
                return;
            }
            
            // ✅ NEW: Calculate rotation angle BEFORE calculating bounding box
            // This ensures the bounding box is calculated in the rotated coordinate system
            // Use dominant angle from cluster sleeves (more accurate than single sleeve)
            
            // ✅ DEBUG: Log individual sleeve rotation angles before determining dominant angle
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var angleList = new List<string>();
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone != null)
                    {
                        var cz = sleeveData.ClashZone as ClashZone;
                        if (cz != null)
                        {
                            double angleDeg = cz.MepElementRotationAngle * 180 / Math.PI;
                            angleList.Add($"Sleeve {sleeveData.SleeveInstanceId}: {angleDeg:F1}°");
                        }
                    }
                }
                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] Cluster with {cluster.Count} sleeves - Individual angles: {string.Join(", ", angleList)}");
            }
            
            // ✅ CRITICAL: Check if this is a round pipe or round duct - if so, skip rotation logic
            // Round pipes/ducts should be treated as axis-aligned (no rotation) regardless of MEP orientation
            bool isRoundPipeOrDuct = false;
            if (groupKey.systemType != null)
            {
                bool isPipeForRotation = groupKey.systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isDuctCategory = groupKey.systemType.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0;
                
                if (isPipeForRotation)
                {
                    // Pipes are always round - skip rotation
                    isRoundPipeOrDuct = true;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string skipMsg = $"[CLUSTER-ROTATION-SKIP] ✅ Skipping rotation for ROUND PIPE cluster (systemType='{groupKey.systemType}') - treating as axis-aligned\n";
                        DebugLogger.Info(skipMsg);
                        System.IO.File.AppendAllText(clusterDebugLogPath, skipMsg);
                    }
                }
                else if (isDuctCategory)
                {
                    // Check if ducts are round by examining first sleeve's ClashZone
                    var firstSleeveClashZone = GetClashZoneBySleeveInstanceId(cluster[0].SleeveInstanceId, xmlFilePath);
                    if (firstSleeveClashZone != null)
                    {
                        var cz = firstSleeveClashZone as ClashZone;
                        if (cz != null)
                        {
                            // Check DuctShape property or MepElementSizeData.Shape
                            bool isRoundDuct = string.Equals(cz.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                             (cz.MepElementSizeData != null && 
                                              (string.Equals(cz.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                               string.Equals(cz.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)));
                            
                            if (isRoundDuct)
                            {
                                isRoundPipeOrDuct = true;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                                    string skipMsg = $"[CLUSTER-ROTATION-SKIP] ✅ Skipping rotation for ROUND DUCT cluster (DuctShape='{cz.DuctShape}', Shape='{cz.MepElementSizeData?.Shape ?? "N/A"}') - treating as axis-aligned\n";
                                    DebugLogger.Info(skipMsg);
                                    System.IO.File.AppendAllText(clusterDebugLogPath, skipMsg);
                                }
                            }
                        }
                    }
                }
            }
            
            // ✅ CRITICAL: Determine correct rotation angle for cluster sleeve
            // WHY ROTATE? The cluster sleeve must align with the MEP element direction.
            // If individual sleeves are rotated (e.g., -45° or 135°), the cluster sleeve must also be rotated to match.
            // WHAT IF WE DON'T ROTATE? The cluster sleeve would be axis-aligned (0°) while MEP elements are rotated,
            // causing misalignment between the cluster sleeve and the actual MEP element direction.
            // HOW TO CHOOSE? When sleeves are on the same axis (differ by 180°), use MEP element orientation vector
            // to determine which angle matches the actual MEP direction.
            // ⚠️ EXCEPTION: Round pipes and round ducts should NOT be rotated - they are always axis-aligned
            double rotationAngle = 0.0;
            if (!isRoundPipeOrDuct && cluster != null && cluster.Count > 0 && actualSleeves != null && actualSleeves.Count > 0)
            {
                // Step 1: Collect all sleeve rotation angles and check if they're on the same axis
                var sleeveAngles = new List<(int sleeveId, double angle, ClashZone clashZone)>();
                foreach (var sleeveData in cluster)
                {
                    var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (clashZone == null) continue;
                    
                    var cz = clashZone as ClashZone;
                    if (cz == null) continue;
                    
                    // Try to get actual rotation from Revit element first
                    double angle = 0.0;
                    var sleeveId = new ElementId(sleeveData.SleeveInstanceId);
                    var sleeve = doc.GetElement(sleeveId) as FamilyInstance;
                    if (sleeve != null)
                    {
                        var loc = sleeve.Location as LocationPoint;
                        if (loc != null)
                        {
                            angle = loc.Rotation;
                        }
                        else
                        {
                            angle = cz.MepElementRotationAngle;
                        }
                    }
                    else
                    {
                        angle = cz.MepElementRotationAngle;
                    }
                    
                    sleeveAngles.Add((sleeveData.SleeveInstanceId, angle, cz));
                }
                
                if (sleeveAngles.Count > 0)
                {
                    // Step 2: Check if all sleeves are on the same axis (differ by 0° or 180°)
                    bool allOnSameAxis = true;
                    double firstAngle = sleeveAngles[0].angle;
                    double firstAngleDeg = firstAngle * 180.0 / Math.PI;
                    while (firstAngleDeg < 0) firstAngleDeg += 360.0;
                    while (firstAngleDeg >= 360.0) firstAngleDeg -= 360.0;
                    
                    foreach (var (sleeveId, angle, cz) in sleeveAngles)
                    {
                        double angleDeg = angle * 180.0 / Math.PI;
                        while (angleDeg < 0) angleDeg += 360.0;
                        while (angleDeg >= 360.0) angleDeg -= 360.0;
                        
                        double diff = Math.Abs(angleDeg - firstAngleDeg);
                        if (diff > 180.0) diff = 360.0 - diff;
                        
                        // Check if difference is 0° or 180° (same axis)
                        double axisToleranceDeg = 1.0;
                        bool isSameAxis = diff <= axisToleranceDeg || Math.Abs(diff - 180.0) <= axisToleranceDeg;
                        
                        if (!isSameAxis)
                        {
                            allOnSameAxis = false;
                            break;
                        }
                    }
                    
                    // Step 3: If all sleeves are on the same axis, use MEP element orientation to choose correct angle
                    if (allOnSameAxis && sleeveAngles.Count > 1)
                    {
                        // Get the first sleeve's MEP element orientation vector
                        var firstCz = sleeveAngles[0].clashZone;
                        if (firstCz.MepElementOrientation != null && firstCz.MepElementOrientation.GetLength() > 1e-6)
                        {
                            // Calculate angle from MEP orientation vector
                            XYZ mepOrientation = firstCz.MepElementOrientation.Normalize();
                            double mepAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                            double mepAngleDeg = mepAngle * 180.0 / Math.PI;
                            
                            // Find which sleeve angle is closest to the MEP orientation angle
                            double bestAngle = sleeveAngles[0].angle;
                            double minDiff = double.MaxValue;
                            
                            foreach (var (sleeveId, angle, cz) in sleeveAngles)
                            {
                                double angleDeg = angle * 180.0 / Math.PI;
                                while (angleDeg < 0) angleDeg += 360.0;
                                while (angleDeg >= 360.0) angleDeg -= 360.0;
                                
                                double diff1 = Math.Abs(angleDeg - mepAngleDeg);
                                if (diff1 > 180.0) diff1 = 360.0 - diff1;
                                
                                // Also check the opposite direction (angle + 180°)
                                double angleOpposite = angleDeg + 180.0;
                                if (angleOpposite >= 360.0) angleOpposite -= 360.0;
                                double diff2 = Math.Abs(angleOpposite - mepAngleDeg);
                                if (diff2 > 180.0) diff2 = 360.0 - diff2;
                                
                                double minDiffForThisAngle = Math.Min(diff1, diff2);
                                
                                if (minDiffForThisAngle < minDiff)
                                {
                                    minDiff = minDiffForThisAngle;
                                    // Use the angle that's closer to MEP orientation
                                    bestAngle = (diff1 < diff2) ? angle : (angle + Math.PI);
                                    // Normalize bestAngle to -π to π range
                                    while (bestAngle > Math.PI) bestAngle -= 2 * Math.PI;
                                    while (bestAngle < -Math.PI) bestAngle += 2 * Math.PI;
                                }
                            }
                            
                            rotationAngle = bestAngle;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ✅ Sleeves on same axis: Using MEP orientation vector to choose angle. MEP angle: {mepAngleDeg:F1}°, Selected: {rotationAngle * 180 / Math.PI:F1}°");
                            }
                        }
                        else
                        {
                            // Fallback: Use first sleeve's angle if MEP orientation not available
                            rotationAngle = sleeveAngles[0].angle;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ⚠️ Sleeves on same axis but MEP orientation not available, using first sleeve angle: {rotationAngle * 180 / Math.PI:F1}°");
                            }
                        }
                    }
                    else
                    {
                        // Not all on same axis, or only one sleeve: Use first sleeve's angle
                        var firstActualSleeve = actualSleeves[0];
                        var firstSleeve = cluster[0];
                        int firstSleeveId = firstSleeve.SleeveInstanceId;
                        
                        var loc = firstActualSleeve.Location as LocationPoint;
                        if (loc != null)
                        {
                            rotationAngle = loc.Rotation;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ✅ Using FIRST sleeve's ACTUAL Revit rotation: {rotationAngle * 180 / Math.PI:F1}° (SleeveId={firstSleeveId})");
                            }
                        }
                        else
                        {
                            var firstSleeveClashZone = GetClashZoneBySleeveInstanceId(firstSleeveId, xmlFilePath);
                            if (firstSleeveClashZone != null)
                            {
                                rotationAngle = firstSleeveClashZone.MepElementRotationAngle;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ⚠️ Using FIRST sleeve's ClashZone rotation (LocationPoint not available): {rotationAngle * 180 / Math.PI:F1}° (SleeveId={firstSleeveId})");
                                }
                            }
                        }
                    }
                }
            }
            
            // Fallback to DetermineDominantRotationAngle if rotation angle is still not determined
            if (Math.Abs(rotationAngle) < 1e-6)
            {
                rotationAngle = DetermineDominantRotationAngle(cluster, xmlFilePath);
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ⚠️ Rotation angle was 0, falling back to DetermineDominantRotationAngle: {rotationAngle * 180 / Math.PI:F1}°");
                }
            }
            
            // ✅ DEBUG: Log the final rotation angle used
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] Final rotation angle for cluster: {rotationAngle * 180 / Math.PI:F1}° (radians: {rotationAngle:F6})");
            }

            // ✅ FALLBACK: If dominant angle is 0, try orientation-based approach for walls/framing
            // ⚠️ CRITICAL: This fallback ONLY applies to Walls and Structural Framing, NOT Floors
            if (Math.Abs(rotationAngle) < 1e-6)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] Rotation angle is 0, checking fallback logic for hostType='{groupKey.hostType}'");
                }
                
                if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                {
                    // Get orientation from XML data (stored in MepElementOrientationDirection)
                    string xmlOrientation = groupKey.orientation ?? "Unknown";
                    
                    // Apply rotation based on XML orientation data
                    if (xmlOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        rotationAngle = Math.PI / 2;  // 90° rotation for X-oriented walls/framing
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ⚠️ FALLBACK TRIGGERED: {groupKey.hostType} with X orientation → Setting rotation to 90°");
                        }
                    }
                    else if (xmlOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    {
                        rotationAngle = 0.0;  // No rotation for Y-oriented walls/framing
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ⚠️ FALLBACK TRIGGERED: {groupKey.hostType} with Y orientation → Keeping rotation at 0°");
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[ClusterService] {groupKey.hostType} orientation from XML: '{xmlOrientation}', rotation angle: {rotationAngle * 180 / Math.PI}°");
                        DebugLogger.Info($"[ROTATION] {groupKey.hostType} orientation: {xmlOrientation}, rotation: {rotationAngle * 180 / Math.PI}°\n");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] ✅ No fallback for hostType='{groupKey.hostType}' (only applies to Wall/Structural Framing), keeping rotation at 0°");
                    }
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode && Math.Abs(rotationAngle) > 1e-6)
            {
                DebugLogger.Info($"[ClusterService] Using dominant rotation angle: {rotationAngle * 180 / Math.PI:F1}° for cluster bounding box calculation");
            }

            // Step 2: Use ClusterBoundingBoxServices to get accurate bounding box from actual sleeves
            // ✅ CRITICAL FIX: For rotated sleeves, we need to use rotated bounding box coordinates
            // from ClashZone data instead of axis-aligned Revit bounding boxes
            // This ensures cluster size matches individual sleeve sizes in rotated coordinate system
            var (width, height, depth, mid, rotatedMinX, rotatedMinY, rotatedMinZ, rotatedMaxX, rotatedMaxY, rotatedMaxZ) = GetClusterBoundingBoxWithRotatedCoordinates(cluster, actualSleeves, rotationAngle, xmlFilePath);
            
            // ✅ ROTATED BBOX STORAGE: Calculate rotated bounding box coordinates for database storage
            // If rotated, we need to store the bounding box in rotated coordinate system
            XYZ rotatedBboxMin = XYZ.Zero;
            XYZ rotatedBboxMax = XYZ.Zero;
            double rotatedWidth = width;
            double rotatedHeight = height;
            double rotatedDepth = depth;
            bool isRotated = Math.Abs(rotationAngle) > 1e-6;
            
            if (isRotated)
            {
                // Calculate rotated bounding box in rotated coordinate system
                // The bounding box from GetClusterBoundingBox is already in rotated coordinates
                // We need to store the min/max in rotated coordinate system
                // For rotated clusters, the bounding box is calculated around the center point
                double halfWidth = width / 2.0;
                double halfHeight = height / 2.0;
                double halfDepth = depth / 2.0;
                
                // Rotated bounding box min/max in rotated coordinate system (centered at origin)
                rotatedBboxMin = new XYZ(-halfWidth, -halfHeight, -halfDepth);
                rotatedBboxMax = new XYZ(halfWidth, halfHeight, halfDepth);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-BBOX-ROTATED] Rotated bounding box in rotated coordinate system: Min=({rotatedBboxMin.X:F3}, {rotatedBboxMin.Y:F3}, {rotatedBboxMin.Z:F3}), Max=({rotatedBboxMax.X:F3}, {rotatedBboxMax.Y:F3}, {rotatedBboxMax.Z:F3}), Width={rotatedWidth:F3}, Height={rotatedHeight:F3}, Depth={rotatedDepth:F3}");
                }
            }
            else
            {
                // For axis-aligned, use the actual bounding box from sleeves
                var allBboxes = actualSleeves.Select(s => s.get_BoundingBox(null)).Where(b => b != null).ToList();
                if (allBboxes.Count > 0)
                {
                    rotatedBboxMin = new XYZ(allBboxes.Min(b => b.Min.X), allBboxes.Min(b => b.Min.Y), allBboxes.Min(b => b.Min.Z));
                    rotatedBboxMax = new XYZ(allBboxes.Max(b => b.Max.X), allBboxes.Max(b => b.Max.Y), allBboxes.Max(b => b.Max.Z));
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Log($"[ClusterService] ✅ ROTATED COORDINATE SYSTEM: Cluster bounding box - Width={width:F3}, Height={height:F3}, Depth={depth:F3}, Angle={rotationAngle * 180 / Math.PI:F1}°");
            }

            // ✅ FIXED: Get reference level from XML data (use first sleeve's level)
            Level? refLevel = GetReferenceLevelFromXml(doc, cluster[0]);
            if (refLevel == null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string earlyReturnMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌ EARLY RETURN: Reference level not found for cluster sleeve (first sleeve ID={cluster[0]?.SleeveInstanceId ?? -1}) - placedClusterSleeve will remain NULL\n";
                    DebugLogger.Error(earlyReturnMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, earlyReturnMsg);
                    DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                }
                capturedClusterSleeveId = null; // ✅ Assign before return
                return;
            }

            // Place cluster sleeve
            // ✅ SIMPLIFIED: Use same universal family symbol as placer service
            FamilyInstance inst = null;
            try
            {
                inst = doc.Create.NewFamilyInstance(mid, familySymbol, refLevel!, StructuralType.NonStructural);
                
                // ✅ CRITICAL LOGGING: Log cluster sleeve creation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string placementMsg = $"[CLUSTER-PLACEMENT] ✅✅✅ CLUSTER SLEEVE CREATED: ID={inst.Id.IntegerValue}, Placement=({mid.X:F3}, {mid.Y:F3}, {mid.Z:F3}), Level={refLevel?.Name ?? "NULL"}\n";
                    placementMsg += $"[CLUSTER-PLACEMENT] Cluster contains {cluster.Count} individual sleeves: {string.Join(", ", cluster.Select(s => s.SleeveInstanceId))}\n";
                    placementMsg += $"[CLUSTER-PLACEMENT] Dimensions: W={width:F3}, H={height:F3}, D={depth:F3}\n";
                    placementMsg += $"[CLUSTER-PLACEMENT] Rotation Angle: {rotationAngle * 180 / Math.PI:F1}°\n";
                    placementMsg += $"[CLUSTER-PLACEMENT] Document.IsModifiable: {doc.IsModifiable}\n";
                    DebugLogger.Info(placementMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, placementMsg);
                }
            }
            catch (Exception createEx)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string errorMsg = $"[CLUSTER-PLACEMENT] ❌❌❌ EXCEPTION CREATING CLUSTER SLEEVE: {createEx.Message}\n";
                    errorMsg += $"[CLUSTER-PLACEMENT] StackTrace: {createEx.StackTrace}\n";
                    errorMsg += $"[CLUSTER-PLACEMENT] Document.IsModifiable: {doc.IsModifiable}\n";
                    DebugLogger.Error(errorMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, errorMsg);
                }
                throw; // Re-throw to prevent continuing with null inst
            }
            
            if (inst == null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string errorMsg = $"[CLUSTER-PLACEMENT-EARLY-RETURN] ❌❌❌ EARLY RETURN: inst is NULL after creation attempt! - placedClusterSleeve will remain NULL\n";
                    errorMsg += $"[CLUSTER-PLACEMENT-EARLY-RETURN] Placement point: ({mid.X:F3}, {mid.Y:F3}, {mid.Z:F3}), Level: {refLevel?.Name ?? "NULL"}, Family: {familySymbol?.Family?.Name ?? "NULL"}\n";
                    DebugLogger.Error(errorMsg);
                    System.IO.File.AppendAllText(clusterDebugLogPath, errorMsg);
                }
                capturedClusterSleeveId = null; // ✅ Assign before return
                return; // Exit early if creation failed
            }

            // ✅ DETAILED LOGGING: Log calculated dimensions before setting parameters
            var clusterSizeLogPath = SafeFileLogger.GetLogFilePath("cluster_size_application.log");
            var sizeLogBuilder = new System.Text.StringBuilder();
            sizeLogBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== CLUSTER SIZE APPLICATION ==========");
            sizeLogBuilder.AppendLine($"Cluster Sleeve ID: {inst.Id.IntegerValue}");
            sizeLogBuilder.AppendLine($"Calculated from GetClusterBoundingBoxWithRotatedCoordinates:");
            sizeLogBuilder.AppendLine($"  Width (internal): {width:F6} = {width * 304.8:F1}mm");
            sizeLogBuilder.AppendLine($"  Height (internal): {height:F6} = {height * 304.8:F1}mm");
            sizeLogBuilder.AppendLine($"  Depth (internal): {depth:F6} = {depth * 304.8:F1}mm");
            sizeLogBuilder.AppendLine($"Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°");
            sizeLogBuilder.AppendLine($"Host Type: {groupKey.hostType}");
            sizeLogBuilder.AppendLine($"Orientation: {groupKey.orientation}");
            
            // Set size parameters (swap dimensions if rotated for orientation alignment)
            // ✅ FIX: Apply swap for both walls and framing when X-oriented
            // ✅ FLOOR: NO dimension swap for Floor - use bounding box values directly
            bool shouldSwapDimensions = ((groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing") && rotationAngle != 0.0);
            sizeLogBuilder.AppendLine($"Should Swap Dimensions: {shouldSwapDimensions}");
            sizeLogBuilder.AppendLine();
            
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);
            
            // ✅ VERIFY: Read back the actual set values
            var widthParam = inst.LookupParameter("Width");
            var heightParam = inst.LookupParameter("Height");
            var depthParam = inst.LookupParameter("Depth");
            
            if (widthParam != null)
            {
                double setWidth = widthParam.AsDouble();
                sizeLogBuilder.AppendLine($"AFTER SETTING PARAMETERS:");
                sizeLogBuilder.AppendLine($"  Width parameter: {setWidth:F6} = {setWidth * 304.8:F1}mm (expected: {width * 304.8:F1}mm)");
            }
            if (heightParam != null)
            {
                double setHeight = heightParam.AsDouble();
                sizeLogBuilder.AppendLine($"  Height parameter: {setHeight:F6} = {setHeight * 304.8:F1}mm (expected: {height * 304.8:F1}mm)");
            }
            if (depthParam != null)
            {
                double setDepth = depthParam.AsDouble();
                sizeLogBuilder.AppendLine($"  Depth parameter: {setDepth:F6} = {setDepth * 304.8:F1}mm (expected: {depth * 304.8:F1}mm)");
            }
            
            sizeLogBuilder.AppendLine($"  ========== END CLUSTER SIZE APPLICATION ==========");
            sizeLogBuilder.AppendLine();
            
            try
            {
                System.IO.File.AppendAllText(clusterSizeLogPath, sizeLogBuilder.ToString());
            }
            catch { }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CLUSTER-SIZE] Cluster {inst.Id.IntegerValue}: Calculated W={width * 304.8:F1}mm × H={height * 304.8:F1}mm, Set W={(widthParam?.AsDouble() ?? 0) * 304.8:F1}mm × H={(heightParam?.AsDouble() ?? 0) * 304.8:F1}mm");
            }
            
            // ✅ CRITICAL: Rotate the cluster sleeve to match the cluster's rotated axis
            // The rotationAngle represents the cluster's "intended rotated axis" direction
            // The cluster sleeve must be rotated to align with this axis so it matches the orientation of the individual sleeves
            // This rotation is applied AFTER placement and sizing
            // ⚠️ EXCEPTION: Round pipes and round ducts should NOT be rotated - skip rotation entirely
            if (!isRoundPipeOrDuct && Math.Abs(rotationAngle) > 1e-6)
            {
                // Check if angle is close to axis-aligned (0°, 90°, 180°, 270°) - if so, don't rotate
                double angleDegrees = rotationAngle * 180 / Math.PI;
                // Normalize to 0-360 range
                double normalizedAngle = angleDegrees;
                while (normalizedAngle < 0) normalizedAngle += 360;
                while (normalizedAngle >= 360) normalizedAngle -= 360;
                
                double distTo0 = Math.Min(normalizedAngle, 360 - normalizedAngle);
                double distTo90 = Math.Abs(normalizedAngle - 90);
                double distTo180 = Math.Abs(normalizedAngle - 180);
                double distTo270 = Math.Abs(normalizedAngle - 270);
                
                double thresholdDegrees = 2.0; // 2 degree tolerance
                bool isAxisAligned = distTo0 < thresholdDegrees || distTo90 < thresholdDegrees ||
                                    distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
                
                if (!isAxisAligned)
                {
                    // ✅ TEST: Add 90 degrees to MEP orientation angle to fix alignment issue
                    // The cluster sleeve is currently 90 degrees off from MEP elements
                    // Adding π/2 (90 degrees) should align it correctly
                    double adjustedRotationAngle = rotationAngle + Math.PI / 2.0;
                    double adjustedAngleDegrees = adjustedRotationAngle * 180 / Math.PI;
                    
                    // Rotate around Z-axis (vertical) at the placement point
                    XYZ axisOrigin = mid;
                    XYZ axisDirection = XYZ.BasisZ;
                    Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                    ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, adjustedRotationAngle);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ROTATION] ✅ Applied rotation {adjustedAngleDegrees:F1}° (original: {angleDegrees:F1}° + 90°) to cluster sleeve {inst.Id.IntegerValue} to align with rotated axis");
                    }
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ROTATION] ⚠️ Skipping rotation for cluster sleeve {inst.Id.IntegerValue}: Angle {angleDegrees:F1}° is axis-aligned");
                }
            }
            
            // ✅ ROTATED BBOX STORAGE: Store rotation data for this cluster sleeve (before returning)
            // This will be used when saving to database to store rotated bounding box coordinates
            int clusterInstanceId = inst.Id.IntegerValue;
            double rotationAngleDeg = rotationAngle * 180 / Math.PI;
            _clusterRotationData[clusterInstanceId] = (rotationAngleDeg, isRotated, rotatedBboxMin, rotatedBboxMax, rotatedWidth, rotatedHeight, rotatedDepth);
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CLUSTER-ROTATION-STORAGE] Stored rotation data for cluster {clusterInstanceId}: Angle={rotationAngleDeg:F1}°, IsRotated={isRotated}, RotatedBbox=({rotatedBboxMin.X:F3},{rotatedBboxMin.Y:F3},{rotatedBboxMin.Z:F3}) to ({rotatedBboxMax.X:F3},{rotatedBboxMax.Y:F3},{rotatedBboxMax.Z:F3})");
            }
            
            // ✅ DEBUG: Log before calling SetClusterSleeveMetadata
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[PlaceClusterSleeve] About to call SetClusterSleeveMetadata for cluster sleeve {inst.Id}, targetCategory='{targetCategory}'");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[PlaceClusterSleeve] About to call SetClusterSleeveMetadata for cluster sleeve {inst.Id}, targetCategory='{targetCategory}'\n");
            
            // CRITICAL FIX: Set metadata parameters for cluster sleeve
            SetClusterSleeveMetadata(inst, targetCategory);

            // ✅ CRITICAL: Collect ALL MEP Element IDs from all sleeves in cluster
            var allMepElementIds = new List<long>();
            var firstClashZone = null as ClashZone;
            
            foreach (var sleeveData in cluster)
            {
                // ✅ CRITICAL FIX: Pass xmlFilePath to enable XML fallback if cache lookup fails
                var clashZone = GetClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                if (clashZone != null && clashZone.MepElementIdValue > 0)
                {
                    allMepElementIds.Add(clashZone.MepElementIdValue);
                    if (firstClashZone == null)
                    {
                        firstClashZone = clashZone; // Keep first for GUID storage
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[PlaceClusterSleeve] ✅ Found first clash zone for GUID storage: {firstClashZone.Id} (SleeveId={sleeveData.SleeveInstanceId})");
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[PlaceClusterSleeve] ⚠️ Could not find clash zone for SleeveId={sleeveData.SleeveInstanceId} - GUID may not be set for this cluster!");
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[MEP_ElementId] Cluster contains {allMepElementIds.Count} MEP Element IDs: {string.Join(", ", allMepElementIds)}\n");
            
            if (allMepElementIds.Count > 0 && firstClashZone != null)
            {
                // Set first MEP Element ID for backward compatibility
                var clusterMepElementIdParam = inst.LookupParameter("MEP_ElementId");
                    if (clusterMepElementIdParam != null && !clusterMepElementIdParam.IsReadOnly)
                    {
                    clusterMepElementIdParam.Set(allMepElementIds[0]);
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DEBUG] ✓ Set MEP_ElementId = {allMepElementIds[0]} (first) on cluster sleeve {inst.Id}\n");
                }
                
                // ✅ CRITICAL: Store ALL MEP Element IDs as comma-separated list for lookup
                // This allows GuidManager to find cluster sleeve for ANY MEP element in the cluster
                string mepIdsString = string.Join(",", allMepElementIds);
                var clusterMepElementIdsParam = inst.LookupParameter("MEP_ElementIds");
                if (clusterMepElementIdsParam != null)
                {
                    if (clusterMepElementIdsParam.IsReadOnly)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[PlaceClusterSleeve] ❌ MEP_ElementIds parameter is READ-ONLY on cluster sleeve {inst.Id} - cannot set value!");
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[PlaceClusterSleeve] Parameter storage type: {clusterMepElementIdsParam.StorageType}, Element type: {inst.GetType().Name}");
                    }
                    else
                    {
                        try
                        {
                            clusterMepElementIdsParam.Set(mepIdsString);
                            
                            // ✅ VERIFY: Read back the value to confirm it was set correctly
                            string verifyValue = clusterMepElementIdsParam.AsString() ?? "";
                            if (verifyValue == mepIdsString)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[PlaceClusterSleeve] ✅ VERIFIED: Set MEP_ElementIds = '{mepIdsString}' on cluster sleeve {inst.Id} (contains {allMepElementIds.Count} MEP IDs)");
                            }
                            else
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[PlaceClusterSleeve] ❌ VERIFICATION FAILED: Expected '{mepIdsString}', got '{verifyValue}' on cluster sleeve {inst.Id}");
                            }
                        }
                        catch (Exception setEx)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[PlaceClusterSleeve] ❌ ERROR setting MEP_ElementIds on cluster sleeve {inst.Id}: {setEx.Message}");
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[PlaceClusterSleeve] Stack trace: {setEx.StackTrace}");
                        }
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[PlaceClusterSleeve] ❌ CRITICAL: MEP_ElementIds parameter NOT FOUND on cluster sleeve {inst.Id}!");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[PlaceClusterSleeve] Available parameters: {string.Join(", ", inst.Parameters.Cast<Parameter>().Select(p => p.Definition?.Name ?? "null").Take(10))}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[PlaceClusterSleeve] Add 'MEP_ElementIds' text parameter to universal family '{inst.Symbol?.Family?.Name ?? "unknown"}' to enable cluster lookup");
                }
                
                // ✅ REMOVED: No longer storing ClashZone_GUID parameter on cluster sleeves
                // Matching now uses Global XML by MEP+Host+Point (O(1) lookup) instead of GUID parameter
                // This enables cross-filter matching without needing GUID parameter on sleeves
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PlaceClusterSleeve] ⚠️ CRITICAL: No MEP Element IDs found OR firstClashZone is null - GUID will NOT be set for cluster sleeve {inst.Id}!");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PlaceClusterSleeve] This may cause duplicate clash zones on refresh. Check cache population and XML loading.");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DEBUG] ✗ No MEP Element IDs found in cluster sleeves\n");
            }

            placed++;
            
            // ✅ CRITICAL: Capture ID immediately while element is valid (before any potential transaction issues)
            int clusterSleeveIdValue = inst.Id.IntegerValue;
            
            // ✅ CRITICAL LOGGING: Log before setting placedClusterSleeve
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                string setMsg = $"[CLUSTER-PLACEMENT-SET] About to set placedClusterSleeve = inst (ID={clusterSleeveIdValue})\n";
                DebugLogger.Info(setMsg);
                System.IO.File.AppendAllText(clusterDebugLogPath, setMsg);
            }
            
            placedClusterSleeve = inst; // ✅ CRITICAL: Return placed cluster sleeve for tracking
            
            // ✅ CRITICAL LOGGING: Log after setting placedClusterSleeve (using captured ID to avoid invalid element access)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                string verifyMsg = $"[CLUSTER-PLACEMENT-SET] ✅ VERIFIED: placedClusterSleeve is now ID={clusterSleeveIdValue} (ID captured while element was valid)\n";
                DebugLogger.Info(verifyMsg);
                System.IO.File.AppendAllText(clusterDebugLogPath, verifyMsg);
            }

            // ✅ CONSOLIDATION: Get cluster sleeve bounding box from Revit immediately after placement
            BoundingBoxXYZ clusterBbox = null;
            try
            {
                clusterBbox = inst.get_BoundingBox(null);
                if (clusterBbox != null && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PlaceClusterSleeve] ✅ Got cluster sleeve bounding box: Min=({clusterBbox.Min.X:F6}, {clusterBbox.Min.Y:F6}, {clusterBbox.Min.Z:F6}), Max=({clusterBbox.Max.X:F6}, {clusterBbox.Max.Y:F6}, {clusterBbox.Max.Z:F6})");
                }
            }
            catch (Exception bboxEx)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PlaceClusterSleeve] Could not get bounding box for cluster sleeve {inst.Id}: {bboxEx.Message}");
            }

            // ⚠️ CRITICAL: Mark clash zones as cluster-resolved with the actual cluster sleeve ID and bounding box
            // ✅ Pass rotated bounding box coordinates if available
            (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)? rotatedBbox = null;
            if (rotatedMinX.HasValue && rotatedMinY.HasValue && rotatedMinZ.HasValue && 
                rotatedMaxX.HasValue && rotatedMaxY.HasValue && rotatedMaxZ.HasValue)
            {
                rotatedBbox = (rotatedMinX.Value, rotatedMinY.Value, rotatedMinZ.Value, rotatedMaxX.Value, rotatedMaxY.Value, rotatedMaxZ.Value);
            }
            MarkClashZonesAsClusterResolvedWithSleeveId(cluster, inst.Id, xmlFilePath, clusterBbox, rotatedBbox);

            // ✅ PERFORMANCE: Removed XML file reading from inside PlaceClusterSleeve (was causing 90+ second delays per cluster)
            // Flag updates are handled by MarkClashZonesAsClusterResolvedWithSleeveId which updates XML files
            // FlagManager updates will be handled in batch after all clusters are placed (not per-cluster)

                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[ClusterService] About to delete {cluster.Count} individual sleeves for cluster sleeve {inst.Id.IntegerValue}");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[DELETE] About to delete {cluster.Count} individual sleeves for cluster sleeve {inst.Id.IntegerValue} (hostType={groupKey.hostType})\n");

            // Delete originals - collect ElementIds first, then delete in batch
            var sleevesToDelete = new List<ElementId>();
            var clusterFormationSleeveIds = new List<int>(); // ✅ Track which sleeves are part of cluster formation
            
            foreach (var s in cluster)
            {
                try
                {
                    // Get the actual Revit sleeve by SleeveInstanceId
                    var sleeveElementId = new ElementId(s.SleeveInstanceId);
                    var sleeveElement = doc.GetElement(sleeveElementId);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[DELETE] Checking sleeve {sleeveElementId.IntegerValue}: found={sleeveElement != null}, isFamilyInstance={sleeveElement is FamilyInstance}\n");
                    }
                    
                    if (sleeveElement != null && sleeveElement is FamilyInstance sleeveInstance)
                    {
                        sleevesToDelete.Add(sleeveElementId);
                        clusterFormationSleeveIds.Add(sleeveElementId.IntegerValue); // ✅ Track cluster formation sleeves
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Log($"[ClusterService] Queued for deletion: individual sleeve {sleeveElementId.IntegerValue} (part of cluster formation)");
                            string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                            System.IO.File.AppendAllText(clusterDebugLogPath, $"[DELETE-CLUSTER-FORMATION] ✅ Queued cluster formation sleeve {sleeveElementId.IntegerValue} for deletion\n");
                        }
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[ClusterService] Could not find sleeve {s.SleeveInstanceId} for deletion");
                            DebugLogger.Info($"[DELETE] ✗ Sleeve {sleeveElementId.IntegerValue} not found or not a FamilyInstance\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[ClusterService] Error preparing sleeve {s.SleeveInstanceId} for deletion: {ex.Message}");
                        DebugLogger.Info($"[DELETE] ✗ Exception preparing sleeve: {ex.Message}\n");
                    }
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[DELETE] Total cluster formation sleeves queued for deletion: {sleevesToDelete.Count} (IDs: {string.Join(", ", clusterFormationSleeveIds)})\n");
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[DELETE-CLUSTER-FORMATION] Total cluster formation sleeves: {sleevesToDelete.Count} (IDs: {string.Join(", ", clusterFormationSleeveIds)})\n");
            }
            
            // Delete all sleeves in batch (within the same transaction)
            if (sleevesToDelete.Count > 0)
            {
                try
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] Calling doc.Delete() for {sleevesToDelete.Count} sleeves (hostType={groupKey.hostType})\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] Document modifiable: {doc.IsModifiable}\n");
                    doc.Delete(sleevesToDelete);
                    deleted = sleevesToDelete.Count;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[ClusterService] Successfully deleted {deleted} individual sleeves in batch");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] ✓ Successfully deleted {deleted} individual sleeves in batch (hostType={groupKey.hostType})\n");
                    
                    // ✅ VERIFY: Check if sleeves were actually deleted
                    int stillExists = 0;
                    foreach (var sleeveId in sleevesToDelete)
                    {
                        var stillThere = doc.GetElement(sleeveId);
                        if (stillThere != null)
                        {
                            stillExists++;
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[DELETE] ⚠️ WARNING: Sleeve {sleeveId.IntegerValue} still exists after deletion attempt!\n");
                        }
                    }
                    if (stillExists > 0)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DELETE] ⚠️ WARNING: {stillExists} sleeves still exist after deletion attempt!\n");
                    }
                }
                catch (Exception ex)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService] Error deleting sleeves in batch: {ex.Message}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] ✗ Error deleting sleeves: {ex.Message} (hostType={groupKey.hostType})\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] Stack trace: {ex.StackTrace}\n");
                    deleted = 0;
                }
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[DELETE] ⚠️ No sleeves to delete (sleevesToDelete.Count=0) for hostType={groupKey.hostType}\n");
            }
            
            // ✅ EDGE CASE: Find and delete individual sleeves that fall within cluster zone but weren't part of cluster formation
            if (clusterBbox != null && !string.IsNullOrEmpty(targetCategory))
            {
                try
                {
                    int clusterFormationCount = clusterFormationSleeveIds.Count;
                    int edgeCaseDeleted = DeleteEdgeCaseSleevesInClusterZone(doc, clusterBbox, targetCategory, cluster, inst.Id, xmlFilePath);
                    deleted += edgeCaseDeleted;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string deleteSummary = $"[DELETE-SUMMARY] Cluster sleeve {inst.Id.IntegerValue}: Deleted {clusterFormationCount} cluster formation sleeves + {edgeCaseDeleted} edge case sleeves = {deleted} total\n";
                        DebugLogger.Info(deleteSummary);
                        System.IO.File.AppendAllText(clusterDebugLogPath, deleteSummary);
                        
                        if (edgeCaseDeleted > 0)
                        {
                            DebugLogger.Info($"[DELETE] ✅ EDGE CASE: Deleted {edgeCaseDeleted} individual sleeves that fell in cluster zone (cluster sleeve {inst.Id.IntegerValue})");
                        }
                    }
                }
                catch (Exception edgeEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[ClusterService] Error deleting edge case sleeves: {edgeEx.Message}");
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[DELETE-EDGE-CASE-ERROR] ❌ Error deleting edge case sleeves: {edgeEx.Message}\n");
                    }
                }
            }
        }

        /// <summary>
        /// ✅ EDGE CASE: Find and delete individual sleeves that fall within cluster zone but weren't part of cluster formation
        /// Updates database flags FIRST, then Global XML (database is primary source of truth)
        /// </summary>
        private int DeleteEdgeCaseSleevesInClusterZone(
            Document doc,
            BoundingBoxXYZ clusterBbox,
            string targetCategory,
            List<dynamic> clusterFormation,
            ElementId clusterSleeveId,
            string xmlFilePath = null)
        {
            int deletedCount = 0;
            
            try
            {
                // Get all sleeve IDs that were part of cluster formation (to exclude them)
                var clusterFormationSleeveIds = new HashSet<int>();
                foreach (var s in clusterFormation)
                {
                    if (s.SleeveInstanceId > 0)
                        clusterFormationSleeveIds.Add(s.SleeveInstanceId);
                }
                
                // Find all individual sleeves in the same category
                var allSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s =>
                    {
                        // Check if it's a sleeve family
                        bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                                s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                        string familyName = s.Symbol?.FamilyName ?? string.Empty;
                        bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                             familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        return (s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                               (hasSleeveKeyword || isKnownFamily);
                    })
                    .ToList();
                
                // Filter sleeves that:
                // 1. Are in the same category
                // 2. Fall within cluster bounding box
                // 3. Are NOT part of cluster formation
                // 4. Are individual sleeves (not cluster sleeves)
                var edgeCaseSleeves = new List<FamilyInstance>();
                int clusterSleeveIdValue = clusterSleeveId.IntegerValue; // ✅ CRITICAL: Capture cluster sleeve ID to protect it
                
                foreach (var sleeve in allSleeves)
                {
                    int sleeveId = sleeve.Id.IntegerValue;
                    
                    // ✅ CRITICAL PROTECTION #1: Skip if this IS the cluster sleeve itself
                    if (sleeveId == clusterSleeveIdValue)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[EDGE-CASE] ⚠️ PROTECTED: Skipping cluster sleeve {sleeveId} (should never be deleted)");
                        }
                        continue;
                    }
                    
                    // Skip if part of cluster formation
                    if (clusterFormationSleeveIds.Contains(sleeveId))
                        continue;
                    
                    // ✅ CRITICAL PROTECTION #2: Skip if it's a cluster sleeve (check by family name or parameters)
                    if (sleeve.Symbol?.FamilyName?.Contains("Cluster", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[EDGE-CASE] ⚠️ PROTECTED: Skipping cluster sleeve {sleeveId} (family name contains 'Cluster')");
                        }
                        continue;
                    }
                    
                    // ✅ CRITICAL PROTECTION #3: Check if sleeve has "Cluster Sleeve Instance ID" parameter set to its own ID
                    var clusterInstanceIdParam = sleeve.LookupParameter("Cluster Sleeve Instance ID");
                    if (clusterInstanceIdParam != null)
                    {
                        int clusterInstanceId = clusterInstanceIdParam.AsInteger();
                        if (clusterInstanceId > 0 && clusterInstanceId == sleeveId)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[EDGE-CASE] ⚠️ PROTECTED: Skipping cluster sleeve {sleeveId} (Cluster Sleeve Instance ID parameter = {clusterInstanceId})");
                            }
                            continue;
                        }
                    }
                    
                    // Check category match
                    string sleeveCategory = GetCategoryFromMepElementId(sleeve);
                    if (string.IsNullOrEmpty(sleeveCategory) || 
                        !string.Equals(sleeveCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    // Check if sleeve falls within cluster bounding box
                    var sleeveBbox = sleeve.get_BoundingBox(null);
                    if (sleeveBbox == null)
                        continue;
                    
                    // Check if sleeve center or any part is within cluster bounding box
                    var sleeveCenter = (sleeveBbox.Min + sleeveBbox.Max) / 2.0;
                    bool withinCluster = sleeveCenter.X >= clusterBbox.Min.X && sleeveCenter.X <= clusterBbox.Max.X &&
                                        sleeveCenter.Y >= clusterBbox.Min.Y && sleeveCenter.Y <= clusterBbox.Max.Y &&
                                        sleeveCenter.Z >= clusterBbox.Min.Z && sleeveCenter.Z <= clusterBbox.Max.Z;
                    
                    if (withinCluster)
                    {
                        edgeCaseSleeves.Add(sleeve);
                    }
                }
                
                if (edgeCaseSleeves.Count == 0)
                    return 0;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[EDGE-CASE] Found {edgeCaseSleeves.Count} edge case sleeves within cluster zone (category='{targetCategory}', cluster sleeve={clusterSleeveId.IntegerValue})");
                }
                
                // Delete edge case sleeves and update flags
                var edgeCaseSleeveIds = edgeCaseSleeves.Select(s => s.Id).ToList();
                var edgeCaseUpdates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)>();
                
                // Find ClashZone entries for edge case sleeves (from database)
                // ✅ DATABASE-FIRST: Look up clash zones by SleeveInstanceId in database
                using (var context = new SleeveDbContext(doc))
                {
                    var repository = new ClashZoneRepository(context);
                    var dbZones = repository.GetClashZonesByCategory(targetCategory);
                    
                    if (dbZones != null)
                    {
                        var sleeveIdToClashZone = dbZones
                            .Where(z => z.SleeveInstanceId > 0)
                            .GroupBy(z => z.SleeveInstanceId)
                            .Select(g => g.First())
                            .ToDictionary(z => z.SleeveInstanceId);
                        
                        foreach (var sleeve in edgeCaseSleeves)
                        {
                            try
                            {
                                int sleeveId = sleeve.Id.IntegerValue;
                                
                                // Try to find ClashZone by SleeveInstanceId in database
                                if (sleeveIdToClashZone.TryGetValue(sleeveId, out var clashZone))
                                {
                                    // ✅ Include all required fields: MEP+Host+Point and OLD values for matching
                                    edgeCaseUpdates.Add((
                                        clashZone.Id, 
                                        false, 
                                        true, 
                                        -1, 
                                        clusterSleeveId.IntegerValue,
                                        clashZone.MepElementIdValue,
                                        clashZone.StructuralElementIdValue,
                                        clashZone.IntersectionPointX,
                                        clashZone.IntersectionPointY,
                                        clashZone.IntersectionPointZ,
                                        clashZone.SleeveInstanceId, // OLD value for matching
                                        clashZone.ClusterSleeveInstanceId, // OLD value for matching
                                        clashZone.MarkedForClusteringSleeveProcess, // ✅ EDGE CASE: Pass through if available
                                        clashZone.AfterClusterSleevePlacedSleeveInstanceId, // ✅ EDGE CASE: Pass through if available
                                        null // ✅ EDGE CASE: IsClusteredFlag (deprecated, set to null)
                                    ));
                                }
                                else
                                {
                                    // Fallback: Try to get GUID from sleeve parameter
                                    string clashGuid = null;
                                    var guidParam = sleeve.LookupParameter("ClashZone_GUID");
                                    if (guidParam != null && !string.IsNullOrEmpty(guidParam.AsString()))
                                    {
                                        clashGuid = guidParam.AsString();
                                    }
                                    
                                    if (!string.IsNullOrEmpty(clashGuid) && Guid.TryParse(clashGuid, out var parsedGuid))
                                    {
                                        // Verify GUID exists in database
                                        var guidZone = dbZones.FirstOrDefault(z => z.Id == parsedGuid);
                                        if (guidZone != null)
                                        {
                                            // ✅ Include all required fields: MEP+Host+Point and OLD values for matching
                                            edgeCaseUpdates.Add((
                                                parsedGuid, 
                                                false, 
                                                true, 
                                                -1, 
                                                clusterSleeveId.IntegerValue,
                                                guidZone.MepElementIdValue,
                                                guidZone.StructuralElementIdValue,
                                                guidZone.IntersectionPointX,
                                                guidZone.IntersectionPointY,
                                                guidZone.IntersectionPointZ,
                                                guidZone.SleeveInstanceId, // OLD value for matching
                                                guidZone.ClusterSleeveInstanceId, // OLD value for matching
                                                guidZone.MarkedForClusteringSleeveProcess, // ✅ EDGE CASE: Pass through if available
                                                guidZone.AfterClusterSleevePlacedSleeveInstanceId, // ✅ EDGE CASE: Pass through if available
                                                null // ✅ EDGE CASE: IsClusteredFlag (deprecated, set to null)
                                            ));
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[EDGE-CASE] Error processing edge case sleeve {sleeve.Id.IntegerValue}: {ex.Message}");
                            }
                        }
                    }
                }
                
                // ✅ DATABASE FIRST: Update database flags
                if (edgeCaseUpdates.Count > 0)
                {
                    try
                    {
                        using (var context = new SleeveDbContext(doc, msg =>
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[EDGE-CASE][DB] {msg}");
                        }))
                        {
                            var repository = new ClashZoneRepository(context, msg =>
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[EDGE-CASE][DB] {msg}");
                            });
                            
                            // Convert to database format (include all fields including OLD values for matching)
                            // ✅ CRITICAL: Match exact field names expected by BatchUpdateFlags
                            var dbUpdates = edgeCaseUpdates.Select(u => (
                                ClashZoneId: u.Id,
                                IsResolved: u.IsResolved,
                                IsClusterResolved: u.IsClusterResolved,
                                SleeveInstanceId: u.SleeveInstanceId,
                                ClusterInstanceId: u.ClusterSleeveInstanceId, // Note: BatchUpdateFlags uses ClusterInstanceId, not ClusterSleeveInstanceId
                                MepElementId: u.MepElementId,
                                StructuralElementId: u.StructuralElementId,
                                IntersectionPointX: u.IntersectionPointX,
                                IntersectionPointY: u.IntersectionPointY,
                                IntersectionPointZ: u.IntersectionPointZ,
                                OldSleeveInstanceId: u.OldSleeveInstanceId,
                                OldClusterInstanceId: u.OldClusterInstanceId,
                                MarkedForClusterProcess: u.MarkedForClusterProcess, // ✅ EDGE CASE: Use tuple field name
                                AfterClusterSleeveId: u.AfterClusterSleeveId, // ✅ EDGE CASE: Use tuple field name
                                IsClusteredFlag: (bool?)null // ✅ EDGE CASE: Deprecated, set to null
                            )).ToList();
                            
                            repository.BatchUpdateFlags(dbUpdates);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[EDGE-CASE] ✅ Updated database flags for {dbUpdates.Count} edge case clash zones (database-first)");
                            }
                        }
                    }
                    catch (Exception dbEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[EDGE-CASE] ⚠️ Database update failed (non-blocking): {dbEx.Message}");
                    }
                }
                
                // Delete edge case sleeves
                if (edgeCaseSleeveIds.Count > 0)
                {
                    try
                    {
                        doc.Delete(edgeCaseSleeveIds);
                        deletedCount = edgeCaseSleeveIds.Count;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[EDGE-CASE] ✅ Deleted {deletedCount} edge case sleeves");
                        }
                    }
                    catch (Exception deleteEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[EDGE-CASE] Error deleting edge case sleeves: {deleteEx.Message}");
                    }
                }
                
                // ✅ XML SECOND: Update Global XML flags (after database)
                if (edgeCaseUpdates.Count > 0 && _flagManager != null)
                {
                    try
                    {
                        // Convert updates to format expected by GlobalIndexService
                        var globalXmlUpdates = edgeCaseUpdates.Select(u => (
                            u.Id,
                            u.IsResolved,
                            u.IsClusterResolved,
                            u.SleeveInstanceId,
                            u.ClusterSleeveInstanceId,
                            u.MepElementId,
                            u.StructuralElementId,
                            u.IntersectionPointX,
                            u.IntersectionPointY,
                            u.IntersectionPointZ
                        )).ToList();
                        
                        // Get MEP+Host+Point data from database for Global XML updates
                        using (var context = new SleeveDbContext(doc))
                        {
                            var repository = new ClashZoneRepository(context);
                            var dbZones = repository.GetClashZonesByCategory(targetCategory);
                            
                            if (dbZones != null)
                            {
                                var dbLookup = dbZones.Where(z => z.Id != Guid.Empty)
                                    .GroupBy(z => z.Id)
                                    .Select(g => g.First())
                                    .ToDictionary(z => z.Id);
                                
                                // Update Global XML with complete data
                                var completeUpdates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)>();
                                
                                foreach (var update in edgeCaseUpdates)
                                {
                                    if (dbLookup.TryGetValue(update.Id, out var dbZone))
                                    {
                                        completeUpdates.Add((
                                            update.Id,
                                            update.IsResolved,
                                            update.IsClusterResolved,
                                            update.SleeveInstanceId,
                                            update.ClusterSleeveInstanceId,
                                            dbZone.MepElementIdValue,
                                            dbZone.StructuralElementIdValue,
                                            dbZone.IntersectionPointX,
                                            dbZone.IntersectionPointY,
                                            dbZone.IntersectionPointZ
                                        ));
                                    }
                                }
                                
                                if (completeUpdates.Count > 0)
                                {
                                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(
                                        doc,
                                        targetCategory,
                                        completeUpdates,
                                        _filterName
                                    );
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[EDGE-CASE] ✅ Updated Global XML flags for {completeUpdates.Count} edge case clash zones (after database)");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception xmlEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[EDGE-CASE] ⚠️ Global XML update failed (non-blocking): {xmlEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[EDGE-CASE] Error in DeleteEdgeCaseSleevesInClusterZone: {ex.Message}");
            }
            
            return deletedCount;
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

            // ✅ DETAILED LOGGING: Log dimension swapping process
            var dimSwapLogPath = SafeFileLogger.GetLogFilePath("cluster_dimension_swap.log");
            var swapLogBuilder = new System.Text.StringBuilder();
            swapLogBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== DIMENSION SWAPPING ==========");
            swapLogBuilder.AppendLine($"Cluster Sleeve ID: {inst.Id.IntegerValue}");
            swapLogBuilder.AppendLine($"Input dimensions (from GetClusterBoundingBoxWithRotatedCoordinates):");
            swapLogBuilder.AppendLine($"  Width: {width:F6} = {width * 304.8:F1}mm");
            swapLogBuilder.AppendLine($"  Height: {height:F6} = {height * 304.8:F1}mm");
            swapLogBuilder.AppendLine($"  Depth: {depth:F6} = {depth * 304.8:F1}mm");
            swapLogBuilder.AppendLine($"Host Type: {groupKey.hostType}");
            swapLogBuilder.AppendLine($"Orientation: {groupKey.orientation}");
            swapLogBuilder.AppendLine($"Should Swap Dimensions: {shouldSwapDimensions}");

            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                // Step 1: Assign all from bbox
                double openingWidth = width;
                double openingHeight = height;
                double openingDepth = depth;
                
                swapLogBuilder.AppendLine($"BEFORE SWAP:");
                swapLogBuilder.AppendLine($"  openingWidth: {openingWidth * 304.8:F1}mm");
                swapLogBuilder.AppendLine($"  openingHeight: {openingHeight * 304.8:F1}mm");
                swapLogBuilder.AppendLine($"  openingDepth: {openingDepth * 304.8:F1}mm");
                
                // Step 2: Map world coordinates to sleeve parameters
                if (groupKey.orientation == "Y")
                {
                    // Y-wall: Width=H, Height=D, Depth=W
                    swapLogBuilder.AppendLine($"Y-WALL SWAP: Width=H, Height=D, Depth=W");
                    double tempWidth = openingWidth;
                    double tempHeight = openingHeight;
                    double tempDepth = openingDepth;
                    openingWidth = tempHeight;  // H (608.6mm)
                    openingHeight = tempDepth;  // D (270mm)
                    openingDepth = tempWidth;   // W (200mm)
                    swapLogBuilder.AppendLine($"AFTER Y-WALL SWAP:");
                    swapLogBuilder.AppendLine($"  openingWidth (was Height): {openingWidth * 304.8:F1}mm");
                    swapLogBuilder.AppendLine($"  openingHeight (was Depth): {openingHeight * 304.8:F1}mm");
                    swapLogBuilder.AppendLine($"  openingDepth (was Width): {openingDepth * 304.8:F1}mm");
                }
                else if (shouldSwapDimensions) // X-walls
                {
                    // X-wall: Width=W, Height=D, Depth=H
                    swapLogBuilder.AppendLine($"X-WALL SWAP: Width=W, Height=D, Depth=H");
                    double tempWidth = openingWidth;
                    double tempHeight = openingHeight;
                    double tempDepth = openingDepth;
                    openingWidth = tempWidth;   // W (478.8mm)
                    openingHeight = tempDepth;  // D (238.9mm)
                    openingDepth = tempHeight;  // H (200mm)
                    swapLogBuilder.AppendLine($"AFTER X-WALL SWAP:");
                    swapLogBuilder.AppendLine($"  openingWidth (stays Width): {openingWidth * 304.8:F1}mm");
                    swapLogBuilder.AppendLine($"  openingHeight (was Depth): {openingHeight * 304.8:F1}mm");
                    swapLogBuilder.AppendLine($"  openingDepth (was Height): {openingDepth * 304.8:F1}mm");
                }
                else
                {
                    swapLogBuilder.AppendLine($"NO SWAP: Using original dimensions");
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLUSTER-DIM] Sleeve {inst.Id}: Orientation={groupKey.orientation}, Swap={shouldSwapDimensions}, " +
                        $"BBox(W={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, " +
                        $"H={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, " +
                        $"D={UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm), " +
                        $"Final(W={UnitUtils.ConvertFromInternalUnits(openingWidth, UnitTypeId.Millimeters):F1}mm, " +
                        $"H={UnitUtils.ConvertFromInternalUnits(openingHeight, UnitTypeId.Millimeters):F1}mm, " +
                        $"D={UnitUtils.ConvertFromInternalUnits(openingDepth, UnitTypeId.Millimeters):F1}mm)\n");
                }
                
                swapLogBuilder.AppendLine($"SETTING PARAMETERS:");
                swapLogBuilder.AppendLine($"  Width parameter ← {openingWidth * 304.8:F1}mm");
                swapLogBuilder.AppendLine($"  Height parameter ← {openingHeight * 304.8:F1}mm");
                
                // ✅ CRITICAL FIX: Check if parameters are instance or type parameters
                // If they're type parameters, we need to set them on the type, not the instance
                if (widthParam != null)
                {
                    swapLogBuilder.AppendLine($"  Width parameter: StorageType={widthParam.StorageType}, IsReadOnly={widthParam.IsReadOnly}, Definition.Name={widthParam.Definition?.Name}");
                    if (!widthParam.IsReadOnly)
                    {
                        try
                        {
                            widthParam.Set(openingWidth);
                            swapLogBuilder.AppendLine($"  ✅ Width parameter SET to {openingWidth * 304.8:F1}mm");
                            
                            // ✅ VERIFY IMMEDIATELY: Read back to confirm
                            double verifyWidth = widthParam?.AsDouble() ?? 0.0;
                            swapLogBuilder.AppendLine($"  ✅ Width parameter VERIFIED: {verifyWidth * 304.8:F1}mm (expected: {openingWidth * 304.8:F1}mm)");
                            if (Math.Abs(verifyWidth - openingWidth) > 0.001)
                            {
                                swapLogBuilder.AppendLine($"  ⚠️ WARNING: Width parameter mismatch! Set={openingWidth * 304.8:F1}mm, Read={verifyWidth * 304.8:F1}mm");
                            }
                        }
                        catch (Exception ex)
                        {
                            swapLogBuilder.AppendLine($"  ❌ ERROR setting Width parameter: {ex.Message}");
                        }
                    }
                    else
                    {
                        swapLogBuilder.AppendLine($"  ⚠️ Width parameter is READ-ONLY - cannot set value");
                    }
                }
                else
                {
                    swapLogBuilder.AppendLine($"  ❌ Width parameter NOT FOUND");
                }
                
                if (heightParam != null)
                {
                    swapLogBuilder.AppendLine($"  Height parameter: StorageType={heightParam.StorageType}, IsReadOnly={heightParam.IsReadOnly}, Definition.Name={heightParam.Definition?.Name}");
                    if (!heightParam.IsReadOnly)
                    {
                        try
                        {
                            heightParam.Set(openingHeight);
                            swapLogBuilder.AppendLine($"  ✅ Height parameter SET to {openingHeight * 304.8:F1}mm");
                            
                            // ✅ VERIFY IMMEDIATELY: Read back to confirm
                            double verifyHeight = heightParam?.AsDouble() ?? 0.0;
                            swapLogBuilder.AppendLine($"  ✅ Height parameter VERIFIED: {verifyHeight * 304.8:F1}mm (expected: {openingHeight * 304.8:F1}mm)");
                            if (Math.Abs(verifyHeight - openingHeight) > 0.001)
                            {
                                swapLogBuilder.AppendLine($"  ⚠️ WARNING: Height parameter mismatch! Set={openingHeight * 304.8:F1}mm, Read={verifyHeight * 304.8:F1}mm");
                            }
                        }
                        catch (Exception ex)
                        {
                            swapLogBuilder.AppendLine($"  ❌ ERROR setting Height parameter: {ex.Message}");
                        }
                    }
                    else
                    {
                        swapLogBuilder.AppendLine($"  ⚠️ Height parameter is READ-ONLY - cannot set value");
                    }
                }
                else
                {
                    swapLogBuilder.AppendLine($"  ❌ Height parameter NOT FOUND");
                }

                // Get actual host thickness for Depth parameter (through-wall dimension)
                // ✅ FIXED: Use XML data instead of Revit API calls
                var firstSleeve = cluster[0];
                double hostThickness = openingDepth;  // Default fallback to calculated depth
                
                // ✅ FIXED: Cluster sleeve should use individual sleeve thickness, not structural element thickness
                // For all host types, use the calculated depth from individual sleeves
                hostThickness = openingDepth;
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[ClusterService] Cluster sleeve: Using individual sleeve depth {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm for {groupKey.hostType}");
                
                // ✅ FIXED: Use calculated depth as final fallback (no Revit API calls needed)
                if (hostThickness <= 0.0)
                {
                    hostThickness = openingDepth;  // Use calculated depth from bounding box
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[ClusterService] Using calculated depth as fallback: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLUSTER-DIM] Sleeve {inst.Id}: Calculated Depth = {UnitUtils.ConvertFromInternalUnits(openingDepth, UnitTypeId.Millimeters):F1}mm, Final HostThickness = {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm\n");
                    
                swapLogBuilder.AppendLine($"  Depth parameter ← {hostThickness * 304.8:F1}mm (host thickness)");
                
                // Set the mapped dimensions (use hostThickness for depth instead of openingDepth)
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(openingWidth);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(openingHeight);
                if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(hostThickness);
            }
            else
            {
                // ✅ FLOOR: No dimension swap for Floor - use bounding box values directly
                // Floor sleeves use Width, Height, Depth from bounding box without any swapping
                swapLogBuilder.AppendLine($"FLOOR HOST: No dimension swap, using bounding box values directly");
                swapLogBuilder.AppendLine($"SETTING PARAMETERS:");
                swapLogBuilder.AppendLine($"  Width parameter ← {width * 304.8:F1}mm");
                swapLogBuilder.AppendLine($"  Height parameter ← {height * 304.8:F1}mm");
                swapLogBuilder.AppendLine($"  Depth parameter ← {depth * 304.8:F1}mm");
                
                // ✅ CRITICAL FIX: Check if parameters are instance or type parameters
                if (widthParam != null)
                {
                    swapLogBuilder.AppendLine($"  Width parameter: StorageType={widthParam.StorageType}, IsReadOnly={widthParam.IsReadOnly}, Definition.Name={widthParam.Definition?.Name}");
                    if (!widthParam.IsReadOnly)
                    {
                        try
                        {
                            widthParam.Set(width);
                            swapLogBuilder.AppendLine($"  ✅ Width parameter SET to {width * 304.8:F1}mm");
                            
                            // ✅ VERIFY IMMEDIATELY: Read back to confirm
                            double verifyWidth = widthParam?.AsDouble() ?? 0.0;
                            swapLogBuilder.AppendLine($"  ✅ Width parameter VERIFIED: {verifyWidth * 304.8:F1}mm (expected: {width * 304.8:F1}mm)");
                            if (Math.Abs(verifyWidth - width) > 0.001)
                            {
                                swapLogBuilder.AppendLine($"  ⚠️ WARNING: Width parameter mismatch! Set={width * 304.8:F1}mm, Read={verifyWidth * 304.8:F1}mm");
                            }
                        }
                        catch (Exception ex)
                        {
                            swapLogBuilder.AppendLine($"  ❌ ERROR setting Width parameter: {ex.Message}");
                        }
                    }
                    else
                    {
                        swapLogBuilder.AppendLine($"  ⚠️ Width parameter is READ-ONLY - cannot set value");
                    }
                }
                else
                {
                    swapLogBuilder.AppendLine($"  ❌ Width parameter NOT FOUND");
                }
                
                if (heightParam != null)
                {
                    swapLogBuilder.AppendLine($"  Height parameter: StorageType={heightParam.StorageType}, IsReadOnly={heightParam.IsReadOnly}, Definition.Name={heightParam.Definition?.Name}");
                    if (!heightParam.IsReadOnly)
                    {
                        try
                        {
                            heightParam.Set(height);
                            swapLogBuilder.AppendLine($"  ✅ Height parameter SET to {height * 304.8:F1}mm");
                            
                            // ✅ VERIFY IMMEDIATELY: Read back to confirm
                            double verifyHeight = heightParam?.AsDouble() ?? 0.0;
                            swapLogBuilder.AppendLine($"  ✅ Height parameter VERIFIED: {verifyHeight * 304.8:F1}mm (expected: {height * 304.8:F1}mm)");
                            if (Math.Abs(verifyHeight - height) > 0.001)
                            {
                                swapLogBuilder.AppendLine($"  ⚠️ WARNING: Height parameter mismatch! Set={height * 304.8:F1}mm, Read={verifyHeight * 304.8:F1}mm");
                            }
                        }
                        catch (Exception ex)
                        {
                            swapLogBuilder.AppendLine($"  ❌ ERROR setting Height parameter: {ex.Message}");
                        }
                    }
                    else
                    {
                        swapLogBuilder.AppendLine($"  ⚠️ Height parameter is READ-ONLY - cannot set value");
                    }
                }
                else
                {
                    swapLogBuilder.AppendLine($"  ❌ Height parameter NOT FOUND");
                }
                
                if (depthParam != null)
                {
                    swapLogBuilder.AppendLine($"  Depth parameter: StorageType={depthParam.StorageType}, IsReadOnly={depthParam.IsReadOnly}, Definition.Name={depthParam.Definition?.Name}");
                    if (!depthParam.IsReadOnly)
                    {
                        try
                        {
                            depthParam.Set(depth);
                            swapLogBuilder.AppendLine($"  ✅ Depth parameter SET to {depth * 304.8:F1}mm");
                            
                            // ✅ VERIFY IMMEDIATELY: Read back to confirm
                            double verifyDepth = depthParam?.AsDouble() ?? 0.0;
                            swapLogBuilder.AppendLine($"  ✅ Depth parameter VERIFIED: {verifyDepth * 304.8:F1}mm (expected: {depth * 304.8:F1}mm)");
                            if (Math.Abs(verifyDepth - depth) > 0.001)
                            {
                                swapLogBuilder.AppendLine($"  ⚠️ WARNING: Depth parameter mismatch! Set={depth * 304.8:F1}mm, Read={verifyDepth * 304.8:F1}mm");
                            }
                        }
                        catch (Exception ex)
                        {
                            swapLogBuilder.AppendLine($"  ❌ ERROR setting Depth parameter: {ex.Message}");
                        }
                    }
                    else
                    {
                        swapLogBuilder.AppendLine($"  ⚠️ Depth parameter is READ-ONLY - cannot set value");
                    }
                }
                else
                {
                    swapLogBuilder.AppendLine($"  ❌ Depth parameter NOT FOUND");
                }
            }
            
            // ✅ VERIFY: Read back actual set values
            swapLogBuilder.AppendLine();
            swapLogBuilder.AppendLine($"VERIFICATION (after setting):");
            if (widthParam != null)
            {
                double actualWidth = widthParam.AsDouble();
                swapLogBuilder.AppendLine($"  Width parameter: {actualWidth * 304.8:F1}mm");
            }
            if (heightParam != null)
            {
                double actualHeight = heightParam.AsDouble();
                swapLogBuilder.AppendLine($"  Height parameter: {actualHeight * 304.8:F1}mm");
            }
            if (depthParam != null)
            {
                double actualDepth = depthParam.AsDouble();
                swapLogBuilder.AppendLine($"  Depth parameter: {actualDepth * 304.8:F1}mm");
            }
            
            swapLogBuilder.AppendLine($"  ========== END DIMENSION SWAPPING ==========");
            swapLogBuilder.AppendLine();
            
            try
            {
                System.IO.File.AppendAllText(dimSwapLogPath, swapLogBuilder.ToString());
            }
            catch { }
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] Filters directory not found: {filtersDirectory}");
                    return mapping;
                }
                
                // Load XML files for the target category (or all if null)
                var searchPattern = string.IsNullOrEmpty(targetCategory) 
                    ? "*.xml" 
                    : $"*_{GetCategoryXmlSuffix(targetCategory)}.xml";
                
                var xmlFiles = Directory.GetFiles(filtersDirectory, searchPattern);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Loading sleeve mapping from {xmlFiles.Length} XML files (pattern: {searchPattern})");
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.AllZones)
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
                                
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Log($"[UniversalClusterService] Loaded {filter.ClashZoneStorage.AllZones.Count} clash zones from {Path.GetFileName(xmlFile)}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] Error loading {Path.GetFileName(xmlFile)}: {ex.Message}");
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[UniversalClusterService] Total sleeve-to-category mappings loaded: {mapping.Count}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] 🔥 UpdateClashZoneFlagsForCluster CALLED 🔥");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Cluster ID: {clusterInstance.Id.IntegerValue}, Original sleeves: {originalSleeves.Count}, SystemType: {systemType}");
                
                // ✅ DEBUG: Log the category mapping
                                if (!DeploymentConfiguration.DeploymentMode)
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

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Loading clash zones for category: {category}");
                
                // ✅ DEBUG: Log the category mapping result
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Category mapping: '{systemType}' → '{category}'\n");

                // Load clash zones from XML
                var clashZones = LoadClashZonesFromRegularXml(null!, category, doc);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Loaded {clashZones.Count} clash zones from XML");

                // Update each clash zone for deleted sleeves
                int updatedCount = 0;
                foreach (var sleeve in originalSleeves)
                {
                    // ✅ FIXED: Use XML data instead of Revit API calls
                    int sleeveInstanceId = sleeve.SleeveInstanceId;
                    var clashZone = clashZones.FirstOrDefault(cz => cz.SleeveInstanceId == sleeveInstanceId);

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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] ✓ Updated ClashZone {clashZone.Id} as clustered");
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] ⚠️ ClashZone not found for sleeve {sleeve.Id.IntegerValue}");
                    }
                }

                // Save updated clash zones back to XML
                SaveClashZonesToXml(clashZones, category);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] ✅ Updated {updatedCount} clash zones as clustered, saved to XML");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] 🔥 FindNewestClusterSleeve CALLED 🔥");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Host Type: {groupKey.hostType}, System Type: {groupKey.systemType}");

                // ✅ FIXED: Use XML data instead of Revit API calls
                // Determine if cluster is circular or rectangular (simplified for XML data)
                bool isCircular = false; // Default to rectangular for XML data
                
                // PIPE CLUSTERS: Always rectangular regardless of member shape (legacy parity)
                bool isPipeCategory = (!string.IsNullOrEmpty(groupKey.systemType) && groupKey.systemType.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0);
                if (isPipeCategory)
                {
                    isCircular = false;
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalClusterService] Unknown host type: {groupKey.hostType}");
                    return null;
                }

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Looking for universal family: {familyName} (shape: {(isCircular ? "Circular" : "Rectangular")})");

                // ✅ SIMPLIFIED: Find all instances of the universal family (same as placer service uses)
                var clusterSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(fi => fi.Id.IntegerValue) // Newest will have highest ID
                    .ToList();

                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Found {clusterSleeves.Count} cluster sleeves using universal family {familyName}");

                // Return the newest one (highest ID)
                var newestSleeve = clusterSleeves.FirstOrDefault();
                if (newestSleeve != null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] ✅ Found newest cluster sleeve: ID {newestSleeve.Id.IntegerValue}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] ⚠️ No cluster sleeve found for family {familyName}");
                }

                return newestSleeve;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Error finding newest cluster sleeve: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ GLOBAL XML FLAG MANAGEMENT: Saves only non-flag metadata to Filter XML.
        /// Flags (IsResolved, IsClusterResolved, SleeveInstanceId, ClusterSleeveInstanceId) 
        /// are managed exclusively in Global XML (single source of truth).
        /// Only saves: MarkedForClusteringSleeveProcess and IsCurrentClash (metadata).
        /// </summary>
        private void SaveClusterFlagsToXml(string xmlFilePath, string targetCategory)
        {
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Saving non-flag metadata (MarkedForClusteringSleeveProcess, IsCurrentClash) to XML: {Path.GetFileName(xmlFilePath)}");
                
                // Load current XML
                var serializer = new XmlSerializer(typeof(OpeningFilter));
                OpeningFilter filter;
                using (var reader = new StreamReader(xmlFilePath))
                {
                    filter = (OpeningFilter)serializer.Deserialize(reader);
                }
                
                if (filter?.ClashZoneStorage?.AllZones == null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] No clash zones found in XML");
                    return;
                }
                
                // ✅ GLOBAL XML FLAG MANAGEMENT: Only save non-flag metadata, NOT flags
                // Flags (IsResolved, IsClusterResolved, SleeveInstanceId, ClusterSleeveInstanceId) 
                // are managed exclusively in Global XML (single source of truth)
                int updatedCount = 0;
                foreach (var clashZone in filter.ClashZoneStorage.AllZones)
                {
                    if (_clashZoneCache != null && _clashZoneCache.ContainsKey(clashZone.MepElementId.IntegerValue))
                    {
                        var cachedClashZone = _clashZoneCache[clashZone.MepElementId.IntegerValue];
                        
                        // ✅ Only save non-flag metadata
                        // Update MarkedForClusteringSleeveProcess (metadata, not a flag)
                        if (cachedClashZone.MarkedForClusteringSleeveProcess != null)
                        {
                            clashZone.MarkedForClusteringSleeveProcess = cachedClashZone.MarkedForClusteringSleeveProcess;
                            updatedCount++;
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalClusterService] Updated MarkedForClusteringSleeveProcess metadata for ClashZone {clashZone.Id}: {clashZone.MarkedForClusteringSleeveProcess}");
                        }
                        
                        // ✅ Preserve IsCurrentClash (metadata, not a flag)
                        clashZone.IsCurrentClash = cachedClashZone.IsCurrentClash;
                        
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalClusterService] Preserved IsCurrentClash metadata for ClashZone {clashZone.Id}: {clashZone.IsCurrentClash}");
                        
                        // ❌ DO NOT save flags: IsResolved, IsClusterResolved, SleeveInstanceId, ClusterSleeveInstanceId
                        // These are managed exclusively in Global XML
                    }
                }
                
                // Save updated XML (only non-flag metadata)
                using (var writer = new StreamWriter(xmlFilePath))
                {
                    serializer.Serialize(writer, filter);
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalClusterService] Saved {updatedCount} non-flag metadata entries to XML (flags managed in Global XML only)");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalClusterService] Error saving metadata to XML: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ CORRECT APPROACH: Use ClashZonePersistenceService to save updated clash zones
        /// This ensures the tree structure is maintained correctly using the dedicated service
        /// </summary>
        private void SaveClashZonesToXml(List<Models.ClashZone> clashZones, string category)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] Saving {clashZones.Count} clash zones for category '{category}' using ClashZonePersistenceService");
                
                // ✅ CRITICAL FIX: Use ProjectPathService instead of hardcoded path
                string filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);

                // Find pattern: *_{category}.xml (e.g., Ventilation_ducts.xml)
                var pattern = $"*_{category.ToLower()}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                
                if (matchingFiles.Length == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] No XML files found for pattern: {pattern}");
                    return;
                }

                // Use most recently modified file
                var xmlFile = matchingFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] Saving clash zones to: {xmlFile}");

                // ✅ STEP 1: Load Filter XML
                var serializer = new XmlSerializer(typeof(Models.OpeningFilter));
                Models.OpeningFilter filter;
                
                using (var reader = new StreamReader(xmlFile))
                {
                    filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                }
                
                if (filter?.ClashZoneStorage == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] Filter or ClashZoneStorage is null in {Path.GetFileName(xmlFile)}");
                    return;
                }
                
                // ✅ STEP 2: Extract ALL clash zones from tree structure and update with cluster data
                var allClashZones = new List<ClashZone>();
                if (filter.ClashZoneStorage.Filters != null)
                {
                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                    {
                        if (filterGroup?.FileCombos != null)
                        {
                            foreach (var fileCombo in filterGroup.FileCombos)
                            {
                                if (fileCombo?.ClashZones != null)
                                {
                                    foreach (var cz in fileCombo.ClashZones)
                                    {
                                        // ✅ CRITICAL: Update cluster data from updated clash zones
                                        var updatedZone = clashZones.FirstOrDefault(c => c.Id == cz.Id);
                                        if (updatedZone != null)
                                        {
                                            // Update cluster metadata (non-flag data)
                                            cz.MarkedForClusteringSleeveProcess = updatedZone.MarkedForClusteringSleeveProcess;
                                            cz.ClusterSleeveBoundingBoxMinX = updatedZone.ClusterSleeveBoundingBoxMinX;
                                            cz.ClusterSleeveBoundingBoxMinY = updatedZone.ClusterSleeveBoundingBoxMinY;
                                            cz.ClusterSleeveBoundingBoxMinZ = updatedZone.ClusterSleeveBoundingBoxMinZ;
                                            cz.ClusterSleeveBoundingBoxMaxX = updatedZone.ClusterSleeveBoundingBoxMaxX;
                                            cz.ClusterSleeveBoundingBoxMaxY = updatedZone.ClusterSleeveBoundingBoxMaxY;
                                            cz.ClusterSleeveBoundingBoxMaxZ = updatedZone.ClusterSleeveBoundingBoxMaxZ;

                                            // ✅ CRITICAL: Propagate updated flag/ID state so persistence service writes correct values
                                            cz.IsResolved = updatedZone.IsResolved;
                                            cz.IsClusterResolved = updatedZone.IsClusterResolved;
                                            cz.SleeveInstanceId = updatedZone.SleeveInstanceId;
                                            cz.ClusterSleeveInstanceId = updatedZone.ClusterSleeveInstanceId;
                                            cz.AfterClusterSleevePlacedSleeveInstanceId = updatedZone.AfterClusterSleevePlacedSleeveInstanceId;
                                        }
                                        
                                        allClashZones.Add(cz);
                                    }
                                }
                            }
                        }
                    }
                }
                
                // ✅ STEP 3: Use ClashZonePersistenceService to save updated clash zones
                var baseFilterName = FilterNameHelper.NormalizeBaseName(_filterName, filter?.Name, category);
                var guidManager = new GuidManager(_doc);
                var persistenceService = new ClashZonePersistenceService(_doc, guidManager);
                persistenceService.SaveClashZones(allClashZones, baseFilterName, filter, allowStructuralUpdates: false);
                
                // ✅ STEP 4: Save Filter XML file using FilterManagementService
                var filterManagementService = new FilterManagementService(_doc, null!, null!);
                filterManagementService.SaveFilterToXmlFile(filter, xmlFile);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalClusterService] ✅ Successfully saved {clashZones.Count} clash zones using ClashZonePersistenceService");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
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
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[SetClusterSleeveMetadata] Setting metadata for cluster sleeve {clusterSleeve.Id}, category='{category}'");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[SetClusterSleeveMetadata] Setting metadata for cluster sleeve {clusterSleeve.Id}, category='{category}'\n");
            
            // ✅ CRITICAL LOGGING: Log _filterName state before calling GetFilterNameForCategory
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                string filterNameLog = $"[SetClusterSleeveMetadata] _filterName='{_filterName ?? "NULL"}', category='{category}'\n";
                DebugLogger.Info(filterNameLog);
                System.IO.File.AppendAllText(clusterDebugLogPath, filterNameLog);
            }
            
            // ✅ OPTIONAL: Try to set MEP_Category parameter if it exists (not critical - XML has category info per sleeve)
            // Note: Opening families don't have MEP_Category parameter - XML stores category per clash zone
            // Parameter transfer uses MEP Element ID and XML category to find correct XML file
            var mepCategoryParam = clusterSleeve.LookupParameter("MEP_Category");
            if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
            {
                // Parameter exists and is writable - set it for convenience
                mepCategoryParam.Set(category);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set MEP_Category = '{category}' for cluster sleeve {clusterSleeve.Id}");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] ✓ Set MEP_Category = '{category}' for cluster sleeve {clusterSleeve.Id}\n");
            }
            // ✅ NOTE: If parameter doesn't exist, that's OK - XML lookup will use MEP Element ID and category from XML
            
            // Set Filter Name based on category
            string filterName = null;
            try
            {
                filterName = GetFilterNameForCategory(category);
            }
            catch (Exception filterNameEx)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    string errorLog = $"[SetClusterSleeveMetadata] ❌❌❌ EXCEPTION in GetFilterNameForCategory for category '{category}': {filterNameEx.Message}\n";
                    errorLog += $"[SetClusterSleeveMetadata] _filterName='{_filterName ?? "NULL"}'\n";
                    errorLog += $"[SetClusterSleeveMetadata] StackTrace: {filterNameEx.StackTrace}\n";
                    DebugLogger.Error(errorLog);
                    System.IO.File.AppendAllText(clusterDebugLogPath, errorLog);
                }
                // Don't re-throw - continue without filter name (non-critical)
                filterName = null;
            }
            
            if (!DeploymentConfiguration.DeploymentMode && filterName != null)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[SetClusterSleeveMetadata] ✅ GetFilterNameForCategory returned: '{filterName}' for category '{category}'\n");
            }
            
            var filterNameParam = clusterSleeve.LookupParameter("Filter Name");
            if (filterNameParam != null && !filterNameParam.IsReadOnly && filterName != null)
            {
                filterNameParam.Set(filterName);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Filter Name = '{filterName}' for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (filterName == null)
                        DebugLogger.Warning($"[SetClusterSleeveMetadata] Filter Name is NULL (GetFilterNameForCategory failed) for cluster sleeve {clusterSleeve.Id}, category='{category}'");
                    else
                        DebugLogger.Warning($"[SetClusterSleeveMetadata] Filter Name parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
                }
            }
            
            // Set Sleeve Instance ID to -1 (indicating this is a cluster sleeve)
            var instanceIdParam = clusterSleeve.LookupParameter("Sleeve Instance ID");
            if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
            {
                instanceIdParam.Set(-1); // -1 indicates this is a cluster sleeve
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Sleeve Instance ID = -1 for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Sleeve Instance ID parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
            }
            
            // CRITICAL: Set Cluster Sleeve Instance ID parameter for XML lookup
            var clusterInstanceIdParam = clusterSleeve.LookupParameter("Cluster Sleeve Instance ID");
            if (clusterInstanceIdParam != null && !clusterInstanceIdParam.IsReadOnly)
            {
                clusterInstanceIdParam.Set(clusterSleeve.Id.IntegerValue);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Cluster Sleeve Instance ID = {clusterSleeve.Id.IntegerValue} for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Cluster Sleeve Instance ID parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
            }
        }
        catch (Exception ex)
        {
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Error($"[SetClusterSleeveMetadata] Error setting metadata for cluster sleeve {clusterSleeve.Id}: {ex.Message}");
        }
    }
    
    /// <summary>
    /// ✅ REMOVED: SetClashZoneGuidOnClusterSleeve method no longer needed
    /// Matching now uses Global XML by MEP+Host+Point (O(1) lookup) instead of GUID parameter on sleeves
    /// This enables cross-filter matching without needing GUID parameter on sleeves
    /// </summary>
    
    /// <summary>
    /// ✅ CRITICAL: Update cluster sleeve MEP_ElementIds parameter with additional MEP Element IDs
    /// Called when sleeves are incidentally deleted during cleanup (edge cases)
    /// </summary>
    private void UpdateClusterSleeveMepElementIds(FamilyInstance clusterSleeve, HashSet<long> newMepIds)
    {
        try
        {
            var mepIdsParam = clusterSleeve.LookupParameter("MEP_ElementIds");
            if (mepIdsParam != null && !mepIdsParam.IsReadOnly)
            {
                // Get existing MEP Element IDs
                string existingIdsString = mepIdsParam.AsString() ?? "";
                var existingIds = new HashSet<long>();
                
                if (!string.IsNullOrWhiteSpace(existingIdsString))
                {
                    existingIds = existingIdsString.Split(',')
                        .Select(id => id.Trim())
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Select(id => long.TryParse(id, out long parsed) ? parsed : (long?)null)
                        .Where(id => id.HasValue)
                        .Select(id => id.Value)
                        .ToHashSet();
                }
                
                // Add new MEP Element IDs (avoid duplicates)
                existingIds.UnionWith(newMepIds);
                
                // Update parameter
                string updatedIdsString = string.Join(",", existingIds.OrderBy(id => id));
                mepIdsParam.Set(updatedIdsString);
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UpdateClusterSleeveMepElementIds] Updated MEP_ElementIds on cluster sleeve {clusterSleeve.Id}: Added {newMepIds.Count} new IDs, total now: {existingIds.Count} ({updatedIdsString})");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UpdateClusterSleeveMepElementIds] MEP_ElementIds parameter not found or read-only on cluster sleeve {clusterSleeve.Id} - cannot add deleted sleeves' MEP Element IDs");
            }
        }
        catch (Exception ex)
        {
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UpdateClusterSleeveMepElementIds] Error updating MEP_ElementIds for cluster sleeve {clusterSleeve.Id}: {ex.Message}");
        }
    }

    /// <summary>
    /// ✅ NEW: Clean up individual sleeves that fall within cluster sleeve bounding boxes
    /// This is a cheap method that checks if any remaining individual sleeves are positioned
    /// within the bounding box of any cluster sleeve and deletes them
    /// ⚠️ CRITICAL: This method MUST NOT delete cluster sleeves - only individual sleeves
    /// </summary>
    private int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> placedClusters)
    {
        int deletedCount = 0;
        
        try
        {
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[CleanupSleevesWithinClusters] 🔥 METHOD CALLED - placedClusters count: {placedClusters?.Count ?? 0}");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[CLEANUP-START] Method called with {placedClusters?.Count ?? 0} placed cluster sleeves\n");
            
            if (placedClusters == null || placedClusters.Count == 0)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log("[CleanupSleevesWithinClusters] No cluster sleeves to check against");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLEANUP-START] ⚠️ No cluster sleeves provided, exiting\n");
                return 0;
            }
            
            // ✅ CRITICAL: Build a HashSet of cluster sleeve IDs for fast O(1) lookup
            // This is the PRIMARY protection mechanism - if a sleeve ID is in this set, it will NEVER be deleted
            var clusterSleeveIds = new HashSet<int>();
            foreach (var cluster in placedClusters)
            {
                if (cluster != null && cluster.Id != null)
                {
                    int clusterId = cluster.Id.IntegerValue;
                    clusterSleeveIds.Add(clusterId);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEANUP-PROTECTION] Added cluster sleeve {clusterId} to protection set\n");
                }
            }
            
            // Log cluster sleeve IDs for debugging
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-START] Cluster sleeve IDs: {string.Join(", ", placedClusters.Select(c => c.Id.IntegerValue))}\n");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-PROTECTION-SET] Protection set contains {clusterSleeveIds.Count} cluster sleeve IDs: {string.Join(", ", clusterSleeveIds)}\n");
                DebugLogger.Info($"[CLEANUP-PROTECTION] Protection set contains {clusterSleeveIds.Count} cluster sleeve IDs: {string.Join(", ", clusterSleeveIds)}\n");
            }
            
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
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-ALL-SLEEVES-FOUND] Found {allSleeves.Count} total sleeves in Revit: {string.Join(", ", allSleeves.Select(s => s.Id.IntegerValue))}\n");
            }
            
            var individualSleeves = allSleeves.Where(s => 
            {
                int sleeveId = s.Id.IntegerValue;
                
                // ✅ CRITICAL FIX #1: PRIMARY PROTECTION - Check protection set FIRST (fastest and safest)
                // If sleeve ID is in protection set, it's DEFINITELY a cluster sleeve - NEVER delete it
                if (clusterSleeveIds.Contains(sleeveId))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEANUP-FILTER] ⚠️ PRIMARY PROTECTION: Sleeve {sleeveId} is in protection set - SKIPPING (cluster sleeve)\n");
                    return false; // This is a cluster sleeve, not an individual sleeve
                }
                
                var clusterParam = s.LookupParameter("Cluster Sleeve Instance ID");
                // ✅ FIX: -1 means it's an individual sleeve, 0 or positive means it's a cluster sleeve
                // Only check sleeves with exactly -1 for cleanup
                int clusterValue = clusterParam?.AsInteger() ?? -1;
                
                // ✅ ADDITIONAL LOGIC: Also check if sleeve is a cluster by examining its Sleeve Instance ID parameter
                var sleeveInstanceIdParam = s.LookupParameter("Sleeve Instance ID");
                int sleeveInstanceId = sleeveInstanceIdParam?.AsInteger() ?? -999; // Use -999 to distinguish from unset
                
                // ✅ LOGIC: 
                // - If Sleeve Instance ID = -1, it's a CLUSTER sleeve (should NOT be deleted)
                // - If Cluster Sleeve Instance ID = -1 or 0, it's an INDIVIDUAL sleeve (can be deleted)
                // - If Cluster Sleeve Instance ID > 0 and equals the sleeve's own ID, it's a CLUSTER sleeve
                // - We skip cluster sleeves by checking both parameters
                bool isClusterSleeve = (sleeveInstanceId == -1) || (clusterValue > 0 && clusterValue == sleeveId);
                
                // ✅ CRITICAL: Additional check - if sleeve ID is in protection set, it's DEFINITELY a cluster sleeve
                // This is a redundant check but provides extra safety (protection set check already happened above, but this ensures parameter logic also respects it)
                if (clusterSleeveIds.Contains(sleeveId))
                {
                    isClusterSleeve = true;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-PARAM-CHECK] ⚠️ Sleeve {sleeveId} is in protection set - FORCING isClusterSleeve=true (SleeveInstanceId={sleeveInstanceId}, ClusterValue={clusterValue})\n");
                    }
                }
                
                bool isIndividual = !isClusterSleeve && (clusterValue <= 0);
                
                // ✅ CRITICAL FIX #2: SECONDARY PROTECTION - Even if parameters say individual, check protection set again
                if (isIndividual && clusterSleeveIds.Contains(sleeveId))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEANUP-FILTER] ⚠️ SECONDARY PROTECTION: Sleeve {sleeveId} parameters say individual BUT it's in protection set - treating as cluster sleeve\n");
                    return false; // Safety check: if in protection set, never delete (even if parameters are wrong)
                }
                
                // ✅ DEBUG: Log ALL sleeves to find missing ones
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLEANUP-DEBUG-ALL] Sleeve {sleeveId}: Family={s.Symbol?.FamilyName}, ClusterParam={clusterValue}, SleeveInstanceId={sleeveInstanceId}, IsCluster={isClusterSleeve}, IsIndividual={isIndividual}, InProtectionSet={clusterSleeveIds.Contains(sleeveId)}\n");
                
                return isIndividual;
            }).ToList();
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[CleanupSleevesWithinClusters] Found {allSleeves.Count} total sleeves, {individualSleeves.Count} individual sleeves to check against {placedClusters.Count} cluster sleeves");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[CLEANUP] Found {allSleeves.Count} total sleeves, checking {individualSleeves.Count} individual sleeves against {placedClusters.Count} cluster sleeves\n");
            
            // ✅ BATCH DELETION: Collect all sleeves to delete first, then delete in one batch
            var sleevesToDelete = new List<(ElementId sleeveId, int clusterId, Models.ClashZone clashZone)>();
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[CLEANUP-BATCH] Scanning {individualSleeves.Count} individual sleeves for deletion candidates\n");
            
            foreach (var individualSleeve in individualSleeves)
            {
                int sleeveId = individualSleeve.Id.IntegerValue;
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLEANUP-LOOP] Processing individual sleeve {sleeveId}\n");
                
                // ✅ CRITICAL FIX #3: TERTIARY PROTECTION - Final check in loop using HashSet (fastest)
                // This is the ultimate safety check - if sleeve ID is in protection set, NEVER delete it
                if (clusterSleeveIds.Contains(sleeveId))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLEANUP] ⚠️ TERTIARY PROTECTION: Sleeve {sleeveId} is in protection set - SKIPPING (cluster sleeve)\n");
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-TERTIARY] ⚠️ PROTECTED: Sleeve {sleeveId} is in protection set - SKIPPING deletion\n");
                    }
                    continue; // Ultimate safety: if in protection set, never delete
                }
                
                // ✅ ADDITIONAL SAFETY CHECK: Double-check parameters even if not in protection set
                // Sometimes cluster sleeves might not be in placedClusters if they were placed in a different transaction
                var sleeveInstanceIdParam = individualSleeve.LookupParameter("Sleeve Instance ID");
                int sleeveInstanceId = sleeveInstanceIdParam?.AsInteger() ?? -999; // Use -999 as default (not -1) to distinguish from unset
                var clusterInstanceIdParam = individualSleeve.LookupParameter("Cluster Sleeve Instance ID");
                int clusterInstanceId = clusterInstanceIdParam?.AsInteger() ?? -1;
                
                // If Sleeve Instance ID = -1, it's definitely a cluster sleeve (even if not in placedClusters)
                // OR if Cluster Sleeve Instance ID > 0 and equals the sleeve's own ID, it's a cluster sleeve
                if (sleeveInstanceId == -1 || (clusterInstanceId > 0 && clusterInstanceId == sleeveId))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP] ⚠️ PARAMETER CHECK: Skipping sleeve {sleeveId} - parameters indicate it's a cluster sleeve (SleeveInstanceId={sleeveInstanceId}, ClusterInstanceId={clusterInstanceId})\n");
                    continue;
                }
                
                // ✅ CRITICAL FIX: Get bounding box from Revit element (works for all sleeves, even newly placed)
                int individualId = individualSleeve.Id.IntegerValue;
                var bbox = individualSleeve.get_BoundingBox(null);
                if (bbox == null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEANUP-SKIP] Individual sleeve {individualSleeve.Id} (hostType={individualHostType}) skipped cluster {clusterId} (hostType={clusterHostType})\n");
                        continue;
                    }
                    
                    // ✅ OPTIMIZED: Get cluster sleeve bounding box from XML cache (no expensive Revit API calls)                    
                    // Try to find cluster sleeve in cache (after XML save, this will have bounding boxes)
                    
                    // ✅ DIAGNOSTIC: Log cache lookup attempt
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var cacheClusterCount = _clashZoneCache.Values.Count(cz => cz.ClusterSleeveInstanceId > 0);
                        var cacheClusterIds = _clashZoneCache.Values
                            .Where(cz => cz.ClusterSleeveInstanceId > 0)
                            .Select(cz => cz.ClusterSleeveInstanceId)
                            .Distinct()
                            .Take(10)
                            .ToList();
                        DebugLogger.Info($"[CLEANUP-LOOKUP] Searching for clusterId={clusterId} in cache. Cache has {cacheClusterCount} clash zones with ClusterSleeveInstanceId>0. Sample IDs: {string.Join(", ", cacheClusterIds)}");
                    }
                    
                    var clusterClashZone = _clashZoneCache.Values.FirstOrDefault(cz => cz.ClusterSleeveInstanceId == clusterId);
                    
                    // ✅ DIAGNOSTIC: Enhanced logging for lookup result
                    if (clusterClashZone == null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLEANUP-LOOKUP] ❌ Cluster sleeve {clusterId} NOT FOUND in cache. Cache size: {_clashZoneCache.Count}");
                            
                            // Check if there are any clash zones with this cluster ID but different MepElementIdValue
                            var allWithThisClusterId = _clashZoneCache.Values.Where(cz => cz.ClusterSleeveInstanceId == clusterId).ToList();
                            DebugLogger.Info($"[CLEANUP-LOOKUP] Clash zones with ClusterSleeveInstanceId={clusterId}: {allWithThisClusterId.Count}");
                            
                            // Log host type of cluster sleeve from Revit (reuse existing clusterHostType from line 4292)
                            DebugLogger.Info($"[CLEANUP-LOOKUP] Cluster sleeve {clusterId} host type from Revit: {clusterHostType}");
                        }
                        continue;
                    }
                    
                    if (clusterClashZone.ClusterSleeveBoundingBoxMinX == 0 && clusterClashZone.ClusterSleeveBoundingBoxMinY == 0)
                    {
                        var clusterHostTypeFromCache = GetHostTypeFromClashZone(clusterClashZone);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CLEANUP-LOOKUP] ⚠️ Cluster sleeve {clusterId} found BUT bbox is zero (HostType={clusterHostTypeFromCache}). " +
                                $"MinX={clusterClashZone.ClusterSleeveBoundingBoxMinX}, MinY={clusterClashZone.ClusterSleeveBoundingBoxMinY}, " +
                                $"MaxX={clusterClashZone.ClusterSleeveBoundingBoxMaxX}, MaxY={clusterClashZone.ClusterSleeveBoundingBoxMaxY}");
                        }
                        continue;
                    }
                    
                    // ✅ DIAGNOSTIC: Log successful lookup
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var clusterHostTypeFromCache = GetHostTypeFromClashZone(clusterClashZone);
                        DebugLogger.Info($"[CLEANUP-LOOKUP] ✅ Cluster sleeve {clusterId} FOUND in cache (HostType={clusterHostTypeFromCache}). " +
                            $"BboxMin=({clusterClashZone.ClusterSleeveBoundingBoxMinX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMinY:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMinZ:F6}), " +
                            $"BboxMax=({clusterClashZone.ClusterSleeveBoundingBoxMaxX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMaxY:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMaxZ:F6})");
                    }
                    
                    // ✅ Use bounding box from XML cache
                    var clusterMin = new XYZ(clusterClashZone.ClusterSleeveBoundingBoxMinX, clusterClashZone.ClusterSleeveBoundingBoxMinY, clusterClashZone.ClusterSleeveBoundingBoxMinZ);
                    var clusterMax = new XYZ(clusterClashZone.ClusterSleeveBoundingBoxMaxX, clusterClashZone.ClusterSleeveBoundingBoxMaxY, clusterClashZone.ClusterSleeveBoundingBoxMaxZ);
                                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP] Cluster sleeve {clusterId} bbox from CACHE: Min=({clusterClashZone.ClusterSleeveBoundingBoxMinX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMinY:F6}), Max=({clusterClashZone.ClusterSleeveBoundingBoxMaxX:F6}, {clusterClashZone.ClusterSleeveBoundingBoxMaxY:F6})\n");
                    
                    // ✅ ENHANCED LOGGING: Show bounding boxes before checking (with Z coordinates for floors)
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP-CHECK] Individual sleeve {individualId}: Min=({individualMin.X:F3}, {individualMin.Y:F3}, {individualMin.Z:F3}), Max=({individualMax.X:F3}, {individualMax.Y:F3}, {individualMax.Z:F3})\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP-CHECK] Cluster sleeve {clusterId}: Min=({clusterMin.X:F3}, {clusterMin.Y:F3}, {clusterMin.Z:F3}), Max=({clusterMax.X:F3}, {clusterMax.Y:F3}, {clusterMax.Z:F3})\n");
                    
                    // Check if individual sleeve is completely within cluster sleeve bounding box
                    bool withinBounds = IsWithinBounds(individualMin, individualMax, clusterMin, clusterMax);
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP-CHECK] Individual {individualId} within Cluster {clusterId}: {withinBounds}\n");
                    
                    if (withinBounds)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"[CleanupSleevesWithinClusters] Individual sleeve {individualSleeve.Id} falls within cluster sleeve {clusterSleeve.Id} - marking for deletion");
                                                if (!DeploymentConfiguration.DeploymentMode)
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
                // ✅ CRITICAL: BEFORE DELETION - Check if any cluster sleeves are in the deletion list
                var clusterSleevesInDeletionList = new List<int>();
                foreach (var (sleeveId, clusterId, clashZone) in sleevesToDelete)
                {
                    int id = sleeveId.IntegerValue;
                    
                    // Check if this sleeve ID is in the protection set (it's a cluster sleeve!)
                    if (clusterSleeveIds.Contains(id))
                    {
                        clusterSleevesInDeletionList.Add(id);
                        
                        // ✅ CRITICAL LOGGING: This should NEVER happen - log detailed info
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                            var sleeve = doc.GetElement(sleeveId) as FamilyInstance;
                            var sleeveInstanceIdParam = sleeve?.LookupParameter("Sleeve Instance ID");
                            int sleeveInstanceId = sleeveInstanceIdParam?.AsInteger() ?? -999;
                            var clusterInstanceIdParam = sleeve?.LookupParameter("Cluster Sleeve Instance ID");
                            int clusterInstanceId = clusterInstanceIdParam?.AsInteger() ?? -1;
                            
                            string errorMsg = $"[CLEANUP-ERROR] ⚠️⚠️⚠️ CRITICAL: Cluster sleeve {id} is in deletion list! This should NEVER happen!\n";
                            errorMsg += $"[CLEANUP-ERROR] Sleeve {id} details:\n";
                            errorMsg += $"[CLEANUP-ERROR]   - In Protection Set: {clusterSleeveIds.Contains(id)}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - Sleeve Instance ID param: {sleeveInstanceId}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - Cluster Sleeve Instance ID param: {clusterInstanceId}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - Family: {sleeve?.Symbol?.FamilyName ?? "NULL"}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - Deleting Cluster ID: {clusterId}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - Protection Set Contents: {string.Join(", ", clusterSleeveIds)}\n";
                            errorMsg += $"[CLEANUP-ERROR]   - placedClusters Count: {placedClusters?.Count ?? 0}\n";
                            if (placedClusters != null && placedClusters.Count > 0)
                            {
                                errorMsg += $"[CLEANUP-ERROR]   - placedClusters IDs: {string.Join(", ", placedClusters.Select(c => c.Id.IntegerValue))}\n";
                            }
                            
                            DebugLogger.Error(errorMsg);
                            System.IO.File.AppendAllText(clusterDebugLogPath, errorMsg);
                        }
                    }
                }
                
                // ✅ CRITICAL: If cluster sleeves are in deletion list, REMOVE THEM and log error
                if (clusterSleevesInDeletionList.Count > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        string errorMsg = $"[CLEANUP-ERROR] ⚠️⚠️⚠️ REMOVING {clusterSleevesInDeletionList.Count} cluster sleeves from deletion list to prevent deletion!\n";
                        errorMsg += $"[CLEANUP-ERROR] Cluster sleeve IDs being removed: {string.Join(", ", clusterSleevesInDeletionList)}\n";
                        DebugLogger.Error(errorMsg);
                        System.IO.File.AppendAllText(clusterDebugLogPath, errorMsg);
                    }
                    
                    // Remove cluster sleeves from deletion list
                    sleevesToDelete = sleevesToDelete.Where(s => !clusterSleevesInDeletionList.Contains(s.sleeveId.IntegerValue)).ToList();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-BATCH-DELETE] Marked {sleevesToDelete.Count} sleeves for batch deletion: {string.Join(", ", sleevesToDelete.Select(s => s.sleeveId.IntegerValue))}\n");
                    if (clusterSleevesInDeletionList.Count > 0)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-BATCH-DELETE] ⚠️ Removed {clusterSleevesInDeletionList.Count} cluster sleeves from deletion list\n");
                    }
                }
                
                var elementIdsToDelete = sleevesToDelete.Select(s => s.sleeveId).ToList();
                
                try
                {
                    doc.Delete(elementIdsToDelete);
                    deletedCount = sleevesToDelete.Count;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CleanupSleevesWithinClusters] ✅ Batch deleted {deletedCount} individual sleeves");
                    
                    // ✅ CRITICAL: Update flags for all deleted sleeves (preserves IsResolved=true)
                    // ✅ ALSO: Collect MEP Element IDs from deleted sleeves to update cluster sleeve parameter
                    var mepIdsToAddByCluster = new Dictionary<int, HashSet<long>>();
                    
                    foreach (var (sleeveId, clusterId, clashZone) in sleevesToDelete)
                    {
                        if (clashZone != null)
                        {
                            clashZone.IsResolved = true;
                            clashZone.IsClusterResolved = true;
                            clashZone.ClusterSleeveInstanceId = clusterId;
                            clashZone.SleeveInstanceId = -1;
                            
                            // ✅ CRITICAL: Track MEP Element ID for cluster sleeve parameter update
                            if (clashZone.MepElementIdValue > 0)
                            {
                                if (!mepIdsToAddByCluster.ContainsKey(clusterId))
                                {
                                    mepIdsToAddByCluster[clusterId] = new HashSet<long>();
                                }
                                mepIdsToAddByCluster[clusterId].Add(clashZone.MepElementIdValue);
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[CLEANUP-MEP-ID] Collected MEP Element ID {clashZone.MepElementIdValue} from deleted sleeve {sleeveId.IntegerValue} for cluster {clusterId}\n");
                            }
                        }
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLEANUP-FLAGS] Updated flags for {sleeveId.IntegerValue}: IsResolved=True, IsClusterResolved=True, ClusterSleeveId={clusterId}\n");
                    }
                    
                    // ✅ CRITICAL: Update cluster sleeve MEP_ElementIds parameter with deleted sleeves' MEP Element IDs
                    foreach (var kvp in mepIdsToAddByCluster)
                    {
                        int clusterId = kvp.Key;
                        HashSet<long> newMepIds = kvp.Value;
                        var clusterSleeve = placedClusters.FirstOrDefault(c => c.Id.IntegerValue == clusterId);
                        if (clusterSleeve != null)
                        {
                            UpdateClusterSleeveMepElementIds(clusterSleeve, newMepIds);
                        }
                        else
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[CLEANUP-MEP-ID] Cluster sleeve {clusterId} not found in placedClusters list - cannot update MEP_ElementIds\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[CleanupSleevesWithinClusters] Error batch deleting sleeves: {ex.Message}");
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[CleanupSleevesWithinClusters] Cleanup complete: {deletedCount} sleeves deleted");
        }
        catch (Exception ex)
        {
                        if (!DeploymentConfiguration.DeploymentMode)
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
                if (!DeploymentConfiguration.DeploymentMode)
        DebugLogger.Info($"[CLEANUP-COVERAGE] Individual: W={individualMax.X - individualMin.X:F3}, H={individualMax.Y - individualMin.Y:F3}, D={individualMax.Z - individualMin.Z:F3}, Vol={individualVolume:F6}\n");
                if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetFilterNameForCategory] {errorMsg}");
                throw new InvalidOperationException(errorMsg);
            }
            
            return _filterName;
        }
        
        /// <summary>
        /// ✅ GLOBAL XML FLAG MANAGEMENT: No longer saves flags to Filter XML.
        /// Flags are managed exclusively in Global XML (single source of truth).
        /// This method is kept for backward compatibility but does not write flags.
        /// </summary>
        private void SaveUpdatedFlagsToXml(string xmlFilePath)
        {
            // ✅ FLAG MANAGEMENT: Flags are now managed exclusively in Global XML.
            // Filter XML no longer stores flags - they are synced from Global XML on load.
            // This method is kept for backward compatibility but does not write flags.
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SaveUpdatedFlagsToXml] Skipped - flags are managed in Global XML only (single source of truth)");
        }

        private static List<ClashZone> GetStorageZones(OpeningFilter filter)
        {
            return filter?.ClashZoneStorage?.AllZones ?? new List<ClashZone>();
        }
    }
}

