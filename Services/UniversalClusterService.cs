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
                            
                            PlaceClusterSleeve(doc, cluster, groupKey, targetCategory, out int placed1, out int deleted1, xmlFilePath);
                            
                            try 
                            { 
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(placementDebugPath3, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING] PlaceClusterSleeve returned: placed={placed1}, deleted={deleted1}\n");
                                }
                            } 
                            catch { }
                            placedCount += placed1;
                            deletedCount += deleted1;
                            
                            // ✅ PERFORMANCE: Track cluster sleeve for cleanup (if placed successfully)
                            // Note: PlaceClusterSleeve creates the instance internally, so we need to find it
                            // We'll track it by finding the newest cluster sleeve with matching properties
                            if (placed1 > 0)
                            {
                                // ✅ PERFORMANCE: Skip FindNewestClusterSleeve (was slow) - cluster sleeves will be tracked via cleanup pass
                                // Cleanup will find all cluster sleeves by family name, so we don't need to track individually
                            }
                            
                            // ✅ PERFORMANCE: Removed duplicate flag updates - PlaceClusterSleeve already handles flag updates internally
                            // Flags are updated inside PlaceClusterSleeve via MarkClashZonesAsClusterResolvedWithSleeveId and FlagManager
                            // No need to call UpdateClashZoneFlagsForCluster again (was causing 90+ second delays)
                            
                            // Marking happens inside PlaceClusterSleeve with actual cluster sleeve ID
                        }
                        catch (Exception ex)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalClusterService] Error placing cluster: {ex.Message}");
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
                    DebugLogger.Log($"[UniversalClusterService] Returning {placedClusters.Count} placed cluster sleeves for coordinate update");
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
            if (!isPath1Replay && comboId.HasValue && filterId.HasValue && placedCount > 0)
            {
                try
                {
                    SaveClusterDataToDatabase(doc, placedClusters, comboId.Value, filterId.Value, targetCategory, xmlFilePath);
                }
                catch (Exception saveEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[CLUSTERING] Error saving cluster data to database: {saveEx.Message}");
                    }
                    // Non-critical error, continue
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

                        // Apply rotation if needed
                        if (clusterData.IsRotated && Math.Abs(clusterData.RotationAngleDeg) > 1e-6)
                        {
                            var rotationAngle = clusterData.RotationAngleDeg * Math.PI / 180.0;
                            var rotationAxis = XYZ.BasisZ;
                            var rotation = Transform.CreateRotationAtPoint(rotationAxis, rotationAngle, placementPoint);
                            ElementTransformUtils.MoveElement(doc, clusterSleeve.Id, rotation.OfPoint(placementPoint) - placementPoint);
                            ElementTransformUtils.RotateElement(doc, clusterSleeve.Id, Line.CreateBound(placementPoint, placementPoint + XYZ.BasisZ), rotationAngle);
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
            string xmlFilePath)
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

                        // Get ClashZoneIds from MEP_ElementIds parameter or from XML
                        List<Guid> clashZoneIds = new List<Guid>();
                        var mepElementIdsParam = clusterSleeve.LookupParameter("MEP_ElementIds");
                        if (mepElementIdsParam != null && mepElementIdsParam.HasValue)
                        {
                            // Parse MEP Element IDs and find corresponding clash zones
                            // This is simplified - in production, you might want to store ClashZoneIds directly
                            // For now, we'll try to get from XML cache
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
        private void MarkClashZonesAsClusterResolvedWithSleeveId(List<dynamic> cluster, ElementId clusterSleeveId, string xmlFilePath = null, BoundingBoxXYZ clusterBbox = null)
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
                    
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[MarkClusterResolved] Updating cluster flags in {xmlFiles.Length} XML file(s): {(string.IsNullOrEmpty(xmlFilePath) ? "ALL" : Path.GetFileName(xmlFilePath))} for cluster sleeve {clusterSleeveId.IntegerValue}\n");
                }
                
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

                        if (filter?.ClashZoneStorage?.AllZones != null)
                        {
                            bool updated = false;
                            
                            foreach (var sleeve in cluster)
                            {
                                // ✅ FIXED: Use sleeve instance ID from XML data
                                int sleeveInstanceId = sleeve.SleeveInstanceId;
                                
                                // Find and mark clash zone as cluster-resolved using sleeve instance ID
                                var clashZone = GetStorageZones(filter).FirstOrDefault(cz => cz.SleeveInstanceId == sleeveInstanceId);
                                if (clashZone != null)
                                {
                                    clashZone.IsClusterResolved = true;
                                    clashZone.ClusterSleeveId = clusterSleeveId;
                                    clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue; // ✅ FIX: Store integer for XML serialization
                                    clashZone.LastUpdated = DateTime.Now;
                                    
                                    // ✅ CONSOLIDATION: Set cluster bounding box coordinates from Revit immediately after placement
                                    // This eliminates the need for SleeveCoordinateService to update cluster bounding boxes later
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
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[MarkClusterResolved] ⚠️ Cluster bounding box not available for ClashZone {clashZone.Id} (cluster sleeve may not be fully initialized)");
                                    }
                                    
                                    // ✅ CORRECT: Individual sleeve was placed then deleted during clustering
                                    // Set clustering history flag to distinguish from non-proximity sleeves
                                    clashZone.MarkedForClusteringSleeveProcess = true; // ✅ FIX: Mark as part of clustering history
                                    clashZone.IsResolved = true; // Individual sleeve was placed (then deleted)
                                    
                                    // ✅ STORAGE: Save original SleeveInstanceId BEFORE clearing it
                                    clashZone.AfterClusterSleevePlacedSleeveInstanceId = clashZone.SleeveInstanceId;
                                    
                                    clashZone.SleeveInstanceId = -1; // Individual sleeve was deleted
                                    clashZone.SleeveFamilyName = string.Empty; // Individual sleeve family cleared
                                    
                                    // ✅ FLAG MANAGER: Persist cluster resolution to Global XML so future runs respect the flag
                                    try
                                    {
                                        var categoryName = clashZone.MepElementCategory ?? string.Empty;
                                        var baseFilterName = FilterNameHelper.NormalizeBaseName(_filterName, filter?.Name, categoryName);
                                        
                                        // ✅ CRITICAL: Update flags in memory first
                                        clashZone.IsClusterResolved = true;
                                        clashZone.ClusterSleeveInstanceId = clusterSleeveId.IntegerValue;
                                        clashZone.IsResolved = true;
                                        clashZone.SleeveInstanceId = -1;
                                        
                                        // ✅ CRITICAL: Update Global XML immediately (don't wait for batch)
                                        _flagManager?.UpdateFlagsForPlacement(clashZone, clusterSleeveId.IntegerValue, isCluster: true, categoryName, baseFilterName);
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[GLOBAL-XML] ✅ Updated Global XML for ClashZone {clashZone.Id} with cluster sleeve {clusterSleeveId.IntegerValue}");
                                    }
                                    catch (Exception flagEx)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[MarkClusterResolved] ⚠️ Failed to update Global XML flags for clash zone {clashZone.Id}: {flagEx.Message}");
                                    }
                                    
                                    // ⚠️ CRITICAL: Log flag state AFTER cluster sleeve placement
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] ClashZone {clashZone.Id}: Cluster sleeve {clusterSleeveId.IntegerValue} placed\n");
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [CLUSTER-PLACED] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                                    
                                    updated = true;
                                    markedCount++;
                                    
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[DEBUG] Marked ClashZone {clashZone.Id} as cluster-resolved with cluster sleeve {clusterSleeveId.IntegerValue} (cleared individual flags)\n");
                                }
                                else
                                {
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[DEBUG] ✗ Clash zone not found for SleeveInstanceId={sleeveInstanceId}\n");
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
                    
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[ResetFlags] Processing {xmlFiles.Length} XML file(s): {(string.IsNullOrEmpty(xmlFilePath) ? "ALL" : Path.GetFileName(xmlFilePath))}\n");
                }
                
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
                        
                        if (filter?.ClashZoneStorage?.AllZones != null)
                        {
                            bool modified = false;
                            
                            var storageZonesForResetFlags = GetStorageZones(filter);
                            if (storageZonesForResetFlags.Count > 0)
                            {
                                foreach (var clashZone in storageZonesForResetFlags)
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
                                                                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                                                            if (!DeploymentConfiguration.DeploymentMode)
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
                                                                                            if (!DeploymentConfiguration.DeploymentMode)
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
                                                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[DEBUG] Reset individual sleeve flag for ClashZone {clashZone.Id} - sleeve {clashZone.SleeveInstanceId} was deleted\n");
                                            }
                                        }
                                    }
                                }
                                
                                // ✅ FIX: Only save if modifications were made, and reader is already closed
                                if (modified)
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
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[DEBUG] ✓ Saved reset flags to {Path.GetFileName(xmlFile)}\n");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[UniversalClusterService] Error processing XML file {xmlFile}: {ex.Message}");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DEBUG] ✗ Error resetting flags in {Path.GetFileName(xmlFile)}: {ex.Message}\n");
                    }
                }

                if (resetCount > 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DEBUG] ✓ Reset flags for {resetCount} deleted sleeves (cluster + individual)\n");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DEBUG] ✓ No deleted sleeves found - all flags preserved\n");
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
                                
                            if (BoundingBoxesOverlapFromXml(clusterSleeve, otherSleeve, toleranceDist))
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
                
                double minDistance;
                
                if (hostType == "Floor")
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
                    System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [CLUSTERING-DISTANCE] Sleeve {sleeve1.SleeveInstanceId} vs {sleeve2.SleeveInstanceId}: " +
                        $"Host={hostType}, Ori={orientation}, Distance={UnitUtils.ConvertFromInternalUnits(minDistance, UnitTypeId.Millimeters):F1}mm, " +
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
            
            // ✅ DEBUG: Log coordinates and orientation
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect1: Min=({current.SleeveBoundingBoxMinX:F3}, {current.SleeveBoundingBoxMinY:F3}), Max=({current.SleeveBoundingBoxMaxX:F3}, {current.SleeveBoundingBoxMaxY:F3})\n");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect2: Min=({other.SleeveBoundingBoxMinX:F3}, {other.SleeveBoundingBoxMinY:F3}), Max=({other.SleeveBoundingBoxMaxX:F3}, {other.SleeveBoundingBoxMaxY:F3})\n");
                        if (!DeploymentConfiguration.DeploymentMode)
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DISTANCE-DEBUG] Using Y,Z coordinates for Y-oriented walls\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DISTANCE-DEBUG] Rect1 YZ: Min=({current.SleeveBoundingBoxMinY:F3}, {current.SleeveBoundingBoxMinZ:F3}), Max=({current.SleeveBoundingBoxMaxY:F3}, {current.SleeveBoundingBoxMaxZ:F3})\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                                if (!DeploymentConfiguration.DeploymentMode)
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
                // ✅ PERFORMANCE: First try cache lookup (fast O(N) search)
                foreach (var clashZone in _clashZoneCache.Values)
                {
                    if (clashZone.SleeveInstanceId == sleeveInstanceId)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"[UniversalClusterService] Found clash zone for Sleeve Instance ID {sleeveInstanceId} in cache");
                        return clashZone;
                        }
                    }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] Sleeve Instance ID {sleeveInstanceId} not found in cache (cache size: {_clashZoneCache?.Count ?? 0}) - trying XML fallback");
                
                // ✅ CRITICAL FIX: Fallback to XML if cache lookup fails (similar to GetClashZoneByMepElementId)
                // This ensures GUID is always set even if cache is incomplete
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalClusterService] Filters directory does not exist for XML fallback");
                    return null;
                }
                
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
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalClusterService] ❌ Sleeve Instance ID {sleeveInstanceId} not found in cache or XML files");
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
                DebugLogger.Log($"Unknown host type for cluster group, skipping. HostType={groupKey.hostType}");
                return;
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
                    DebugLogger.Error($"[ClusterService] Failed to load universal family '{familyName}'");
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
                    DebugLogger.Error($"[ClusterService] Still no family found for '{familyName}' after loading attempt");
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
                DebugLogger.Error($"[ClusterService] No actual sleeves found for cluster group, skipping placement.");
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
            
            double rotationAngle = DetermineDominantRotationAngle(cluster, xmlFilePath);
            
            // ✅ DEBUG: Log the determined rotation angle
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CLUSTER-ANGLE-DEBUG] DetermineDominantRotationAngle returned: {rotationAngle * 180 / Math.PI:F1}° (radians: {rotationAngle:F6})");
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
            // ✅ NEW: Pass rotation angle to calculate bounding box in rotated coordinate system
            var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(actualSleeves, rotationAngle);
            
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
                    DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                return;
            }

            // Place cluster sleeve
            // ✅ SIMPLIFIED: Use same universal family symbol as placer service
            FamilyInstance inst = doc.Create.NewFamilyInstance(mid, familySymbol, refLevel!, StructuralType.NonStructural);

            // ✅ FIXED: Apply rotation if needed (for cases where rotation wasn't fully accounted for in bounding box)
            // ⚠️ CRITICAL: Only apply rotation if angle is NOT axis-aligned (0°, 90°, 180°, 270°)
            // The DetermineDominantRotationAngle method should return 0.0 for axis-aligned angles,
            // but we add an extra safety check here to prevent accidental rotation
            if (Math.Abs(rotationAngle) > 1e-6)
            {
                // Double-check: if angle is close to 0°, 90°, 180°, or 270°, don't rotate
                double angleDegrees = rotationAngle * 180 / Math.PI;
                // Normalize to 0-360 range
                while (angleDegrees < 0) angleDegrees += 360;
                while (angleDegrees >= 360) angleDegrees -= 360;
                
                double distTo0 = Math.Min(angleDegrees, 360 - angleDegrees);
                double distTo90 = Math.Abs(angleDegrees - 90);
                double distTo180 = Math.Abs(angleDegrees - 180);
                double distTo270 = Math.Abs(angleDegrees - 270);
                
                double thresholdDegrees = 2.0; // 2 degree tolerance
                bool isAxisAligned = distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                                    distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
                
                if (isAxisAligned)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ROTATION] ⚠️ Skipping rotation for cluster sleeve {inst.Id}: Angle {angleDegrees:F1}° is axis-aligned (distTo0={distTo0:F1}°, distTo90={distTo90:F1}°, distTo180={distTo180:F1}°, distTo270={distTo270:F1}°)");
                    }
                }
                else
                {
                    XYZ axisOrigin = mid;
                    XYZ axisDirection = XYZ.BasisZ;
                    Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                    ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ROTATION] ✅ Applied rotation {angleDegrees:F1}° to cluster sleeve {inst.Id}");
                    }
                }
            }

            // Set size parameters (swap dimensions if rotated for orientation alignment)
            // ✅ FIX: Apply swap for both walls and framing when X-oriented
            bool shouldSwapDimensions = ((groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing") && rotationAngle != 0.0);
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);
            
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
            MarkClashZonesAsClusterResolvedWithSleeveId(cluster, inst.Id, xmlFilePath, clusterBbox);

            // ✅ PERFORMANCE: Removed XML file reading from inside PlaceClusterSleeve (was causing 90+ second delays per cluster)
            // Flag updates are handled by MarkClashZonesAsClusterResolvedWithSleeveId which updates XML files
            // FlagManager updates will be handled in batch after all clusters are placed (not per-cluster)

                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[ClusterService] About to delete {cluster.Count} individual sleeves for cluster sleeve {inst.Id.IntegerValue}");
                        if (!DeploymentConfiguration.DeploymentMode)
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
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] Checking sleeve {sleeveElementId.IntegerValue}: found={sleeveElement != null}, isFamilyInstance={sleeveElement is FamilyInstance}\n");
                    
                    if (sleeveElement != null && sleeveElement is FamilyInstance sleeveInstance)
                    {
                        sleevesToDelete.Add(sleeveElementId);
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"[ClusterService] Queued for deletion: individual sleeve {sleeveElementId.IntegerValue}");
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[ClusterService] Could not find sleeve {s.SleeveInstanceId} for deletion");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[DELETE] ✗ Sleeve {sleeveElementId.IntegerValue} not found or not a FamilyInstance\n");
                    }
                }
                catch (Exception ex)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ClusterService] Error preparing sleeve {s.SleeveInstanceId} for deletion: {ex.Message}");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[DELETE] ✗ Exception preparing sleeve: {ex.Message}\n");
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[DELETE] Total sleeves queued for deletion: {sleevesToDelete.Count}\n");
            
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
                    int edgeCaseDeleted = DeleteEdgeCaseSleevesInClusterZone(doc, clusterBbox, targetCategory, cluster, inst.Id, xmlFilePath);
                    deleted += edgeCaseDeleted;
                    
                    if (edgeCaseDeleted > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[DELETE] ✅ EDGE CASE: Deleted {edgeCaseDeleted} individual sleeves that fell in cluster zone (cluster sleeve {inst.Id.IntegerValue})");
                    }
                }
                catch (Exception edgeEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[ClusterService] Error deleting edge case sleeves: {edgeEx.Message}");
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
                foreach (var sleeve in allSleeves)
                {
                    // Skip if part of cluster formation
                    if (clusterFormationSleeveIds.Contains(sleeve.Id.IntegerValue))
                        continue;
                    
                    // Skip if it's a cluster sleeve (check by family name or parameters)
                    if (sleeve.Symbol?.FamilyName?.Contains("Cluster", StringComparison.OrdinalIgnoreCase) == true)
                        continue;
                    
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
                var edgeCaseUpdates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId)>();
                
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
                                        clashZone.ClusterSleeveInstanceId // OLD value for matching
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
                                                guidZone.ClusterSleeveInstanceId // OLD value for matching
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
                                OldClusterInstanceId: u.OldClusterInstanceId
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
                var clashZones = LoadClashZonesFromRegularXml(null, category, doc);
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
                var filterManagementService = new FilterManagementService(_doc, null, null);
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
            string filterName = GetFilterNameForCategory(category);
            var filterNameParam = clusterSleeve.LookupParameter("Filter Name");
            if (filterNameParam != null && !filterNameParam.IsReadOnly)
            {
                filterNameParam.Set(filterName);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[SetClusterSleeveMetadata] Set Filter Name = '{filterName}' for cluster sleeve {clusterSleeve.Id}");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[SetClusterSleeveMetadata] Filter Name parameter not found or read-only on cluster sleeve {clusterSleeve.Id}");
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
            
            // Log cluster sleeve IDs for debugging
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-START] Cluster sleeve IDs: {string.Join(", ", placedClusters.Select(c => c.Id.IntegerValue))}\n");
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLEANUP-DEBUG-ALL] Sleeve {s.Id}: Family={s.Symbol?.FamilyName}, ClusterParam={clusterValue}, SleeveInstanceId={sleeveInstanceId}, IsCluster={isClusterSleeve}, IsIndividual={isIndividual}\n");
                
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
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLEANUP-LOOP] Processing individual sleeve {individualSleeve.Id}\n");
                
                // ✅ FIX: Skip if this sleeve is actually a cluster sleeve in our placedClusters list
                if (placedClusters.Any(c => c.Id.IntegerValue == individualSleeve.Id.IntegerValue))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLEANUP] Skipping sleeve {individualSleeve.Id} - it's a cluster sleeve\n");
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[CLEANUP-BATCH-DELETE] Marked {sleevesToDelete.Count} sleeves for batch deletion: {string.Join(", ", sleevesToDelete.Select(s => s.sleeveId.IntegerValue))}\n");
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

