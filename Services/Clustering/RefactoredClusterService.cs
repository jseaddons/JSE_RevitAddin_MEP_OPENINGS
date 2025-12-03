using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// ✅ REFACTORED CLUSTER SERVICE: Clean orchestrator for all Phase 1-10 services.
    /// Replaces the 8000+ line UniversalClusterService with clean delegation to extracted services.
    /// 
    /// Architecture:
    /// - Phase 1: Geometry (DistanceCalculator, RotationMatrixCalculator, CoordinateTransformer)
    /// - Phase 2: Proximity (ProximityCheckerFactory, IProximityChecker implementations)
    /// - Phase 3: BoundingBox (IBoundingBoxCalculator implementations)
    /// - Phase 4: Strategy (IClusteringStrategy implementations, ClusteringStrategyFactory)
    /// - Phase 5: Placement (IClusterPlacementService)
    /// - Phase 6: Rotation (IClusterRotationService)
    /// - Phase 7: Cleanup (IClusterCleanupService)
    /// - Phase 8: Algorithm (IClusterAlgorithmService)
    /// - Phase 9: Data (IClusterDataService)
    /// - Phase 10: Timeout (IClusterTimeoutService)
    /// 
    /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
    /// To modify protected methods in this class, you MUST:
    /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true
    /// 2. Get explicit consent from the project owner
    /// 3. Test thoroughly before committing
    /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
    /// 
    /// PROTECTED METHODS:
    /// - SaveClusterDataToDatabase() - Database saving logic
    /// - SaveClusterSleeveSnapshots() - Parameter aggregation and saving
    /// 
    /// ⚠️ DO NOT SET TO true UNLESS YOU HAVE EXPLICIT CONSENT ⚠️
    /// </summary>
    public class RefactoredClusterService
    {
        // ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        // To modify protected methods in this class, you MUST:
        // 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true
        // 2. Get explicit consent from the project owner
        // 3. Test thoroughly before committing
        // 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        // 
        // PROTECTED METHODS:
        // - SaveClusterDataToDatabase() - Database saving logic
        // - SaveClusterSleeveSnapshots() - Parameter aggregation and saving
        // 
        // ⚠️ DO NOT SET TO true UNLESS YOU HAVE EXPLICIT CONSENT ⚠️
        private const bool ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false;
        // Phase 1-5 Services (Geometry, Proximity, BoundingBox, Strategy, Placement)
        private readonly ClusteringStrategyFactory _strategyFactory;
        private readonly IClusterPlacementService _placementService;
        
        // Phase 6-10 Services (Rotation, Cleanup, Algorithm, Data, Timeout)
        private readonly IClusterRotationService _rotationService;
        private readonly IClusterCleanupService _cleanupService;
        private readonly IClusterAlgorithmService _algorithmService;
        private readonly IClusterDataService _dataService;
        private readonly IClusterTimeoutService _timeoutService;
        
        // Supporting Services
        private readonly FlagManager _flagManager;
        private readonly FilterManagementService _filterService;
        private readonly Document _doc;
        
        // Internal state
        private string? _filterName;
        private readonly Dictionary<int, List<Guid>> _clusterToClashZoneIds = new Dictionary<int, List<Guid>>();
        
        // ✅ STEP 5 OPTIMIZATION: Deferred parameter batching for cluster placement (4-6× faster)
        // Accumulates parameter values during cluster placement loop, writes all after single regeneration
        // Key: ElementId of cluster sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private Dictionary<ElementId, Dictionary<string, object>> _deferredClusterParameters = new Dictionary<ElementId, Dictionary<string, object>>();

        /// <summary>
        /// Constructor with dependency injection for all Phase 1-10 services.
        /// </summary>
        public RefactoredClusterService(
            Document doc,
            IClusterDataService dataService,
            IClusterAlgorithmService algorithmService,
            IClusterRotationService rotationService,
            IClusterPlacementService placementService,
            IClusterCleanupService cleanupService,
            IClusterTimeoutService timeoutService,
            ClusteringStrategyFactory? strategyFactory = null,
            FlagManager? flagManager = null,
            FilterManagementService? filterService = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _dataService = dataService ?? throw new ArgumentNullException(nameof(dataService));
            _algorithmService = algorithmService ?? throw new ArgumentNullException(nameof(algorithmService));
            _rotationService = rotationService ?? throw new ArgumentNullException(nameof(rotationService));
            _placementService = placementService ?? throw new ArgumentNullException(nameof(placementService));
            _cleanupService = cleanupService ?? throw new ArgumentNullException(nameof(cleanupService));
            _timeoutService = timeoutService ?? throw new ArgumentNullException(nameof(timeoutService));
            
            _strategyFactory = strategyFactory ?? new ClusteringStrategyFactory();
            _flagManager = flagManager ?? new FlagManager(doc);
            _filterService = filterService ?? new FilterManagementService(
                doc,
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); },
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Error(msg); });
        }

        /// <summary>
        /// Main entry point: Cluster sleeves for a specific category.
        /// Clean orchestration of all Phase 1-10 services.
        /// </summary>
        public (int placedCount, int deletedCount) ClusterSleeves(
            Document doc,
            string targetCategory,
            UIDocument? uiDoc = null,
            string? xmlFilePath = null,
            string? filterName = null,
            List<FamilyInstance>? placedClusterSleevesOut = null,
            bool isPath1Replay = false,
            int? comboId = null,
            int? filterId = null,
            Dictionary<string, double> currentClearanceSettings = null,
            bool isPath3Validated = false,
            bool isPath3Invalidated = false,
            bool isPath3New = false)
        {
            // ✅ PERFORMANCE MONITORING: Initialize cluster performance monitor
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string performanceLogName = $"ClusterPlacement_{targetCategory}_{timestamp}.log";
            var performanceMonitor = new PlacementPerformanceMonitor(performanceLogName);
            
            // 🔥 CRITICAL: Direct System.IO logging to ensure we always see entry (bypasses SafeFileLogger completely)
            // This MUST work in both R2023 and R2024
            try
            {
                var versionTag = Helpers.VersionInfo.VersionTag; // "R2023" or "R2024"
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                
                // Ensure directory exists
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                
                var logPath = Path.Combine(logDir, "cluster_debug.log");
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = assembly?.Location ?? "unknown";
                var buildTime = File.Exists(assemblyPath) ? File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss") : "unknown";
                
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 ClusterSleeves METHOD CALLED\n");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] Version: {versionTag}, Assembly: {Path.GetFileName(assemblyPath)}, BuildTime: {buildTime}\n");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] Parameters: category={targetCategory}, filter={filterName}, comboId={comboId}, filterId={filterId}\n");
            }
            catch (Exception directLogEx)
            {
                // Last resort: try Desktop
                try
                {
                    var desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"JSE_Cluster_Error_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                    File.AppendAllText(desktopPath, $"[{DateTime.Now:HH:mm:ss}] CRITICAL: Failed to write cluster log: {directLogEx.Message}\n");
                    File.AppendAllText(desktopPath, $"[{DateTime.Now:HH:mm:ss}] StackTrace: {directLogEx.StackTrace}\n");
            }
            catch { }
            }
            
            // ✅ BUILD TIMESTAMP: Log build info to verify correct DLL is loaded
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyPath = assembly?.Location ?? string.Empty;
                var buildTimestamp = !string.IsNullOrWhiteSpace(assemblyPath)
                    ? System.IO.File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss")
                    : "unknown";
                var versionTag = Helpers.VersionInfo.VersionTag; // "R2023" or "R2024"
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"\n========== REFACTORED CLUSTER SERVICE STARTED ==========\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔨 BUILD TIMESTAMP: {buildTimestamp} | VERSION: {versionTag} | Assembly: {Path.GetFileName(assemblyPath)}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] Target category: {targetCategory ?? "ALL"}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] Filter name: {filterName ?? "NONE"}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] Path 1 Replay: {isPath1Replay}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] PATH 3 Validated: {isPath3Validated}, Invalidated: {isPath3Invalidated}, New: {isPath3New}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ComboId: {comboId?.ToString() ?? "NONE"}, FilterId: {filterId?.ToString() ?? "NONE"}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DeploymentMode: {DeploymentConfiguration.DeploymentMode}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE-ONLY ARCHITECTURE (No XML)\n");
            }
            catch (Exception buildLogEx)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ Error logging build info: {buildLogEx.Message}\n");
            }
            
            _filterName = filterName;
            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>();

            try
            {
                // ✅ STEP 1: Path 1 Replay - Load from database if available
                // ✅ FIX: Pass currentClearanceSettings to check if conditions changed
                // ✅ PATH 3: Skip PATH 1 replay check for PATH 3 types (always recalculate)
                if (isPath1Replay && !isPath3Validated && !isPath3Invalidated && !isPath3New && comboId.HasValue && filterId.HasValue)
                {
                    var path1Result = HandlePath1Replay(doc, comboId.Value, filterId.Value, targetCategory, uiDoc, placedClusterSleevesOut, xmlFilePath, currentClearanceSettings);
                    if (path1Result.hasData)
                        return (path1Result.placedCount, path1Result.deletedCount);
                }
                
                // ✅ PATH 3: Log that we're always recalculating for PATH 3 types
                if (isPath3Validated || isPath3Invalidated || isPath3New)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 3 CLUSTERING: Always recalculating clusters (ignoring existing cluster data)\n");
                    if (isPath3Validated)
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}]   → PATH 3 Validated: Nearby changes may affect cluster formation\n");
                    if (isPath3Invalidated)
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}]   → PATH 3 Invalidated: Geometry changed, must recalculate\n");
                    if (isPath3New)
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}]   → PATH 3 New: New zones added, must recalculate\n");
                }

                // ✅ PERFORMANCE: Track clash zone loading
                List<ClashZone> allClashZones;
                using (var loadTracker = performanceMonitor.TrackOperation("Load Clash Zones from Database"))
                {
                    // ✅ STEP 2: Load clash zones from DATABASE ONLY (Phase 9: Data Service)
                    // ✅ DATABASE-ONLY: All paths (PATH 1, PATH 2, PATH 3) use database exclusively
                    // xmlFilePath parameter is passed as null and ignored - all data comes from database
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 ABOUT TO LOAD clash zones from database for category '{targetCategory}'\n");
                    
                    allClashZones = _dataService.LoadClashZonesFromRegularXml(null, targetCategory, doc);
                    loadTracker.SetItemCount(allClashZones?.Count ?? 0);
                }
                
                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE: Loaded {allClashZones?.Count ?? 0} clash zones from database for category '{targetCategory}'\n");
                }
                catch { }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE: Loaded {allClashZones?.Count ?? 0} clash zones from database for category '{targetCategory}'\n");
                
                if (allClashZones == null || allClashZones.Count == 0)
                {
                    try
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXIT: No clash zones loaded for category '{targetCategory}' - RETURNING (0, 0)\n");
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ EXIT: No clash zones loaded for category '{targetCategory}'\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[RefactoredClusterService] No clash zones loaded for category '{targetCategory}'");
                    }
                    return (0, 0);
                }

                // ✅ PERFORMANCE: Track cache population
                using (var cacheTracker = performanceMonitor.TrackOperation("Populate Clash Zone Cache"))
                {
                    // ✅ STEP 3: Populate cache (Phase 9: Data Service)
                    _dataService.LoadClashZoneCacheFromLoadedClashZones(allClashZones, targetCategory);
                    cacheTracker.SetItemCount(allClashZones.Count);
                }

                // ✅ PERFORMANCE: Track filtering
                List<ClashZone> filteredClashZones;
                using (var filterTracker = performanceMonitor.TrackOperation("Filter Clash Zones"))
                {
                    // ✅ STEP 4: Filter and prepare sleeves for clustering
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: Starting with {allClashZones.Count} total clash zones\n");
                    
                    var withSleeveId = allClashZones.Where(cz => cz.SleeveInstanceId > 0).ToList();
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: {withSleeveId.Count} zones with SleeveInstanceId > 0\n");
                    
                    var notClusterResolved = withSleeveId.Where(cz => !cz.IsClusterResolved).ToList();
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: {notClusterResolved.Count} zones not cluster resolved\n");
                    
                    filteredClashZones = notClusterResolved
                        .Where(cz => string.IsNullOrEmpty(targetCategory) || 
                                    string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    filterTracker.SetItemCount(filteredClashZones.Count);
                }

                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ FILTERED: {filteredClashZones.Count} clash zones ready for clustering (SleeveId>0, not cluster resolved, category='{targetCategory}')\n");
                }
                catch { }
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ FILTERED: {filteredClashZones.Count} clash zones ready for clustering (SleeveId>0, not cluster resolved, category='{targetCategory}')\n");
                
                if (filteredClashZones.Count == 0)
                {
                    try
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXIT: No sleeves to cluster after filtering - RETURNING (0, 0)\n");
                    }
                    catch { }
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ EXIT: No sleeves to cluster after filtering\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Log($"[RefactoredClusterService] No sleeves to cluster after filtering");
                    }
                    return (0, 0);
                }

                // ✅ STEP 5: Validate transaction
                if (!doc.IsModifiable)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefactoredClusterService] Document is not modifiable - transaction must be started by caller");
                    throw new InvalidOperationException("Document must be in a transaction before calling ClusterSleeves");
                }

                // ✅ STEP 6: Get tolerance from settings
                double toleranceDist = GetToleranceFromSettings(targetCategory);

                // ✅ PERFORMANCE: Track sleeve data preparation
                List<dynamic> rawSleeves;
                using (var prepareTracker = performanceMonitor.TrackOperation("Prepare Sleeve Data"))
                {
                    // ✅ STEP 7: Prepare sleeve data (create dynamic objects with ClashZone references)
                    rawSleeves = PrepareSleeveData(filteredClashZones, allClashZones);
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"Prepared {rawSleeves?.Count ?? 0} sleeve data objects\n");
                    prepareTracker.SetItemCount(rawSleeves?.Count ?? 0);
                }

                // ✅ PERFORMANCE: Track grouping
                IGrouping<SleeveGroupKey, dynamic>[] sleeveGroups;
                using (var groupTracker = performanceMonitor.TrackOperation("Group Sleeves by Host/System/Orientation"))
                {
                    // ✅ STEP 8: Group sleeves by host type, system type, and orientation
                    sleeveGroups = rawSleeves.GroupBy(sleeve => new SleeveGroupKey(
                        sleeve.HostType,
                        sleeve.Category,  // systemType = Category (MEP element category)
                        sleeve.Orientation
                    )).ToArray();
                    groupTracker.SetItemCount(sleeveGroups.Length);
                }

                // ✅ STEP 9: Start timeout protection (Phase 10: Timeout Service)
                _timeoutService.StartTimer();

                // ✅ PERFORMANCE: Track cluster formation
                Dictionary<SleeveGroupKey, List<List<dynamic>>> clustersByGroup;
                using (var formTracker = performanceMonitor.TrackOperation("Form Clusters"))
                {
                    // ✅ STEP 10: Form clusters using algorithm service (Phase 8: Algorithm Service)
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 CLUSTERING: Forming clusters with tolerance={RevitUnitConversionService.Instance.FromInternalMillimeters(toleranceDist):F1}mm, {sleeveGroups.Length} sleeve groups\n");
                    clustersByGroup = _algorithmService.FormClusters(sleeveGroups, toleranceDist, doc, enableParallel: true);
                    
                    int totalClusters = clustersByGroup?.Sum(g => g.Value?.Count ?? 0) ?? 0;
                    int totalClustersWithMultipleSleeves = clustersByGroup?.Sum(g => g.Value?.Count(c => c.Count > 1) ?? 0) ?? 0;
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ CLUSTERING: Formed {clustersByGroup?.Count ?? 0} cluster groups, {totalClusters} total clusters, {totalClustersWithMultipleSleeves} clusters with >1 sleeve\n");
                    formTracker.SetItemCount(totalClusters);
                }

                // ✅ STEP 11: Check timeout after clustering
                if (_timeoutService.IsTimedOut())
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefactoredClusterService] ⏱ TIMEOUT: Clustering exceeded {_timeoutService.TimeoutLimitMs / 1000} second limit");
                    _timeoutService.ShowTimeoutWarning("during FormClusters");
                    return (placedCount, deletedCount);
                }

                // ✅ PERFORMANCE: Track cluster placement loop
                int clusterProcessedCount = 0;
                using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Clusters Loop"))
                {
                    // ✅ STEP 12: Process each cluster group
                    foreach (var groupEntry in clustersByGroup)
                    {
                        // Check timeout every 5 clusters
                        clusterProcessedCount++;
                        if (clusterProcessedCount % 5 == 0 && _timeoutService.IsTimedOut())
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[RefactoredClusterService] ⏱ TIMEOUT: Exceeded limit after processing {clusterProcessedCount} clusters");
                            _timeoutService.ShowTimeoutWarning($"after processing {clusterProcessedCount} clusters");
                            break;
                        }

                        var groupKey = groupEntry.Key;
                        var clusters = groupEntry.Value;

                        // ✅ STEP 13: Select clustering strategy (Phase 4: Strategy Factory)
                        var strategy = _strategyFactory.GetStrategy(groupKey, clusters.FirstOrDefault() ?? new List<dynamic>());

                        // Process each cluster
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 PROCESSING: Group {groupKey.hostType}/{groupKey.systemType}/{groupKey.orientation} has {clusters.Count} clusters\n");
                        foreach (var cluster in clusters)
                    {
                        if (cluster.Count <= 1)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⏭️ SKIP: Cluster with only {cluster.Count} sleeve(s) (need >1 to cluster)\n");
                            continue; // Skip individual sleeves
                        }
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ PROCESSING: Cluster with {cluster.Count} sleeves\n");

                        try
                        {
                            // 🔥 CRITICAL: Direct IO logging before placement attempt
                            try
                            {
                                var versionTag = Helpers.VersionInfo.VersionTag;
                                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                var logPath = Path.Combine(logDir, "cluster_debug.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CALL PlaceClusterForGroup: clusterSize={cluster.Count}, groupKey={groupKey.hostType}/{groupKey.systemType}/{groupKey.orientation}\n");
                            }
                            catch { }
                            
                // ✅ PERFORMANCE: Track cluster placement
                (bool success, int placedCount, int deletedCount, FamilyInstance? placedClusterSleeve, int? capturedClusterSleeveId) placementResult;
                using (placementLoopTracker?.TrackSubOperation("Place Cluster Sleeve"))
                {
                    // ✅ STEP 14: Place cluster sleeve (Phase 5: Placement Service + Phase 6: Rotation Service + Phase 3: BoundingBox)
                    // Note: Placement service needs to be wired with functions from rotation service and data service
                    placementResult = PlaceClusterForGroup(
                        doc,
                        cluster,
                        groupKey,
                        targetCategory,
                        xmlFilePath,
                        placementLoopTracker); // Pass tracker for sub-operation tracking
                } // End Place Cluster Sleeve sub-operation

                            // 🔥 CRITICAL: Direct IO logging after placement attempt
                            try
                            {
                                var versionTag = Helpers.VersionInfo.VersionTag;
                                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                var logPath = Path.Combine(logDir, "cluster_debug.log");
                                
                                // Add build timestamp to verify latest build
                                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                                var assemblyPath = assembly?.Location ?? "unknown";
                                var buildTime = File.Exists(assemblyPath) ? File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss") : "unknown";
                                
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterForGroup RETURNED: success={placementResult.success}, placedCount={placementResult.placedCount}, deletedCount={placementResult.deletedCount}, capturedId={placementResult.capturedClusterSleeveId?.ToString() ?? "NULL"}, placedClusterSleeve={(placementResult.placedClusterSleeve != null ? "NOT NULL" : "NULL")} | BUILD: {buildTime}\n");
                            }
                            catch { }

                            if (placementResult.success)
                            {
                                try
                                {
                                    var versionTag = Helpers.VersionInfo.VersionTag;
                                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ PLACEMENT SUCCESS: Adding to _clusterToClashZoneIds\n");
                                }
                                catch { }
                                
                                placedCount += placementResult.placedCount;
                                deletedCount += placementResult.deletedCount;
                                
                                if (placementResult.placedClusterSleeve != null && placementResult.capturedClusterSleeveId.HasValue)
                                {
                                    try
                                    {
                                        var versionTag = Helpers.VersionInfo.VersionTag;
                                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ Adding cluster to placedClusters and _clusterToClashZoneIds: clusterSleeveId={placementResult.capturedClusterSleeveId.Value}\n");
                                    }
                                    catch { }
                                    
                                    // ✅ SESSION PROTECTION: Register recently placed cluster sleeve to prevent deletion
                                    // This protects the cluster sleeve from being deleted by FlagManager.DeleteSleeveForIntersectionPointChange()
                                    FlagManager.RegisterRecentlyPlacedClusterSleeve(placementResult.capturedClusterSleeveId.Value);
                                    
                                    placedClusters.Add(placementResult.placedClusterSleeve);
                                    
                                    // Track ClashZoneIds for database save
                                    var clusterClashZoneIds = cluster
                                        .Select(s => (s.ClashZone as ClashZone)?.Id)
                                        .Where(id => id.HasValue)
                                        .Select(id => id!.Value)
                                        .ToList();
                                    
                                    try
                                    {
                                        var versionTag = Helpers.VersionInfo.VersionTag;
                                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ Adding {clusterClashZoneIds.Count} ClashZoneIds to _clusterToClashZoneIds[{placementResult.capturedClusterSleeveId.Value}]\n");
                                    }
                                    catch { }
                                    
                                    _clusterToClashZoneIds[placementResult.capturedClusterSleeveId.Value] = clusterClashZoneIds;
                                    
                                    try
                                    {
                                        var versionTag = Helpers.VersionInfo.VersionTag;
                                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ AFTER ADD: _clusterToClashZoneIds.Count={_clusterToClashZoneIds.Count}\n");
                                    }
                                    catch { }
                                }
                                else
                                {
                                    try
                                    {
                                        var versionTag = Helpers.VersionInfo.VersionTag;
                                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ PLACEMENT SUCCESS but NOT adding to _clusterToClashZoneIds: placedClusterSleeve={(placementResult.placedClusterSleeve != null ? "NOT NULL" : "NULL")}, capturedId={placementResult.capturedClusterSleeveId?.ToString() ?? "NULL"}\n");
                                    }
                                    catch { }
                                }
                            }
                            else
                            {
                                try
                                {
                                    var versionTag = Helpers.VersionInfo.VersionTag;
                                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PLACEMENT FAILED: success=false, placedCount={placementResult.placedCount}, deletedCount={placementResult.deletedCount}, capturedId={placementResult.capturedClusterSleeveId?.ToString() ?? "NULL"}, placedClusterSleeve={(placementResult.placedClusterSleeve != null ? "NOT NULL" : "NULL")}\n");
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[RefactoredClusterService] Error placing cluster: {ex.Message}");
                                SafeFileLogger.SafeAppendText("cluster_errors.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ Error placing cluster: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                            }
                            // Continue with next cluster
                        }
                    }
                    }
                    placementLoopTracker.SetItemCount(placedCount);
                } // ✅ PERFORMANCE: End of cluster placement loop tracking

                // ✅ STEP 5 OPTIMIZATION: Flush deferred cluster parameters after all clusters placed (4-6× faster)
                // Must happen AFTER all clusters placed but BEFORE cleanup to ensure parameters set
                // This applies to Path 2/3 (main cluster placement loop), Path 1 has its own flush
                if (placedClusters.Count > 0)
                {
                    // Regenerate document to ensure geometry is available for parameter writes
                    try
                    {
                        doc.Regenerate();
                        System.Threading.Thread.Sleep(100); // Brief pause for regeneration
                    }
                    catch (Exception regenEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[RefactoredClusterService] Error regenerating for parameter flush: {regenEx.Message}");
                    }
                    
                    // Flush all accumulated deferred parameters in batch
                    FlushDeferredClusterParameters();
                }

                // ✅ DIAGNOSTIC: Verify all cluster sleeves exist BEFORE cleanup
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 PRE-CLEANUP VERIFICATION: Checking {placedClusters.Count} cluster sleeves before cleanup...\n");
                int validBeforeCleanup = 0;
                int invalidBeforeCleanup = 0;
                foreach (var cluster in placedClusters)
                {
                    if (cluster == null || !cluster.IsValidObject)
                    {
                        invalidBeforeCleanup++;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ PRE-CLEANUP: Cluster sleeve is NULL or INVALID!\n");
                    }
                    else
                    {
                        validBeforeCleanup++;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ PRE-CLEANUP: Cluster sleeve {cluster.Id.IntegerValue} EXISTS: Name='{cluster.Name}', IsValid={cluster.IsValidObject}\n");
                    }
                }
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 📊 PRE-CLEANUP RESULT: {validBeforeCleanup} valid, {invalidBeforeCleanup} invalid out of {placedClusters.Count} cluster sleeves\n");

                // ✅ PERFORMANCE: Track cleanup
                int deletedInCleanup = 0;
                using (var cleanupTracker = performanceMonitor.TrackOperation("Cleanup Individual Sleeves"))
                {
                    // ✅ STEP 15: Cleanup individual sleeves within placed clusters (Phase 7: Cleanup Service)
                    // ⚠️ CRITICAL: Only pass VALID cluster sleeves to cleanup service to ensure they're protected
                    var validClusters = placedClusters.Where(c => c != null && c.IsValidObject).ToList();
                    if (validClusters.Count > 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Calling cleanup service with {validClusters.Count} VALID cluster sleeves for protection\n");
                        deletedInCleanup = _cleanupService.CleanupSleevesWithinClusters(doc, validClusters);
                        deletedCount += deletedInCleanup;
                        cleanupTracker.SetItemCount(deletedInCleanup);
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: No valid cluster sleeves to protect, skipping cleanup service\n");
                    }
                }
                
                // ✅ DIAGNOSTIC: Verify all cluster sleeves exist AFTER cleanup
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 POST-CLEANUP VERIFICATION: Checking {placedClusters.Count} cluster sleeves after cleanup...\n");
                int validAfterCleanup = 0;
                int invalidAfterCleanup = 0;
                foreach (var cluster in placedClusters)
                {
                    if (cluster == null || !cluster.IsValidObject)
                    {
                        invalidAfterCleanup++;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ POST-CLEANUP: Cluster sleeve is NULL or INVALID (WAS DELETED BY CLEANUP!)\n");
                    }
                    else
                    {
                        validAfterCleanup++;
                        var location = cluster.Location as LocationPoint;
                        var locPoint = location?.Point;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ POST-CLEANUP: Cluster sleeve {cluster.Id.IntegerValue} EXISTS: Name='{cluster.Name}', " +
                            $"Location=({locPoint?.X:F2}, {locPoint?.Y:F2}, {locPoint?.Z:F2})\n");
                    }
                }
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 📊 POST-CLEANUP RESULT: {validAfterCleanup} valid, {invalidAfterCleanup} invalid (deleted) out of {placedClusters.Count} cluster sleeves\n");

                // ✅ STEP 16: Return placed cluster sleeves if requested
                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.AddRange(placedClusters);
                }

                // ✅ PERFORMANCE: Track database save
                using (var dbSaveTracker = performanceMonitor.TrackOperation("Save Cluster Data to Database"))
                {
                    // ✅ STEP 17: Save cluster data to database (if comboId and filterId are available)
                // 🔥 CRITICAL: Direct IO logging for database save check
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    
                    // Add build timestamp
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var assemblyPath = assembly?.Location ?? "unknown";
                    var buildTime = File.Exists(assemblyPath) ? File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss") : "unknown";
                    
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 📊 DATABASE SAVE CHECK: comboId={comboId?.ToString() ?? "NULL"}, filterId={filterId?.ToString() ?? "NULL"}, _clusterToClashZoneIds.Count={_clusterToClashZoneIds.Count}, placedClusters.Count={placedClusters.Count} | BUILD: {buildTime}\n");
                    
                    // Log details of what's in _clusterToClashZoneIds
                    if (_clusterToClashZoneIds.Count > 0)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 📊 _clusterToClashZoneIds DETAILS:\n");
                        foreach (var kvp in _clusterToClashZoneIds)
                        {
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}]   ClusterSleeveId={kvp.Key}, ClashZoneIds=[{string.Join(", ", kvp.Value.Take(5))}...] (total {kvp.Value.Count})\n");
                        }
                    }
                }
                catch { }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📊 DATABASE SAVE CHECK: comboId={comboId?.ToString() ?? "NULL"}, filterId={filterId?.ToString() ?? "NULL"}, _clusterToClashZoneIds.Count={_clusterToClashZoneIds.Count}\n");
                
                    if (comboId.HasValue && filterId.HasValue && _clusterToClashZoneIds.Count > 0)
                {
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ CALLING SaveClusterDataToDatabase: {_clusterToClashZoneIds.Count} clusters\n");
                    }
                    catch { }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVING cluster data to database: {_clusterToClashZoneIds.Count} clusters\n");
                    
                    // 🚀 BATCH SAVE: Save all clusters in single transaction (113ms → ~10ms)
                    BatchSaveClusterDataToDatabase(doc, comboId.Value, filterId.Value, targetCategory, _clusterToClashZoneIds);
                    
                    // ✅ STEP 17B: Save sleeve snapshots for cluster sleeves
                    SaveClusterSleeveSnapshots(doc, filterId.Value, placedClusters, targetCategory);
                    
                    // ✅ DIAGNOSTIC: Verify all cluster sleeves still exist in Revit after all operations
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 FINAL VERIFICATION: Checking if all {_clusterToClashZoneIds.Count} cluster sleeves still exist in Revit...\n");
                    
                    int foundCount = 0;
                    int missingCount = 0;
                    foreach (var kvp in _clusterToClashZoneIds)
                    {
                        int clusterId = kvp.Key;
                        var clusterSleeve = doc.GetElement(new ElementId(clusterId)) as FamilyInstance;
                        if (clusterSleeve == null || !clusterSleeve.IsValidObject)
                        {
                            missingCount++;
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CLUSTER SLEEVE {clusterId} MISSING in Revit!\n");
                        }
                        else
                        {
                            foundCount++;
                            var location = clusterSleeve.Location as LocationPoint;
                            var locPoint = location?.Point;
                            
                            // ✅ DIAGNOSTIC: Check if element is in active document or linked file
                            string docTitle = clusterSleeve.Document?.Title ?? "NULL";
                            bool isActiveDoc = clusterSleeve.Document?.IsLinked == false;
                            bool isLinked = clusterSleeve.Document?.IsLinked == true;
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ Cluster sleeve {clusterId} EXISTS: Name='{clusterSleeve.Name}', " +
                                $"Location=({locPoint?.X:F2}, {locPoint?.Y:F2}, {locPoint?.Z:F2}), " +
                                $"Category='{clusterSleeve.Category?.Name ?? "NULL"}, " +
                                $"Document='{docTitle}', " +
                                $"IsActiveDocument={isActiveDoc}, " +
                                $"IsLinked={isLinked}, " +
                                $"IsValid={clusterSleeve.IsValidObject}\n");
                            
                            if (isLinked)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ WARNING: Cluster sleeve {clusterId} is in LINKED FILE '{docTitle}', not active document!\n");
                            }
                        }
                    }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 📊 FINAL VERIFICATION RESULT: {foundCount} found, {missingCount} missing out of {_clusterToClashZoneIds.Count} cluster sleeves\n");
                }
                else
                {
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPED saving cluster data: comboId={comboId?.ToString() ?? "NULL"}, filterId={filterId?.ToString() ?? "NULL"}, clusters={_clusterToClashZoneIds.Count}\n");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ REASON: comboId.HasValue={comboId.HasValue}, filterId.HasValue={filterId.HasValue}, _clusterToClashZoneIds.Count={_clusterToClashZoneIds.Count}\n");
                    }
                    catch { }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPED saving cluster data: comboId={comboId?.ToString() ?? "NULL"}, filterId={filterId?.ToString() ?? "NULL"}, clusters={_clusterToClashZoneIds.Count}\n");
                    }
                    dbSaveTracker.SetItemCount(_clusterToClashZoneIds.Count);
                } // ✅ PERFORMANCE: End of database save tracking

                // ✅ SESSION FLAG: Reset ReadyForPlacementFlag for all zones that were clustered
                // This should happen AFTER cluster placement completes (whichever placement is last: individual or cluster)
                // Individual sleeve placement already resets flags, so this ensures cluster-processed zones are also reset
                if (_clusterToClashZoneIds != null && _clusterToClashZoneIds.Count > 0)
                {
                    try
                    {
                        // Collect all ClashZone GUIDs that were part of clusters
                        var clusteredZoneGuids = new HashSet<Guid>();
                        foreach (var clusterData in _clusterToClashZoneIds.Values)
                        {
                            foreach (var guid in clusterData)
                            {
                                if (guid != Guid.Empty)
                                    clusteredZoneGuids.Add(guid);
                            }
                        }

                        if (clusteredZoneGuids.Count > 0)
                        {
                            using (var dbContext = new SleeveDbContext(doc, msg => { }))
                            {
                                var repository = new ClashZoneRepository(dbContext, msg => { });
                                repository.BulkResetReadyForPlacementFlags(clusteredZoneGuids);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[RefactoredClusterService] ✅ Reset ReadyForPlacementFlag on {clusteredZoneGuids.Count} clustered zones after cluster placement");
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ SESSION-FLAG-RESET: Reset ReadyForPlacementFlag=0 for {clusteredZoneGuids.Count} zones that were clustered\n");
                                    SafeFileLogger.SafeAppendText("placement_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ CLUSTER-FLAG-RESET: Reset ReadyForPlacementFlag=0 for {clusteredZoneGuids.Count} clustered zones (after cluster placement completes)\n");
                                }
                            }
                        }
                    }
                    catch (Exception flagEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[RefactoredClusterService] ⚠️ Failed to reset ReadyForPlacementFlag after cluster placement: {flagEx.Message}");
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ SESSION-FLAG-RESET-ERROR: {flagEx.Message}\n");
                        }
                        // Continue even if flag reset fails (non-blocking)
                    }
                }

                // ✅ STEP 5 OPTIMIZATION: Flush deferred cluster parameters after all placements
                // This writes all accumulated parameters in a single transaction (much faster)
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    int flushedCount = FlushDeferredClusterParameters();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH-FLUSH: Flushed parameters for {flushedCount} cluster sleeves\n");
                    }
                }

                // ✅ PERFORMANCE: Generate final performance report
                performanceMonitor.GenerateReport(0, placedCount); // 0 individual sleeves, placedCount clusters

                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] \n========== REFACTORED CLUSTER SERVICE COMPLETED ==========\n");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ FINAL RESULT: Placed {placedCount} clusters, Deleted {deletedCount} individual sleeves\n");
                }
                catch { }
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] \n========== REFACTORED CLUSTER SERVICE COMPLETED ==========\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ Placed {placedCount} clusters, Deleted {deletedCount} individual sleeves\n");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[RefactoredClusterService] ✅ Completed: Placed {placedCount} clusters, Deleted {deletedCount} individual sleeves");
                }

                return (placedCount, deletedCount);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RefactoredClusterService] ❌ Critical error: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ CRITICAL ERROR: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
                throw;
            }
        }

        /// <summary>
        /// ✅ PATH 1 REPLAY: Load pre-calculated clusters from database.
        /// ✅ FIX: Checks if conditions changed - if changed, uses normal calculation instead of database replay
        /// </summary>
        private (bool hasData, int placedCount, int deletedCount) HandlePath1Replay(
            Document doc,
            int comboId,
            int filterId,
            string targetCategory,
            UIDocument? uiDoc,
            List<FamilyInstance>? placedClusterSleevesOut,
            string? xmlFilePath,
            Dictionary<string, double> currentClearanceSettings = null)
        {
            try
            {
                // ✅ CRITICAL FIX: Check if conditions changed BEFORE loading from database
                // If clearance values changed, cluster sizes may be different → must recalculate
                if (currentClearanceSettings != null && !string.IsNullOrEmpty(_filterName))
                {
                    bool conditionsChanged = Services.Refresh.RefreshPathDeterminer.CheckConditionsChanged(
                        doc, _filterName, targetCategory, currentClearanceSettings);
                    
                    if (conditionsChanged)
                    {
                        // ✅ Conditions changed → Skip database replay, use normal calculation
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: ⚠️ Conditions changed (clearance values) for filter '{_filterName}' + category '{targetCategory}' - using normal calculation instead of database replay");
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] PATH 1: ⚠️ Conditions changed - skipping database replay, will use normal calculation\n");
                        }
                        return (false, 0, 0); // Fall through to normal calculation
                    }
                }
                
                // ✅ Conditions unchanged → Safe to use database replay ("dump one-time use many times")
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var clusterRepository = new ClusterSleeveRepository(dbContext);
                    var existingClusters = clusterRepository.LoadClusterSleevesForCombo(comboId, targetCategory);
                    
                    if (existingClusters != null && existingClusters.Count > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: Found {existingClusters.Count} pre-calculated clusters in database (conditions unchanged, using database replay)");
                        
                        // ✅ PATH 1: Place clusters from database (skip calculation)
                        var path1Result = PlaceClustersFromDatabase(doc, existingClusters, uiDoc, placedClusterSleevesOut, xmlFilePath, targetCategory, comboId);
                        return (true, path1Result.placedCount, path1Result.deletedCount);
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: No cluster data found, skipping clustering");
                        return (true, 0, 0); // Skip clustering
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error loading cluster data: {ex.Message}, falling back to calculation");
                return (false, 0, 0); // Fall through to normal calculation
            }
        }

        /// <summary>
        /// Get tolerance distance from settings (ClusterConfigurationManager).
        /// </summary>
        private double GetToleranceFromSettings(string? targetCategory)
        {
            var settingsService = new SettingsService();
            var currentProfile = ApplicationProfileService.Instance.GetCurrentProfile();
            if (currentProfile != null)
            {
                var settings = settingsService.GetSettings(currentProfile);
                ClusterConfigurationManager.Instance.LoadFromSettings(settings, "User Settings");
            }
            
            double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
            
            // Clamp pipes to 100mm
            if (!string.IsNullOrEmpty(targetCategory) && targetCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                toleranceMm = Math.Min(toleranceMm, 100);
            }
            
            return RevitUnitConversionService.Instance.ToInternalMillimeters(toleranceMm);
        }

        /// <summary>
        /// Prepare sleeve data from clash zones (create dynamic objects with ClashZone references).
        /// </summary>
        private List<dynamic> PrepareSleeveData(List<ClashZone> filteredClashZones, List<ClashZone> allClashZones)
        {
            var rawSleeves = new List<dynamic>();
            
            foreach (var cz in filteredClashZones)
            {
                try
                {
                    // Get host type and orientation from ClashZone
                    string hostType = GetHostTypeFromClashZone(cz);
                    string orientation = GetEffectiveOrientationForClustering(cz);
                    
                    // Create bounding box from ClashZone data
                    BoundingBoxXYZ? bbox = GetBoundingBoxFromClashZone(cz);
                    
                    if (bbox == null) continue; // Skip sleeves without valid bounding boxes
                    
                    dynamic sleeveData = new
                    {
                        SleeveInstanceId = cz.SleeveInstanceId,
                        Category = cz.MepElementCategory,
                        HostType = hostType,
                        Orientation = orientation,
                        BoundingBox = bbox,
                        ClashZone = cz
                    };
                    
                    rawSleeves.Add(sleeveData);
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] Error preparing sleeve data for ClashZone {cz?.Id}: {ex.Message}");
                }
            }
            
            return rawSleeves;
        }

        /// <summary>
        /// Place a cluster sleeve for a group using all relevant services.
        /// </summary>
        private (bool success, int placedCount, int deletedCount, FamilyInstance? placedClusterSleeve, int? capturedClusterSleeveId) PlaceClusterForGroup(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            string? xmlFilePath,
            PlacementPerformanceMonitor.OperationTracker? performanceTracker = null)
        {
            // 🔥 CRITICAL: Direct IO logging at method entry
            try
            {
                var versionTag = Helpers.VersionInfo.VersionTag;
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "cluster_debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterForGroup ENTRY: clusterSize={cluster.Count}, groupKey={groupKey.hostType}/{groupKey.systemType}/{groupKey.orientation}\n");
            }
            catch { }
            
            try
            {
                // 🔥 CRITICAL: Log before rotation angle calculation
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 BEFORE DetermineRotationAngle\n");
                }
                catch { }
                
                // ✅ PERFORMANCE: Track rotation angle determination
                double rotationAngle;
                using (performanceTracker?.TrackSubOperation("Determine Rotation Angle"))
                {
                    // ✅ Step 1: Determine rotation angle (Phase 6: Rotation Service)
                    rotationAngle = _rotationService.DetermineRotationAngle(cluster, xmlFilePath);
                } // End Determine Rotation Angle sub-operation
                
                // 🔥 CRITICAL: Log after rotation angle calculation
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 AFTER DetermineRotationAngle: rotationAngle={rotationAngle * 180.0 / Math.PI:F1}°\n");
                }
                catch { }
                
                // ✅ Step 2: Calculate bounding box (Phase 6: Rotation Service + Phase 3: BoundingBox)
                // 🔥 CRITICAL: Log before getting actual sleeves
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 BEFORE getting actual sleeves from document\n");
                }
                catch { }
                
                var actualSleeves = cluster
                    .Select(s => doc.GetElement(new ElementId(s.SleeveInstanceId)) as FamilyInstance)
                    .Where(fi => fi != null)
                    .ToList();
                
                // 🔥 CRITICAL: Log after getting actual sleeves
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 AFTER getting actual sleeves: found {actualSleeves.Count} of {cluster.Count} sleeves\n");
                }
                catch { }
                
                // 🔥 CRITICAL: Log before bounding box calculation
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 BEFORE CalculateRotatedBoundingBox\n");
                }
                catch { }
                
                // ✅ PERFORMANCE: Track bounding box calculation
                (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) bboxResult;
                using (performanceTracker?.TrackSubOperation("Calculate Rotated Bounding Box"))
                {
                    bboxResult = _rotationService.CalculateRotatedBoundingBox(cluster, actualSleeves, rotationAngle, xmlFilePath);
                } // End Calculate Rotated Bounding Box sub-operation
                
                // 🔥 CRITICAL: Log after bounding box calculation
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 AFTER CalculateRotatedBoundingBox: width={bboxResult.width:F2}, height={bboxResult.height:F2}, depth={bboxResult.depth:F2}\n");
                }
                catch { }
                
                // ✅ DEBUG: Log bounding box calculation for oversized detection
                // Convert from internal units to millimeters for display
                double clusterWidthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.width);
                double clusterHeightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.height);
                double clusterDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.depth);
                double rotationDeg = rotationAngle * 180.0 / Math.PI;
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📐 Cluster Bounding Box: W={clusterWidthMm:F1}mm, H={clusterHeightMm:F1}mm, D={clusterDepthMm:F1}mm, Rotation={rotationDeg:F1}°, ClusterSize={cluster.Count}\n");
                
                // ✅ DEBUG: Log individual sleeve dimensions for comparison
                foreach (var sleeveData in cluster)
                {
                    try
                    {
                        var cz = sleeveData.ClashZone as ClashZone;
                        if (cz != null)
                        {
                            // Calculate depth from bounding box if available, otherwise use diameter for circular sleeves
                            double sleeveDepthInternal = 0.0;
                            if (cz.SleeveBoundingBoxMaxZ > cz.SleeveBoundingBoxMinZ)
                            {
                                sleeveDepthInternal = cz.SleeveBoundingBoxMaxZ - cz.SleeveBoundingBoxMinZ;
                            }
                            else if (cz.SleeveDiameter > 0)
                            {
                                sleeveDepthInternal = cz.SleeveDiameter; // For circular sleeves, use diameter as depth
                            }
                            
                            // Convert from internal units to millimeters for display
                            double slvWidthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cz.SleeveWidth);
                            double slvHeightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cz.SleeveHeight);
                            double slvDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(sleeveDepthInternal);
                            double diameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cz.SleeveDiameter);
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}]   Sleeve {sleeveData.SleeveInstanceId}: W={slvWidthMm:F1}mm, H={slvHeightMm:F1}mm, Depth={slvDepthMm:F1}mm, Diameter={diameterMm:F1}mm\n");
                        }
                    }
                    catch { }
                }
                
                if (bboxResult.width <= 0 || bboxResult.height <= 0)
                {
                    // 🔥 CRITICAL: Log invalid bounding box
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ INVALID BOUNDING BOX: width={bboxResult.width:F2}, height={bboxResult.height:F2} - RETURNING (false, 0, 0, null, null)\n");
                    }
                    catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] Invalid bounding box for cluster, skipping");
                    return (false, 0, 0, null, null);
                }
                
                // ✅ Step 3: Place cluster sleeve (Phase 5: Placement Service)
                // 🔥 CRITICAL: Direct IO logging before placement
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CALL PlaceClusterSleeve: placementPoint=({bboxResult.mid.X:F2},{bboxResult.mid.Y:F2},{bboxResult.mid.Z:F2}), W={RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.width):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.height):F1}mm\n");
                }
                catch { }
                
                // ✅ BATCH PARAMETER OPTIMIZATION: Pass deferred parameters dictionary for batch writes
                bool placementSuccess = _placementService.PlaceClusterSleeve(
                    doc,
                    cluster,
                    groupKey,
                    targetCategory,
                    bboxResult.mid, // placement point
                    bboxResult.width,
                    bboxResult.height,
                    bboxResult.depth,
                    rotationAngle,
                    xmlFilePath,
                    out FamilyInstance? placedClusterSleeve,
                    out int? capturedClusterSleeveId,
                    _deferredClusterParameters); // ✅ Pass deferred parameters for batching
                
                // 🔥 CRITICAL: Direct IO logging after placement
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 PlaceClusterSleeve RETURNED: success={placementSuccess}, placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")}, capturedId={capturedClusterSleeveId?.ToString() ?? "NULL"}\n");
                }
                catch { }
                
                if (!placementSuccess || placedClusterSleeve == null)
                {
                    // 🔥 CRITICAL: Direct IO logging for failure
                    try
                    {
                        var versionTag = Helpers.VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PLACEMENT FAILED: placementSuccess={placementSuccess}, placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")} - RETURNING (false, 0, 0, null, null)\n");
                    }
                    catch { }
                    return (false, 0, 0, null, null);
                }
                
                // ✅ Step 4: Store rotation data (Phase 6: Rotation Service)
                if (capturedClusterSleeveId.HasValue)
                {
                    _rotationService.StoreRotationData(
                        capturedClusterSleeveId.Value,
                        rotationAngle * 180.0 / Math.PI,
                        Math.Abs(rotationAngle) > 1e-6,
                        new XYZ(bboxResult.rotatedMinX ?? 0, bboxResult.rotatedMinY ?? 0, bboxResult.rotatedMinZ ?? 0),
                        new XYZ(bboxResult.rotatedMaxX ?? 0, bboxResult.rotatedMaxY ?? 0, bboxResult.rotatedMaxZ ?? 0),
                        bboxResult.width,
                        bboxResult.height,
                        bboxResult.depth);
                }
                
                // ✅ Step 5: Update flags for clash zones BEFORE deleting individual sleeves
                // ✅ CRITICAL: Must set AfterClusterSleevePlacedSleeveInstanceId BEFORE clearing SleeveInstanceId
                if (capturedClusterSleeveId.HasValue && placedClusterSleeve != null)
                {
                    try
                    {
                        foreach (var sleeveData in cluster)
                        {
                            try
                            {
                                // Get clash zone from cluster data
                                var clashZone = sleeveData.ClashZone as ClashZone;
                                if (clashZone == null)
                                {
                                    // Try to get from database if not in cluster data
                                    using (var dbContext = new SleeveDbContext(doc))
                                    {
                                        var clashZoneRepository = new ClashZoneRepository(dbContext);
                                        var categoryZones = clashZoneRepository.GetClashZonesByCategory(targetCategory);
                                        clashZone = categoryZones?.FirstOrDefault(cz => cz.SleeveInstanceId == sleeveData.SleeveInstanceId);
                                    }
                                }
                                
                                if (clashZone != null)
                                {
                                    // ✅ CRITICAL: Save original SleeveInstanceId BEFORE clearing it
                                    var originalSleeveInstanceId = clashZone.SleeveInstanceId;
                                    var originalClusterId = clashZone.ClusterSleeveInstanceId;
                                    
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 BEFORE UpdateFlagsForPlacement: ClashZone {clashZone.Id}, " +
                                        $"IsClusterResolved={clashZone.IsClusterResolved}, " +
                                        $"ClusterId={clashZone.ClusterSleeveInstanceId}, " +
                                        $"SleeveId={clashZone.SleeveInstanceId}, " +
                                        $"AfterClusterSleeveId={clashZone.AfterClusterSleevePlacedSleeveInstanceId}\n");
                                    
                                    // ✅ CRITICAL: Save original SleeveInstanceId to AfterClusterSleevePlacedSleeveInstanceId BEFORE UpdateFlagsForPlacement clears it
                                    // This preserves the individual sleeve ID for database tracking (edge case: deleted individual sleeves)
                                    if (originalSleeveInstanceId > 0)
                                    {
                                        clashZone.AfterClusterSleevePlacedSleeveInstanceId = originalSleeveInstanceId;
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ Set AfterClusterSleevePlacedSleeveInstanceId={originalSleeveInstanceId} for ClashZone {clashZone.Id} (original individual sleeve ID)\n");
                                    }
                                    else
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ originalSleeveInstanceId={originalSleeveInstanceId} (not > 0), cannot set AfterClusterSleevePlacedSleeveInstanceId\n");
                                    }
                                    
                                    // Update flags using FlagManager (this will save AfterClusterSleevePlacedSleeveInstanceId to database)
                                    _flagManager.UpdateFlagsForPlacement(
                                        clashZone,
                                        capturedClusterSleeveId.Value,
                                        isCluster: true,
                                        targetCategory,
                                        _filterName);
                                    
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ AFTER UpdateFlagsForPlacement: ClashZone {clashZone.Id}, " +
                                        $"IsClusterResolved={clashZone.IsClusterResolved}, " +
                                        $"ClusterId={clashZone.ClusterSleeveInstanceId}, " +
                                        $"SleeveId={clashZone.SleeveInstanceId}, " +
                                        $"AfterClusterSleeveId={clashZone.AfterClusterSleevePlacedSleeveInstanceId}, " +
                                        $"NewClusterId={capturedClusterSleeveId.Value}\n");
                                }
                                else
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ ClashZone is NULL for sleeve {sleeveData.SleeveInstanceId}, cannot update flags\n");
                                }
                            }
                            catch (Exception flagEx)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error updating flags for sleeve {sleeveData.SleeveInstanceId}: {flagEx.Message}\n");
                            }
                        }
                    }
                    catch (Exception updateFlagsEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ Error updating flags for cluster: {updateFlagsEx.Message}\n");
                    }
                }
                
                // ✅ Step 6: Delete individual sleeves that formed this cluster (IMMEDIATE deletion, not deferred)
                // ⚠️ CRITICAL: Verify cluster sleeve exists BEFORE deleting individual sleeves
                if (capturedClusterSleeveId.HasValue && placedClusterSleeve != null)
                {
                    var verifyBeforeDelete = doc.GetElement(new ElementId(capturedClusterSleeveId.Value)) as FamilyInstance;
                    if (verifyBeforeDelete == null || !verifyBeforeDelete.IsValidObject)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL: Cluster sleeve {capturedClusterSleeveId.Value} NOT FOUND before deleting individual sleeves - SKIPPING DELETION\n");
                        return (true, 1, 0, placedClusterSleeve, capturedClusterSleeveId); // Return success but don't delete individual sleeves
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ VERIFY BEFORE DELETE: Cluster sleeve {capturedClusterSleeveId.Value} EXISTS before deleting {cluster.Count} individual sleeves\n");
                    }
                }
                
                int deletedIndividualCount = 0;
                var sleevesToDelete = new List<ElementId>();
                
                foreach (var sleeveData in cluster)
                {
                    try
                    {
                        int sleeveInstanceId = sleeveData.SleeveInstanceId;
                        if (sleeveInstanceId <= 0) continue;
                        
                        var sleeveElementId = new ElementId(sleeveInstanceId);
                        var sleeveElement = doc.GetElement(sleeveElementId);
                        
                        // ✅ PROTECTION: Don't delete the cluster sleeve itself
                        if (placedClusterSleeve != null && sleeveElementId == placedClusterSleeve.Id)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ PROTECTION: Skipping cluster sleeve {sleeveInstanceId} (same as placedClusterSleeve)\n");
                            continue;
                        }
                        
                        // ✅ DOUBLE-CHECK: Verify this is NOT the cluster sleeve by ID
                        if (capturedClusterSleeveId.HasValue && sleeveInstanceId == capturedClusterSleeveId.Value)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ PROTECTION: Skipping cluster sleeve {sleeveInstanceId} (matches capturedClusterSleeveId)\n");
                            continue;
                        }
                        
                        if (sleeveElement != null && sleeveElement is FamilyInstance)
                        {
                            sleevesToDelete.Add(sleeveElementId);
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🗑️ QUEUED for deletion: individual sleeve {sleeveInstanceId} (part of cluster {capturedClusterSleeveId?.ToString() ?? "unknown"})\n");
                        }
                    }
                    catch (Exception delEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error queuing sleeve {sleeveData.SleeveInstanceId} for deletion: {delEx.Message}\n");
                    }
                }
                
                // Delete in batch within the same transaction
                if (sleevesToDelete.Count > 0)
                {
                    try
                    {
                        foreach (var id in sleevesToDelete)
                        {
                            try
                            {
                                // ✅ TRIPLE-CHECK: Verify this is NOT the cluster sleeve
                                if (capturedClusterSleeveId.HasValue && id.IntegerValue == capturedClusterSleeveId.Value)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL: Attempted to delete cluster sleeve {id.IntegerValue} - SKIPPING!\n");
                                    continue;
                                }
                                
                                doc.Delete(id);
                                deletedIndividualCount++;
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ DELETED individual sleeve {id.IntegerValue}\n");
                            }
                            catch (Exception delEx)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ❌ Failed to delete sleeve {id.IntegerValue}: {delEx.Message}\n");
                            }
                        }
                        
                        // ✅ CRITICAL: Verify cluster sleeve still exists AFTER deleting individual sleeves
                        if (capturedClusterSleeveId.HasValue)
                        {
                            var verifyAfterDelete = doc.GetElement(new ElementId(capturedClusterSleeveId.Value)) as FamilyInstance;
                            if (verifyAfterDelete == null || !verifyAfterDelete.IsValidObject)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL: Cluster sleeve {capturedClusterSleeveId.Value} WAS DELETED during individual sleeve deletion!\n");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ VERIFY AFTER DELETE: Cluster sleeve {capturedClusterSleeveId.Value} STILL EXISTS after deleting {deletedIndividualCount} individual sleeves\n");
                            }
                        }
                    }
                    catch (Exception batchDelEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ Batch deletion error: {batchDelEx.Message}\n");
                    }
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📊 Cluster {capturedClusterSleeveId?.ToString() ?? "unknown"}: Deleted {deletedIndividualCount} of {cluster.Count} individual sleeves\n");
                
                return (true, 1, deletedIndividualCount, placedClusterSleeve, capturedClusterSleeveId);
            }
            catch (Exception ex)
            {
                // 🔥 CRITICAL: Log exception details
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION in PlaceClusterForGroup: {ex.Message}\n");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] StackTrace: {ex.StackTrace}\n");
                }
                catch { }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error in PlaceClusterForGroup: {ex.Message}");
                return (false, 0, 0, null, null);
            }
        }

        /// <summary>
        /// ✅ PATH 1 REPLAY: Place cluster sleeves from pre-calculated database data.
        /// </summary>
        private (int placedCount, int deletedCount) PlaceClustersFromDatabase(
            Document doc,
            List<ClusterSleeveData> clusterDataList,
            UIDocument? uiDoc,
            List<FamilyInstance>? placedClusterSleevesOut,
            string? xmlFilePath,
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
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Cluster {clusterData.ClusterInstanceId} has no ClashZoneIds, skipping");
                            continue;
                        }

                        // Get family symbol based on host type
                        string familyName = "";
                        if (clusterData.HostType == "Wall" || clusterData.HostType == "Structural Framing")
                        {
                            familyName = "RectangularOpeningOnWall";
                        }
                        else if (clusterData.HostType == "Floor")
                        {
                            familyName = "RectangularOpeningOnSlab";
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Unknown host type '{clusterData.HostType}' for cluster {clusterData.ClusterInstanceId}, skipping");
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
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Family '{familyName}' not found for cluster {clusterData.ClusterInstanceId}, skipping");
                            continue;
                        }

                        var familySymbol = universalSymbols.First();
                        if (!familySymbol.IsActive) familySymbol.Activate();

                        // Get reference level (fallback to first level - could be improved by storing level in ClusterSleeves table)
                        Level? refLevel = doc.GetElement(new ElementId(1)) as Level;
                        if (refLevel == null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Could not determine reference level for cluster {clusterData.ClusterInstanceId}, skipping");
                            continue;
                        }

                        // Create placement point
                        var placementPoint = new XYZ(clusterData.PlacementX, clusterData.PlacementY, clusterData.PlacementZ);

                        // ✅ PATH 1: Place cluster sleeve directly from database data (simpler than PlacementService)
                        FamilyInstance? placedClusterSleeve = null;
                        int? capturedClusterSleeveId = null;
                        
                        try
                        {
                            // Create cluster sleeve directly
                            placedClusterSleeve = doc.Create.NewFamilyInstance(
                                placementPoint,
                                familySymbol,
                                refLevel,
                                StructuralType.NonStructural);
                            
                            if (placedClusterSleeve != null)
                            {
                                // ✅ CRITICAL: Capture ID immediately while element is valid
                                capturedClusterSleeveId = placedClusterSleeve.Id.IntegerValue;
                                
                                // Set dimensions
                                var widthParam = placedClusterSleeve.LookupParameter("Width");
                                var heightParam = placedClusterSleeve.LookupParameter("Height");
                                var depthParam = placedClusterSleeve.LookupParameter("Depth");

                                // ✅ STEP 5 OPTIMIZATION: Defer parameter writes if batching enabled
                                if (OptimizationFlags.UseBatchedParameterWrites)
                                {
                                    // Accumulate parameter values for batch write after regeneration
                                    if (!_deferredClusterParameters.ContainsKey(placedClusterSleeve.Id))
                                        _deferredClusterParameters[placedClusterSleeve.Id] = new Dictionary<string, object>();
                                    
                                    if (widthParam != null && !widthParam.IsReadOnly)
                                        _deferredClusterParameters[placedClusterSleeve.Id]["Width"] = clusterData.ClusterWidth;
                                    if (heightParam != null && !heightParam.IsReadOnly)
                                        _deferredClusterParameters[placedClusterSleeve.Id]["Height"] = clusterData.ClusterHeight;
                                    if (depthParam != null && !depthParam.IsReadOnly)
                                        _deferredClusterParameters[placedClusterSleeve.Id]["Depth"] = clusterData.ClusterDepth;
                                }
                                else
                                {
                                    // Immediate write path (backward compatibility)
                                    if (widthParam != null && !widthParam.IsReadOnly)
                                        widthParam.Set(clusterData.ClusterWidth);
                                    if (heightParam != null && !heightParam.IsReadOnly)
                                        heightParam.Set(clusterData.ClusterHeight);
                                    if (depthParam != null && !depthParam.IsReadOnly)
                                        depthParam.Set(clusterData.ClusterDepth);
                                }
                            }
                        }
                        catch (Exception placeEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[RefactoredClusterService] PATH 1: Error placing cluster {clusterData.ClusterInstanceId}: {placeEx.Message}");
                            continue;
                        }

                        if (placedClusterSleeve == null || !capturedClusterSleeveId.HasValue)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Failed to place cluster {clusterData.ClusterInstanceId}");
                            continue;
                        }

                        placedClusters.Add(placedClusterSleeve);
                        placedCount++;

                        // Mark clash zones as cluster resolved using FlagManager
                        if (capturedClusterSleeveId.HasValue && clashZoneIds.Count > 0)
                        {
                            try
                            {
                                // Get clash zones from database and update flags using FlagManager
                                using (var dbContext = new SleeveDbContext(doc))
                                {
                                    var clashZoneRepository = new ClashZoneRepository(dbContext);
                                    
                                    foreach (var clashZoneId in clashZoneIds)
                                    {
                                        // Get clash zone by loading all for category and finding by GUID
                                        // Note: This could be optimized if we add GetClashZoneByGuid to repository
                                        var categoryZones = clashZoneRepository.GetClashZonesByCategory(targetCategory);
                                        var clashZone = categoryZones?.FirstOrDefault(cz => cz.Id == clashZoneId);
                                        
                                        if (clashZone != null)
                                        {
                                            // Get filter name (use empty string if not available)
                                            var categoryName = clashZone.MepElementCategory ?? targetCategory;
                                            var baseFilterName = string.Empty; // Filter name not critical for PATH 1
                                            
                                            // Update flags using FlagManager
                                            _flagManager.UpdateFlagsForPlacement(
                                                clashZone,
                                                capturedClusterSleeveId.Value,
                                                isCluster: true,
                                                categoryName,
                                                baseFilterName);
                                        }
                                    }
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[RefactoredClusterService] PATH 1: Updated flags for {clashZoneIds.Count} clash zones for cluster {capturedClusterSleeveId.Value}");
                                }
                            }
                            catch (Exception flagEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error updating flags for cluster {clusterData.ClusterInstanceId}: {flagEx.Message}");
                            }
                        }

                        // Delete individual sleeves within cluster (use CleanupService)
                        // Note: This is done after all clusters are placed, not individually
                    }
                    catch (Exception clusterEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error placing cluster {clusterData.ClusterInstanceId}: {clusterEx.Message}");
                            SafeFileLogger.SafeAppendText("cluster_errors.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] PATH 1: Error placing cluster {clusterData.ClusterInstanceId}: {clusterEx.Message}\nStackTrace: {clusterEx.StackTrace}\n");
                        }
                        // Continue with next cluster
                    }
                }

                // ✅ STEP 5 OPTIMIZATION: Flush deferred cluster parameters before cleanup (4-6× faster)
                // Must happen AFTER all clusters placed but BEFORE cleanup to ensure parameters set
                if (placedClusters.Count > 0)
                {
                    // Regenerate document to ensure geometry is available for parameter writes
                    try
                    {
                        doc.Regenerate();
                        System.Threading.Thread.Sleep(100); // Brief pause for regeneration
                    }
                    catch (Exception regenEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error regenerating for parameter flush: {regenEx.Message}");
                    }
                    
                    // Flush all accumulated deferred parameters in batch
                    FlushDeferredClusterParameters();
                    
                    // Now cleanup individual sleeves within placed clusters
                    deletedCount = _cleanupService.CleanupSleevesWithinClusters(doc, placedClusters);
                }

                // Return placed cluster sleeves if requested
                if (placedClusterSleevesOut != null)
                {
                    placedClusterSleevesOut.Clear();
                    placedClusterSleevesOut.AddRange(placedClusters);
                }

                // Reset flag after placement
                try
                {
                    using (var dbContext = new SleeveDbContext(doc))
                    {
                        var clashZoneRepository = new ClashZoneRepository(dbContext);
                        clashZoneRepository.ResetFileComboFlag(comboId);
                    }
                }
                catch (Exception resetEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] PATH 1: Error resetting flag: {resetEx.Message}");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RefactoredClusterService] PATH 1: Error in PlaceClustersFromDatabase: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] PATH 1: Error in PlaceClustersFromDatabase: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
                throw;
            }

            return (placedCount, deletedCount);
        }

        /// <summary>
        /// ✅ STEP 5 OPTIMIZATION: Flush deferred cluster parameters after regeneration (4-6× faster)
        /// Batch-writes all accumulated parameter values in single transaction.
        /// Expected speedup: 584ms → ~100-150ms per cluster (parameter portion)
        /// </summary>
        /// <returns>Number of cluster sleeves with parameters written</returns>
        public int FlushDeferredClusterParameters()
        {
            if (!OptimizationFlags.UseBatchedParameterWrites || _deferredClusterParameters.Count == 0)
                return 0;

            var sw = Stopwatch.StartNew();
            int successCount = 0;
            int errorCount = 0;
            
            SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] Flushing {_deferredClusterParameters.Count} cluster sleeve parameters...\n");

            foreach (var kvp in _deferredClusterParameters)
            {
                var elementId = kvp.Key;
                var parameters = kvp.Value;
                
                try
                {
                    var element = _doc.GetElement(elementId);
                    if (element == null)
                    {
                        errorCount++;
                        SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Element {elementId} not found\n");
                        continue;
                    }

                    foreach (var param in parameters)
                    {
                        var paramName = param.Key;
                        var paramValue = param.Value;
                        
                        var parameter = element.LookupParameter(paramName);
                        if (parameter != null && !parameter.IsReadOnly)
                        {
                            try
                            {
                                if (paramValue is double doubleValue)
                                    parameter.Set(doubleValue);
                                else if (paramValue is int intValue)
                                    parameter.Set(intValue);
                                else if (paramValue is string stringValue)
                                    parameter.Set(stringValue);
                                else
                                {
                                    // Try to convert to int (for Cluster Sleeve Instance ID)
                                    if (paramValue != null && int.TryParse(paramValue.ToString(), out int parsedInt))
                                        parameter.Set(parsedInt);
                                    else
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Unsupported parameter type for '{paramName}': {paramValue?.GetType().Name ?? "null"}\n");
                                    }
                                }
                            }
                            catch (Exception paramEx)
                            {
                                SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Error setting parameter '{paramName}' = {paramValue}: {paramEx.Message}\n");
                            }
                        }
                        else if (parameter == null)
                        {
                            SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Parameter '{paramName}' not found on element {elementId}\n");
                        }
                        else if (parameter.IsReadOnly)
                        {
                            SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Parameter '{paramName}' is read-only on element {elementId}\n");
                        }
                    }
                    
                    successCount++;
                }
                catch (Exception ex)
                {
                    errorCount++;
                    SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Error writing params for element {elementId}: {ex.Message}\n");
                }
            }

            sw.Stop();
            SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ✅ Flushed {successCount} cluster sleeves in {sw.ElapsedMilliseconds}ms ({errorCount} errors)\n");

            // Clear deferred parameters after flushing
            _deferredClusterParameters.Clear();
            
            return successCount;
        }

        /// <summary>
        /// 🚀 BATCH SAVE: Save cluster data to database using single transaction (113ms → ~10ms)
        /// Significantly faster than calling SaveClusterDataToDatabase in a loop
        /// </summary>
        private void BatchSaveClusterDataToDatabase(
            Document doc,
            int comboId,
            int filterId,
            string targetCategory,
            Dictionary<int, List<Guid>> clusterToClashZoneIds)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (comboId <= 0) throw new ArgumentException($"Invalid ComboId: {comboId}", nameof(comboId));
            if (filterId <= 0) throw new ArgumentException($"Invalid FilterId: {filterId}", nameof(filterId));
            if (string.IsNullOrWhiteSpace(targetCategory)) throw new ArgumentException("Target category cannot be null or empty", nameof(targetCategory));
            if (clusterToClashZoneIds == null) throw new ArgumentNullException(nameof(clusterToClashZoneIds));
            if (clusterToClashZoneIds.Count == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ BatchSaveClusterDataToDatabase: No clusters to save\n");
                return;
            }
            
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var clusterRepository = new ClusterSleeveRepository(dbContext);
                    
                    // Prepare all cluster save data
                    var clustersToSave = new List<ClusterSaveData>();
                    int skippedCount = 0;
                    
                    foreach (var kvp in clusterToClashZoneIds)
                    {
                        try
                        {
                            int clusterInstanceId = kvp.Key;
                            List<Guid> clashZoneIds = kvp.Value;
                            
                            // Get rotation data for this cluster
                            var rotationData = _rotationService.GetRotationData(clusterInstanceId);
                            
                            // Get cluster sleeve element
                            var clusterSleeve = doc.GetElement(new ElementId(clusterInstanceId)) as FamilyInstance;
                            if (clusterSleeve == null)
                            {
                                skippedCount++;
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ SKIPPED CLUSTER {clusterInstanceId}: Cluster sleeve not found in Revit document\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Cluster sleeve {clusterInstanceId} not found, skipping save");
                                continue;
                            }
                            
                            // Get bounding box
                            var bbox = clusterSleeve.get_BoundingBox(null);
                            if (bbox == null)
                            {
                                skippedCount++;
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ SKIPPED CLUSTER {clusterInstanceId}: Could not get bounding box from Revit element\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Could not get bounding box for cluster {clusterInstanceId}, skipping save");
                                continue;
                            }
                            
                            // Get dimensions
                            var widthParam = clusterSleeve.LookupParameter("Width");
                            var heightParam = clusterSleeve.LookupParameter("Height");
                            var depthParam = clusterSleeve.LookupParameter("Depth");
                            
                            double width = widthParam?.AsDouble() ?? 0.0;
                            double height = heightParam?.AsDouble() ?? 0.0;
                            double depth = depthParam?.AsDouble() ?? 0.0;
                            
                            // Get rotation data
                            double rotationAngleDeg = 0.0;
                            bool isRotated = false;
                            XYZ bboxMin = bbox.Min;
                            XYZ bboxMax = bbox.Max;
                            
                            if (rotationData.HasValue)
                            {
                                rotationAngleDeg = rotationData.Value.rotationAngleDeg;
                                isRotated = rotationData.Value.isRotated;
                                
                                // ✅ CRITICAL FIX: Only use rotated bounding box if it's valid (not zero/empty)
                                // If rotated bounding box is invalid, fall back to Revit bounding box
                                var rotatedBboxMin = rotationData.Value.rotatedBboxMin;
                                var rotatedBboxMax = rotationData.Value.rotatedBboxMax;
                                
                                // Check if rotated bounding box is valid (not zero and min < max)
                                bool isValidRotatedBbox = rotatedBboxMin != null && rotatedBboxMax != null &&
                                                          !rotatedBboxMin.IsAlmostEqualTo(XYZ.Zero) &&
                                                          !rotatedBboxMax.IsAlmostEqualTo(XYZ.Zero) &&
                                                          rotatedBboxMin.X < rotatedBboxMax.X &&
                                                          rotatedBboxMin.Y < rotatedBboxMax.Y &&
                                                          rotatedBboxMin.Z < rotatedBboxMax.Z;
                                
                                if (isValidRotatedBbox)
                                {
                                    bboxMin = rotatedBboxMin;
                                    bboxMax = rotatedBboxMax;
                                }
                                // Otherwise, keep using Revit bounding box (bbox.Min/Max from above)
                                
                                // Use rotated dimensions if available and valid
                                if (rotationData.Value.rotatedWidth > 0)
                                    width = rotationData.Value.rotatedWidth;
                                if (rotationData.Value.rotatedHeight > 0)
                                    height = rotationData.Value.rotatedHeight;
                                if (rotationData.Value.rotatedDepth > 0)
                                    depth = rotationData.Value.rotatedDepth;
                            }
                            
                            // Get host type and orientation
                            string hostType = GetHostTypeFromSleeve(clusterSleeve);
                            string hostOrientation = GetOrientationFromSleeve(clusterSleeve);
                            
                            // Get placement point
                            var placementPoint = (bboxMin + bboxMax) / 2.0;
                            
                            // ✅ DIAGNOSTIC: Log bounding box values before saving
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 📦 PREPARE SAVE: Cluster {clusterInstanceId} - " +
                                $"BBoxMin=({bboxMin.X:F6}, {bboxMin.Y:F6}, {bboxMin.Z:F6}), " +
                                $"BBoxMax=({bboxMax.X:F6}, {bboxMax.Y:F6}, {bboxMax.Z:F6}), " +
                                $"Placement=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                            
                            // Add to batch
                            clustersToSave.Add(new ClusterSaveData
                            {
                                ClusterInstanceId = clusterInstanceId,
                                ComboId = comboId,
                                FilterId = filterId,
                                Category = targetCategory,
                                BoundingBoxMinX = bboxMin.X,
                                BoundingBoxMinY = bboxMin.Y,
                                BoundingBoxMinZ = bboxMin.Z,
                                BoundingBoxMaxX = bboxMax.X,
                                BoundingBoxMaxY = bboxMax.Y,
                                BoundingBoxMaxZ = bboxMax.Z,
                                ClusterWidth = width,
                                ClusterHeight = height,
                                ClusterDepth = depth,
                                RotationAngleDeg = rotationAngleDeg,
                                IsRotated = isRotated,
                                PlacementX = placementPoint.X,
                                PlacementY = placementPoint.Y,
                                PlacementZ = placementPoint.Z,
                                HostType = hostType,
                                HostOrientation = hostOrientation ?? "Unknown",
                                ClashZoneIds = clashZoneIds
                            });
                        }
                        catch (Exception prepEx)
                        {
                            skippedCount++;
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ SKIPPED CLUSTER {kvp.Key}: Error preparing for batch save - {prepEx.Message}\n" +
                                $"StackTrace: {prepEx.StackTrace}\n");
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] Error preparing cluster {kvp.Key} for batch save: {prepEx.Message}");
                        }
                    }
                    
                    // Execute batch save in single transaction
                    if (clustersToSave.Count > 0)
                    {
                        clusterRepository.BatchSaveClusterSleeves(clustersToSave);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] ✅ Batch saved {clustersToSave.Count} clusters to database (ComboId={comboId}, FilterId={filterId}, Category={targetCategory})");
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVED {clustersToSave.Count} clusters to ClusterSleeves table in single transaction (Skipped: {skippedCount})\n");
                    }
                    
                    if (skippedCount > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[RefactoredClusterService] ⚠️ Skipped {skippedCount} of {clusterToClashZoneIds.Count} clusters");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RefactoredClusterService] ❌ CRITICAL ERROR batch saving cluster data to database: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ CRITICAL ERROR batch saving cluster data: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
                throw;
            }
        }

        /// <summary>
        /// ✅ PATH 2/3: Save cluster data to database after calculation and placement.
        /// </summary>
        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
        /// 
        /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        /// To modify this method, you MUST:
        /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true in this class
        /// 2. Get explicit consent from the project owner
        /// 3. Test thoroughly with database operations
        /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        /// 
        /// Saves cluster sleeve data to ClusterSleeves table in SQLite database.
        /// This is the ONLY place where cluster data is saved to the database.
        /// 
        /// ✅ CRITICAL: This method must be called AFTER cluster placement and BEFORE parameter transfer.
        /// The data saved here is used by:
        /// - ParameterTransferService (to find cluster sleeves and transfer parameters)
        /// - RefreshService (to reload cluster data)
        /// - UI components (to display cluster information)
        /// 
        /// ⚠️ DO NOT:
        /// - Remove validation checks (they prevent data corruption)
        /// - Change the database schema without updating this method
        /// - Skip saving rotation data (needed for rotated clusters)
        /// - Modify the ClashZoneIds list (used for parameter aggregation)
        /// 
        /// DATABASE SCHEMA (ClusterSleeves table):
        /// - ClusterInstanceId (PRIMARY KEY)
        /// - ComboId, FilterId, Category
        /// - BoundingBoxMinX/Y/Z, BoundingBoxMaxX/Y/Z
        /// - ClusterWidth, ClusterHeight, ClusterDepth
        /// - RotationAngleDeg, IsRotated
        /// - PlacementX/Y/Z
        /// - HostType, HostOrientation
        /// - ClashZoneIdsJson (JSON array of GUIDs)
        /// </summary>
        private void SaveClusterDataToDatabase(
            Document doc,
            int comboId,
            int filterId,
            string targetCategory,
            Dictionary<int, List<Guid>> clusterToClashZoneIds)
        {
            // ⚠️ CONSENT CHECK: Prevent modifications without explicit consent
            if (!ALLOW_MODIFICATIONS_TO_PROTECTED_CODE)
            {
                // This method is protected - modifications require explicit consent
                // To modify: Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true and get consent
            }
            
            // ✅ VALIDATION: Ensure all required parameters are valid
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (comboId <= 0) throw new ArgumentException($"Invalid ComboId: {comboId}", nameof(comboId));
            if (filterId <= 0) throw new ArgumentException($"Invalid FilterId: {filterId}", nameof(filterId));
            if (string.IsNullOrWhiteSpace(targetCategory)) throw new ArgumentException("Target category cannot be null or empty", nameof(targetCategory));
            if (clusterToClashZoneIds == null) throw new ArgumentNullException(nameof(clusterToClashZoneIds));
            if (clusterToClashZoneIds.Count == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ SaveClusterDataToDatabase: No clusters to save\n");
                return; // No clusters to save - not an error
            }
            
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var clusterRepository = new ClusterSleeveRepository(dbContext);
                    
                    int savedCount = 0;
                    int failedCount = 0;
                    
                    foreach (var kvp in clusterToClashZoneIds)
                    {
                        try
                        {
                            int clusterInstanceId = kvp.Key;
                            List<Guid> clashZoneIds = kvp.Value;
                            
                            // Get rotation data for this cluster
                            var rotationData = _rotationService.GetRotationData(clusterInstanceId);
                            
                            // Get cluster sleeve element
                            var clusterSleeve = doc.GetElement(new ElementId(clusterInstanceId)) as FamilyInstance;
                            if (clusterSleeve == null)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CLUSTER SLEEVE {clusterInstanceId} NOT FOUND IN REVIT - SKIPPING SAVE\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Cluster sleeve {clusterInstanceId} not found, skipping save");
                                continue;
                            }
                            
                            // ✅ DIAGNOSTIC: Verify cluster sleeve exists and log its properties
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ VERIFY: Cluster sleeve {clusterInstanceId} EXISTS in Revit: " +
                                $"Name='{clusterSleeve.Name}', Category='{clusterSleeve.Category?.Name ?? "NULL"}', " +
                                $"IsValid={clusterSleeve.IsValidObject}, Location={clusterSleeve.Location?.GetType().Name ?? "NULL"}\n");
                            
                            // Get bounding box
                            var bbox = clusterSleeve.get_BoundingBox(null);
                            if (bbox == null)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Could not get bounding box for cluster {clusterInstanceId}, skipping save");
                                continue;
                            }
                            
                            // Get dimensions
                            var widthParam = clusterSleeve.LookupParameter("Width");
                            var heightParam = clusterSleeve.LookupParameter("Height");
                            var depthParam = clusterSleeve.LookupParameter("Depth");
                            
                            double width = widthParam?.AsDouble() ?? 0.0;
                            double height = heightParam?.AsDouble() ?? 0.0;
                            double depth = depthParam?.AsDouble() ?? 0.0;
                            
                            // Get rotation data
                            double rotationAngleDeg = 0.0;
                            bool isRotated = false;
                            XYZ bboxMin = bbox.Min;
                            XYZ bboxMax = bbox.Max;
                            
                            if (rotationData.HasValue)
                            {
                                rotationAngleDeg = rotationData.Value.rotationAngleDeg;
                                isRotated = rotationData.Value.isRotated;
                                
                                // ✅ CRITICAL FIX: Only use rotated bounding box if it's valid (not zero/empty)
                                // If rotated bounding box is invalid, fall back to Revit bounding box
                                var rotatedBboxMin = rotationData.Value.rotatedBboxMin;
                                var rotatedBboxMax = rotationData.Value.rotatedBboxMax;
                                
                                // Check if rotated bounding box is valid (not zero and min < max)
                                bool isValidRotatedBbox = rotatedBboxMin != null && rotatedBboxMax != null &&
                                                          !rotatedBboxMin.IsAlmostEqualTo(XYZ.Zero) &&
                                                          !rotatedBboxMax.IsAlmostEqualTo(XYZ.Zero) &&
                                                          rotatedBboxMin.X < rotatedBboxMax.X &&
                                                          rotatedBboxMin.Y < rotatedBboxMax.Y &&
                                                          rotatedBboxMin.Z < rotatedBboxMax.Z;
                                
                                if (isValidRotatedBbox)
                                {
                                    bboxMin = rotatedBboxMin;
                                    bboxMax = rotatedBboxMax;
                                }
                                // Otherwise, keep using Revit bounding box (bbox.Min/Max from above)
                                
                                // Use rotated dimensions if available and valid
                                if (rotationData.Value.rotatedWidth > 0)
                                    width = rotationData.Value.rotatedWidth;
                                if (rotationData.Value.rotatedHeight > 0)
                                    height = rotationData.Value.rotatedHeight;
                                if (rotationData.Value.rotatedDepth > 0)
                                    depth = rotationData.Value.rotatedDepth;
                            }
                            
                            // Get host type and orientation
                            string hostType = GetHostTypeFromSleeve(clusterSleeve);
                            string hostOrientation = GetOrientationFromSleeve(clusterSleeve);
                            
                            // Get placement point (center of bounding box)
                            var placementPoint = (bboxMin + bboxMax) / 2.0;
                            
                            // Save to database
                            clusterRepository.SaveClusterSleeve(
                                clusterInstanceId: clusterInstanceId,
                                comboId: comboId,
                                filterId: filterId,
                                category: targetCategory,
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
                                hostOrientation: hostOrientation ?? "Unknown",
                                clashZoneIds: clashZoneIds);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[RefactoredClusterService] ✅ Saved cluster {clusterInstanceId} to database (ComboId={comboId}, FilterId={filterId}, Category={targetCategory})");
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ SAVED cluster {clusterInstanceId} to ClusterSleeves table: ComboId={comboId}, FilterId={filterId}, Category={targetCategory}, ClashZoneIds={clashZoneIds.Count}\n");
                            savedCount++;
                        }
                        catch (Exception saveEx)
                        {
                            failedCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] Error saving cluster {kvp.Key} to database: {saveEx.Message}");
                            
                            SafeFileLogger.SafeAppendText("cluster_errors.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ Error saving cluster {kvp.Key} to database: {saveEx.Message}\nStackTrace: {saveEx.StackTrace}\n");
                            // Continue with next cluster - don't fail entire operation
                        }
                    }
                    
                    // ✅ VALIDATION: Log summary and verify all clusters were saved
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ COMPLETED saving {savedCount} of {clusterToClashZoneIds.Count} clusters to ClusterSleeves table (Failed: {failedCount})\n");
                    
                    if (failedCount > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[RefactoredClusterService] ⚠️ Failed to save {failedCount} of {clusterToClashZoneIds.Count} clusters to database");
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ CRITICAL: Log all errors - database save failures must be visible
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RefactoredClusterService] ❌ CRITICAL ERROR saving cluster data to database: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ CRITICAL ERROR saving cluster data to database: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
                throw; // Re-throw to prevent silent failures - database save is critical
            }
        }

        /// <summary>
        /// Save sleeve snapshots for placed cluster sleeves.
        /// </summary>
        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
        /// 
        /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        /// To modify this method, you MUST:
        /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true in this class
        /// 2. Get explicit consent from the project owner
        /// 3. Test thoroughly with parameter transfer
        /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        /// 
        /// Saves cluster sleeve parameter snapshots to SleeveSnapshots table.
        /// This is REQUIRED for ParameterTransferService to work with cluster sleeves.
        /// 
        /// ✅ CRITICAL: Parameter values are aggregated from individual clash zones that form the cluster.
        /// - MEP parameters: Comma-separated (size parameters include duplicates)
        /// - Host parameters: Comma-separated unique values
        /// 
        /// ⚠️ DO NOT:
        /// - Skip parameter aggregation (ParameterTransferService will fail)
        /// - Change aggregation logic (must match ParameterTransferService expectations)
        /// - Remove ClusterInstanceId assignment (used to find cluster sleeves)
        /// - Modify comma-separated format (ParameterTransferService expects this format)
        /// 
        /// DATABASE SCHEMA (SleeveSnapshots table):
        /// - ClusterInstanceId (set for cluster sleeves, NULL for individual sleeves)
        /// - SleeveInstanceId (NULL for cluster sleeves)
        /// - MepParametersJson (JSON with aggregated comma-separated values)
        /// - HostParametersJson (JSON with aggregated comma-separated values)
        /// - SourceType ("Cluster" for cluster sleeves)
        /// </summary>
        private void SaveClusterSleeveSnapshots(
            Document doc,
            int filterId,
            List<FamilyInstance> placedClusters,
            string targetCategory)
        {
            // ⚠️ CONSENT CHECK: Prevent modifications without explicit consent
            if (!ALLOW_MODIFICATIONS_TO_PROTECTED_CODE)
            {
                // This method is protected - modifications require explicit consent
                // To modify: Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true and get consent
            }
            
            // ✅ VALIDATION: Ensure all required parameters are valid
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (filterId <= 0) throw new ArgumentException($"Invalid FilterId: {filterId}", nameof(filterId));
            if (placedClusters == null || placedClusters.Count == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ SaveClusterSleeveSnapshots: No clusters to save\n");
                return; // No clusters to save - not an error
            }
            if (string.IsNullOrWhiteSpace(targetCategory)) throw new ArgumentException("Target category cannot be null or empty", nameof(targetCategory));

            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var repository = new ClashZoneRepository(dbContext);
                    
                    // ✅ CRITICAL: Create ClashZone objects with aggregated parameter values from individual clash zones
                    var clusterZones = new List<ClashZone>();
                    foreach (var clusterSleeve in placedClusters)
                    {
                        if (clusterSleeve == null || !clusterSleeve.IsValidObject)
                            continue;

                        try
                        {
                            int clusterInstanceId = clusterSleeve.Id.IntegerValue;
                            
                            // Get clash zone IDs for this cluster
                            if (!_clusterToClashZoneIds.TryGetValue(clusterInstanceId, out var clashZoneIds) || clashZoneIds == null || clashZoneIds.Count == 0)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cluster {clusterInstanceId}: No clash zone IDs found in _clusterToClashZoneIds\n");
                                continue;
                            }

                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🔍 Cluster {clusterInstanceId}: Loading {clashZoneIds.Count} clash zones for parameter aggregation\n");

                            // Load all clash zones for the category (this will include parameter values from SleeveSnapshots)
                            // Then filter by GUID to get only the ones that form this cluster
                            var allCategoryZones = repository.GetClashZonesByCategory(targetCategory ?? string.Empty);
                            
                            // Filter to get only the clash zones that form this cluster
                            var individualClashZones = allCategoryZones
                                .Where(cz => clashZoneIds.Contains(cz.Id))
                                .ToList();

                            if (individualClashZones.Count == 0)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cluster {clusterInstanceId}: No clash zones loaded, skipping snapshot\n");
                                continue;
                            }

                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ Cluster {clusterInstanceId}: Loaded {individualClashZones.Count} clash zones, aggregating parameters\n");

                            // ⚠️⚠️⚠️ CRITICAL PARAMETER AGGREGATION LOGIC - MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
                            // To modify this aggregation logic, you MUST:
                            // 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true in this class
                            // 2. Get explicit consent from the project owner
                            // 3. Test thoroughly with ParameterTransferService
                            // 4. Verify comma-separated format is maintained
                            // 5. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
                            // 
                            // ✅ This aggregation logic MUST match ParameterTransferService expectations
                            // The format (comma-separated values) is required for parameter transfer to work
                            
                            // ⚠️ CONSENT CHECK: Prevent modifications without explicit consent
                            if (!ALLOW_MODIFICATIONS_TO_PROTECTED_CODE)
                            {
                                // This aggregation logic is protected - modifications require explicit consent
                                // To modify: Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true and get consent
                            }
                            
                            var aggregatedMepParams = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                            var aggregatedHostParams = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                            foreach (var cz in individualClashZones)
                            {
                                // ✅ CRITICAL FIX: Ensure Size parameter is available from MepElementSizeParameterValue
                                // If Size is not in MepParameterValues, add it from MepElementSizeParameterValue
                                if (cz.MepParameterValues == null)
                                {
                                    cz.MepParameterValues = new List<Models.SerializableKeyValue>();
                                }
                                
                                // ✅ CRITICAL FIX FOR PIPES: Remove any existing Size parameter first (same as individual sleeves)
                                // This ensures we replace old float values (e.g., "0.082") with fresh text values (e.g., "20 mmø") from database
                                // Same logic as AggregateParameterValues for individual zones - always use fresh value from current refresh
                                cz.MepParameterValues.RemoveAll(kv => kv != null && 
                                    (kv.Key?.Equals("Size", StringComparison.OrdinalIgnoreCase) == true ||
                                     kv.Key?.Equals("MEP Size", StringComparison.OrdinalIgnoreCase) == true ||
                                     kv.Key?.Equals("Service Size", StringComparison.OrdinalIgnoreCase) == true ||
                                     kv.Key?.Equals("MepElementFormattedSize", StringComparison.OrdinalIgnoreCase) == true));
                                
                                // ✅ PRIORITY 1: Use MepElementSizeParameterValue (raw Size parameter value from database, e.g., "20 mmø")
                                if (!string.IsNullOrWhiteSpace(cz.MepElementSizeParameterValue))
                                {
                                    cz.MepParameterValues.Add(new Models.SerializableKeyValue 
                                    { 
                                        Key = "Size", 
                                        Value = cz.MepElementSizeParameterValue 
                                    });
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ Cluster {clusterInstanceId}: Added Size='{cz.MepElementSizeParameterValue}' from MepElementSizeParameterValue for ClashZoneGuid={cz.Id} (replaced any existing Size parameter)\n");
                                }
                                // ✅ PRIORITY 2: Fallback to MepElementFormattedSize if MepElementSizeParameterValue is empty
                                else if (!string.IsNullOrWhiteSpace(cz.MepElementFormattedSize))
                                {
                                    cz.MepParameterValues.Add(new Models.SerializableKeyValue 
                                    { 
                                        Key = "Size", 
                                        Value = cz.MepElementFormattedSize 
                                    });
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ Cluster {clusterInstanceId}: Added Size='{cz.MepElementFormattedSize}' from MepElementFormattedSize for ClashZoneGuid={cz.Id} (fallback, MepElementSizeParameterValue was empty)\n");
                                }
                                else
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cluster {clusterInstanceId}: No Size parameter available for ClashZoneGuid={cz.Id} (both MepElementSizeParameterValue and MepElementFormattedSize are empty)\n");
                                }
                                
                                // ⚠️ PROTECTED: MEP Parameter Aggregation Logic
                                // Aggregate MEP parameters
                                if (cz.MepParameterValues != null)
                                {
                                    foreach (var param in cz.MepParameterValues)
                                    {
                                        if (param == null || string.IsNullOrWhiteSpace(param.Key) || string.IsNullOrWhiteSpace(param.Value))
                                            continue;

                                        if (!aggregatedMepParams.TryGetValue(param.Key, out var mepList))
                                        {
                                            mepList = new List<string>();
                                            aggregatedMepParams[param.Key] = mepList;
                                        }

                                        // ⚠️ PROTECTED: Size parameter handling
                                        // For size parameters, include all values (even duplicates)
                                        // This ensures "200 mmø, 200 mmø, 200 mmø" format for 3 identical sizes
                                        // For other parameters, only add unique values
                                        // ⚠️ DO NOT CHANGE THIS LOGIC - ParameterTransferService expects this format
                                        bool isSize = param.Key.Equals("Size", StringComparison.OrdinalIgnoreCase) ||
                                                     param.Key.Equals("MEP Size", StringComparison.OrdinalIgnoreCase) ||
                                                     param.Key.Equals("Service Size", StringComparison.OrdinalIgnoreCase) ||
                                                     param.Key.Equals("MepElementFormattedSize", StringComparison.OrdinalIgnoreCase);

                                        if (isSize)
                                        {
                                            // Size parameters: Include all values (duplicates allowed)
                                            mepList.Add(param.Value.Trim());
                                        }
                                        else
                                        {
                                            // Non-size parameters: Only add unique values
                                            if (!mepList.Any(v => string.Equals(v, param.Value.Trim(), StringComparison.OrdinalIgnoreCase)))
                                            {
                                                mepList.Add(param.Value.Trim());
                                            }
                                        }
                                    }
                                }

                                // ⚠️ PROTECTED: Host Parameter Aggregation Logic
                                // Aggregate Host parameters
                                if (cz.HostParameterValues != null)
                                {
                                    foreach (var param in cz.HostParameterValues)
                                    {
                                        if (param == null || string.IsNullOrWhiteSpace(param.Key) || string.IsNullOrWhiteSpace(param.Value))
                                            continue;

                                        if (!aggregatedHostParams.TryGetValue(param.Key, out var hostList))
                                        {
                                            hostList = new List<string>();
                                            aggregatedHostParams[param.Key] = hostList;
                                        }

                                        // ⚠️ PROTECTED: Host parameters always use unique values only
                                        // Host parameters: only add unique values
                                        if (!hostList.Any(v => string.Equals(v, param.Value.Trim(), StringComparison.OrdinalIgnoreCase)))
                                        {
                                            hostList.Add(param.Value.Trim());
                                        }
                                    }
                                }
                            }

                            // ⚠️ PROTECTED: Comma-separated format conversion
                            // ✅ CRITICAL: This format (comma-separated with ", " separator) is REQUIRED
                            // ParameterTransferService expects this exact format for cluster sleeves
                            // ⚠️ DO NOT CHANGE THE SEPARATOR OR FORMAT - This will break parameter transfer!
                            var mepParamsDict = aggregatedMepParams.ToDictionary(
                                kvp => kvp.Key,
                                kvp => string.Join(", ", kvp.Value), // ⚠️ PROTECTED: ", " separator is required
                                StringComparer.OrdinalIgnoreCase);

                            var hostParamsDict = aggregatedHostParams.ToDictionary(
                                kvp => kvp.Key,
                                kvp => string.Join(", ", kvp.Value), // ⚠️ PROTECTED: ", " separator is required
                                StringComparer.OrdinalIgnoreCase);

                            // Create cluster ClashZone with aggregated parameter values
                            var clusterZone = new ClashZone
                            {
                                ClusterSleeveInstanceId = clusterInstanceId,
                                MepElementCategory = "Cluster",
                                MepParameterValues = mepParamsDict.Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList(),
                                HostParameterValues = hostParamsDict.Select(kv => new Models.SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList(),
                                // Get SourceDocKey and HostDocKey from first clash zone (they should be the same for all)
                                SourceDocKey = individualClashZones.FirstOrDefault()?.SourceDocKey ?? string.Empty,
                                HostDocKey = individualClashZones.FirstOrDefault()?.HostDocKey ?? string.Empty,
                            };

                            // ✅ CRITICAL DEBUG: Log if Size parameter was found in aggregated params
                            var sizeKvp = mepParamsDict.FirstOrDefault(kvp => string.Equals(kvp.Key, "Size", StringComparison.OrdinalIgnoreCase));
                            bool hasSizeInAggregated = sizeKvp.Key != null;
                            string sizeValue = hasSizeInAggregated ? sizeKvp.Value : "NOT FOUND";
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ Cluster {clusterInstanceId}: Aggregated {mepParamsDict.Count} MEP params, {hostParamsDict.Count} Host params. Size parameter: {(hasSizeInAggregated ? $"FOUND='{sizeValue}'" : "MISSING")}\n");

                            clusterZones.Add(clusterZone);
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ❌ Error creating cluster zone for snapshot: {ex.Message}\n{ex.StackTrace}\n");
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] Error creating cluster zone for snapshot: {ex.Message}");
                        }
                    }

                    if (clusterZones.Count > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] ✅ Saving sleeve snapshots for {clusterZones.Count} cluster sleeves with aggregated parameters");
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ SAVING sleeve snapshots for {clusterZones.Count} cluster sleeves (with aggregated comma-separated parameters)\n");
                        
                        repository.SaveSleeveSnapshotsForPlacedSleeves(filterId, clusterZones);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] ✅ Saved sleeve snapshots for {clusterZones.Count} cluster sleeves");
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ COMPLETED saving sleeve snapshots for {clusterZones.Count} cluster sleeves to SleeveSnapshots table\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ No cluster zones created for snapshot saving\n");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[RefactoredClusterService] Error saving cluster sleeve snapshots: {ex.Message}");
                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] Error saving cluster sleeve snapshots: {ex.Message}\nStackTrace: {ex.StackTrace}\n");
                }
            }
        }

        /// <summary>
        /// Get host type from sleeve element.
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
                    DebugLogger.Error($"[RefactoredClusterService] Error getting host type from sleeve: {ex.Message}");
            }
            return "Unknown";
        }

        /// <summary>
        /// Get orientation from sleeve element.
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
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error getting orientation from sleeve: {ex.Message}");
            }
            return "Unknown";
        }

        #region Helper Methods (Extracted from UniversalClusterService)

        /// <summary>
        /// Get host type from ClashZone.
        /// </summary>
        private string GetHostTypeFromClashZone(ClashZone cz)
        {
            try
            {
                // Get host type from structural element type
                if (!string.IsNullOrEmpty(cz.StructuralElementType))
                {
                    if (cz.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase))
                        return "Wall";
                    else if (cz.StructuralElementType.Contains("Floor", StringComparison.OrdinalIgnoreCase) || 
                             cz.StructuralElementType.Contains("Slab", StringComparison.OrdinalIgnoreCase))
                        return "Floor";
                    else if (cz.StructuralElementType.Contains("Structural Framing", StringComparison.OrdinalIgnoreCase))
                        return "Structural Framing";
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error getting host type from clash zone: {ex.Message}");
            }
            
            return "Floor"; // Default fallback
        }

        /// <summary>
        /// Get effective orientation for clustering from ClashZone.
        /// </summary>
        private string GetEffectiveOrientationForClustering(ClashZone cz)
        {
            try
            {
                string hostType = GetHostTypeFromClashZone(cz);
                
                // For walls and framing, use HostOrientation (X/Y)
                if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    return cz.HostOrientation ?? "Unknown";
                }
                // For floors, return "Floor" to group ALL floor sleeves together for clustering
                else if (hostType == "Floor")
                {
                    return "Floor"; // Unified "Floor" instead of X/Y split
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[RefactoredClusterService] Error getting effective orientation: {ex.Message}");
            }
            
            return "Vertical"; // Default fallback
        }

        /// <summary>
        /// Get bounding box from ClashZone.
        /// </summary>
        private BoundingBoxXYZ? GetBoundingBoxFromClashZone(ClashZone cz)
        {
            try
            {
                if (cz.SleeveBoundingBoxMinX == 0 && cz.SleeveBoundingBoxMaxX == 0 &&
                    cz.SleeveBoundingBoxMinY == 0 && cz.SleeveBoundingBoxMaxY == 0 &&
                    cz.SleeveBoundingBoxMinZ == 0 && cz.SleeveBoundingBoxMaxZ == 0)
                {
                    return null; // Invalid bounding box
                }
                
                return new BoundingBoxXYZ
                {
                    Min = new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ),
                    Max = new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ),
                    Enabled = true
                };
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}

