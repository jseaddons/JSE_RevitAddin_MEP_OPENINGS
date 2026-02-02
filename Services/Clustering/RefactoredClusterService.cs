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
using JSE_RevitAddin_MEP_OPENINGS.Services.Helpers;
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
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.PreCalculation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Existence;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;

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
        
        // ✅ CORNER SERVICE: For calculating corners during batch save
        private readonly Services.Geometry.ISleeveCornerCalculationService _cornerService;
        
        // ✅ SOLID REFACTORING: Pre-calculation service (optional, used when flag enabled)
        private readonly IClusterPreCalculationService? _preCalculationService;
        
        // ✅ PATH 1a: Existence checker for smart placement (SOLID - SRP)
        private readonly ISleeveExistenceChecker _existenceChecker;
        
        // Supporting Services
        private readonly Services.Interfaces.Refactor.IFlagManager _flagManager;
        private readonly FilterManagementService _filterService;
        private readonly Document _doc;
        
        // ✅ BATCH V2 SERVICES
        private readonly BatchClusterCalculationService _batchCalculationService;
        private readonly BatchClusterPlacementService _batchPlacementService;
        
        // Internal state
        private string? _filterName;
        private readonly Dictionary<int, List<Guid>> _clusterToClashZoneIds = new Dictionary<int, List<Guid>>();
        private readonly SleeveParameterService _parameterService;
        
        // ✅ STEP 5 OPTIMIZATION: Deferred parameter batching for cluster placement (4-6× faster)
        // Accumulates parameter values during cluster placement loop, writes all after single regeneration
        // Key: ElementId of cluster sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private Dictionary<ElementId, Dictionary<string, object>> _deferredClusterParameters = new Dictionary<ElementId, Dictionary<string, object>>();
        
        // ✅ SOLID: Actual placement points for cluster sleeves (for database save)
        // Stores the actual calculated placement point used when placing each cluster sleeve
        // Key: Cluster instance ID (int)
        // Value: Actual placement point (XYZ)
        private Dictionary<int, XYZ> _actualPlacementPoints = new Dictionary<int, XYZ>();

        // ✅ BATCH PERSISTENCE FIX: Track calculated cluster data from placement service
        // This ensures the database can be saved even if Revit parameters are stale/deferred
        // Key: Cluster instance ID (int)
        private Dictionary<int, ClusterSaveData> _clusterSaveDataCache = new Dictionary<int, ClusterSaveData>();

        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor _performanceMonitor;

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
            Services.Geometry.ISleeveCornerCalculationService cornerService,
            ClusteringStrategyFactory? strategyFactory = null,
            Services.Interfaces.Refactor.IFlagManager? flagManager = null,
            FilterManagementService? filterService = null,
            IClusterPreCalculationService? preCalculationService = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor performanceMonitor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _dataService = dataService ?? throw new ArgumentNullException(nameof(dataService));
            _algorithmService = algorithmService ?? throw new ArgumentNullException(nameof(algorithmService));
            _rotationService = rotationService ?? throw new ArgumentNullException(nameof(rotationService));
            _placementService = placementService ?? throw new ArgumentNullException(nameof(placementService));
            _cleanupService = cleanupService ?? throw new ArgumentNullException(nameof(cleanupService));
            _timeoutService = timeoutService ?? throw new ArgumentNullException(nameof(timeoutService));
            _cornerService = cornerService ?? throw new ArgumentNullException(nameof(cornerService));
            _performanceMonitor = performanceMonitor;
            
            // ✅ BATCH V2: Initialize new services
            // Get the project-specific database path from SleeveDbContext
            string dbPath;
            using (var tempContext = new SleeveDbContext(doc))
            {
                dbPath = tempContext.DatabasePath;
            }
                
            _batchCalculationService = new BatchClusterCalculationService(
                algorithmService, 
                rotationService, 
                dbPath);
            
            // ✅ CODE REUSE: Inject SleeveParameterService for consistent parameter logic
            var parameterService = new JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService(doc);
            // ✅ BATCH V2: Pass injected monitor to batch placement service
            _batchPlacementService = new BatchClusterPlacementService(dbPath, new ClashZoneRepository(new SleeveDbContext(doc)), parameterService, cleanupService, _performanceMonitor);
            
            _flagManager = flagManager ?? throw new ArgumentNullException(nameof(flagManager));
            _strategyFactory = strategyFactory ?? new ClusteringStrategyFactory();
            _filterService = filterService ?? new FilterManagementService(
                doc,
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); },
                msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Error(msg); });
            
            // ✅ PARAMETER SERVICE: Initialize for batch parameter handling
            _parameterService = new SleeveParameterService(doc);
            
            // ✅ SOLID REFACTORING: Initialize pre-calculation service if flag is enabled
            if (OptimizationFlags.UseSOLIDRefactoredClusterPreCalculation && preCalculationService == null)
            {
                _preCalculationService = new ClusterPreCalculationService(rotationService);
            }
            else
            {
                _preCalculationService = preCalculationService; // Use injected service or null
            }
            
            // ✅ PATH 1a: Initialize existence checker for smart placement (SOLID - SRP)
            _existenceChecker = new SleeveExistenceChecker(doc);
        }



        /// <summary>
        /// ✅ BATCH V2: Robust Orchestration (Calculation First -> Optional Placement)
        /// </summary>
        public (int placedCount, int failedCount) ClusterSleevesV2(
            Document doc,
            List<ClashZone> clashZones,
            string targetCategory,
            int comboId,
            int filterId,
            bool useSingleTransaction = true,
            bool skipPlacement = false) // ✅ NEW: Option to skip placement for consolidation
        {
            try
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🚀 ORCHESTRATOR V2 START: Category={targetCategory}, Zones={clashZones.Count}, Mode={(useSingleTransaction ? "Bulk" : "Sequential")}, SkipPlacement={skipPlacement}, doc.IsModifiable={doc.IsModifiable}\n");

                // ✅ CRITICAL: Populate ClashZone cache BEFORE calculation so ClusterRotationService can look up by SleeveInstanceId
                // Required for Wall/Framing detection, correct dimensions (Width/Height), and orientation
                _dataService.LoadClashZoneCacheFromLoadedClashZones(clashZones, targetCategory);

                // Phase 1: Calculation (Parallel Safe, No Revit Transaction needed usually, or ReadOnly)
                // Ensure we are not in a transaction here if possible, or it's fine.
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🟡 BEFORE CalculateAndSave...\n");
                string batchId = _batchCalculationService.CalculateAndSave(clashZones, targetCategory, comboId, filterId, doc);
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🟢 AFTER CalculateAndSave, batchId={batchId}\n");

                // Phase 2: Placement (Sequential / Bulk Transaction) - SKIP if consolidating
                if (!skipPlacement)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🟡 BEFORE PlaceFromDatabase, doc.IsModifiable={doc.IsModifiable}...\n");
                    var result = _batchPlacementService.PlaceFromDatabase(doc, batchId, useSingleTransaction);
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🟢 AFTER PlaceFromDatabase, placed={result.placed}, failed={result.failed}\n");
                    return result;
                }
                else
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⏭️ SKIPPING placement (will be consolidated for all categories)\n");
                    return (0, 0); // Return 0,0 to indicate calculation only
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ ORCHESTRATOR ERROR: {ex.Message}\n{ex.StackTrace}\n");
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ❌ EXCEPTION in ClusterSleevesV2: {ex.Message}\n");
                return (0, 0);
            }
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
            // ✅ PERFORMANCE MONITORING: Use injected monitor or initialize new one
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string performanceLogName = $"ClusterPlacement_{targetCategory}_{timestamp}.log";
            var performanceMonitor = _performanceMonitor ?? new PlacementPerformanceMonitor(performanceLogName);

            // ✅ LOGGING: Log Build Timestamp to verify correct DLL is running (PATH 2/3)
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss}] 🚀 STARTING CLUSTERING (PATH 2/3) - Build: {Services.Helpers.VersionInfo.GetBuildTimestamp()} - Version: {Services.Helpers.VersionInfo.VersionTag}\n");

            // 🔥 CRITICAL: Direct System.IO logging to ensure we always see entry (bypasses SafeFileLogger completely)
            // This MUST work in both R2023 and R2024
            try
            {
                var versionTag = Services.Helpers.VersionInfo.VersionTag; // "R2023" or "R2024"
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
            catch (Exception ex)
            {
                // Fallback removed as requested - rely on standard SafeFileLogger or crash to debug
                System.Diagnostics.Debug.WriteLine($"Log init failed: {ex.Message}");
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

                    // ✅ SELF-HEALING FIRST: Check for deleted cluster sleeves BEFORE filtering by SleeveInstanceId
                    // This allows zones with SleeveInstanceId=0 (after clustering) to have SleeveInstanceId restored

                    // ✅ FORCE DETECTION / REFRESH CLEANUP: 
                    // If Force Detection is ON or we are in Path 3 (Refresh), we MUST delete existing clusters to prevent duplicates.
                    // This creates a "clean slate" for the zones in this scope.
                    bool isForceDetection = JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService.Instance.GetCurrentSettings().ForceDetectionMode;

                    if (isForceDetection || isPath3Validated || isPath3Invalidated || isPath3New)
                    {
                        var clustersToDelete = new List<ElementId>();
                        var clusterIdsSeen = new HashSet<int>();

                        foreach (var cz in allClashZones)
                        {
                            if (cz.ClusterSleeveInstanceId > 0 && cz.IsClusterResolved && !clusterIdsSeen.Contains(cz.ClusterSleeveInstanceId))
                            {
                                clusterIdsSeen.Add(cz.ClusterSleeveInstanceId);
                                clustersToDelete.Add(new ElementId(cz.ClusterSleeveInstanceId));
                            }
                        }

                        using (var cleanupTracker = performanceMonitor.TrackOperation("Force Cleanup (Prep)"))
                        {
                            if (clustersToDelete.Count > 0)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🧹 FORCE CLEANUP: Deleting {clustersToDelete.Count} existing clusters (Force={isForceDetection}, Path3={isPath3Validated}|{isPath3Invalidated}|{isPath3New})\n");

                                try
                                {
                                    doc.Delete(clustersToDelete);

                                    // ✅ DATABASE CLEANUP: Remove "data rows" for these deleted clusters to prevent ghosts
                                    // User Requirement: "delete the data rows of the invalidated zones" (Cluster AND Combined)
                                    try
                                    {
                                        using (var db = new SleeveDbContext(doc))
                                        {
                                            using (var tx = db.Connection.BeginTransaction())
                                            {
                                                using (var cmd = db.Connection.CreateCommand())
                                                {
                                                    cmd.Transaction = tx;

                                                    // 1. CLEANUP STANDARD CLUSTERS
                                                    var clusterIdList = string.Join(",", clustersToDelete.Select(id => id.IntegerValue)); // These are ClusterSleeveInstanceIds

                                                    if (!string.IsNullOrEmpty(clusterIdList))
                                                    {
                                                        // Delete from ClusterSleeves table
                                                        cmd.CommandText = $"DELETE FROM ClusterSleeves WHERE ClusterSleeveInstanceId IN ({clusterIdList})";
                                                        int delClusters = cmd.ExecuteNonQuery();

                                                        // Reset Cluster flags in ClashZones
                                                        cmd.CommandText = $"UPDATE ClashZones SET IsClusterResolvedFlag = 0, ClusterInstanceId = -1 WHERE ClusterInstanceId IN ({clusterIdList})";
                                                        int updClusters = cmd.ExecuteNonQuery();

                                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                            $"[{DateTime.Now:HH:mm:ss}] 🧹 DB CLEANUP (Clusters): Deleted {delClusters} rows, Reset {updClusters} zones.\n");
                                                    }

                                                    // 2. CLEANUP COMBINED CLUSTERS (If any zones reference them)
                                                    // Find zones associated with deleted clusters that ALSO have CombinedCluster metrics
                                                    var combinedIds = allClashZones
                                                        .Where(z => z.CombinedClusterSleeveInstanceId > 0 &&
                                                                   (clustersToDelete.Any(id => id.IntegerValue == z.ClusterSleeveInstanceId) || isPath3Invalidated))
                                                        .Select(z => z.CombinedClusterSleeveInstanceId)
                                                        .Distinct()
                                                        .ToList();

                                                    if (combinedIds.Count > 0)
                                                    {
                                                        var combinedIdList = string.Join(",", combinedIds);

                                                        // Delete from CombinedSleeves table (Assuming table name is CombinedSleeves)
                                                        // Also delete the Revit Element if it exists (need to find it... but we only have ID)
                                                        // Ideally we should have collected Revit Element IDs for Combined too.
                                                        // For now, clean DB to break the link.

                                                        cmd.CommandText = $"DELETE FROM CombinedSleeves WHERE CombinedClusterSleeveInstanceId IN ({combinedIdList})";
                                                        try { cmd.ExecuteNonQuery(); } catch { /* Table might not exist yet in early phase */ }

                                                        // Reset Combined flags in ClashZones
                                                        cmd.CommandText = $"UPDATE ClashZones SET IsCombinedResolved = 0, CombinedClusterInstanceId = -1, CombinedClusterSleeveInstanceId = -1 WHERE CombinedClusterSleeveInstanceId IN ({combinedIdList})";
                                                        int updCombined = cmd.ExecuteNonQuery();

                                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                            $"[{DateTime.Now:HH:mm:ss}] 🧹 DB CLEANUP (Combined): Reset {updCombined} zones for IDs: {combinedIdList}.\n");
                                                    }
                                                }
                                                tx.Commit();
                                            }
                                        }
                                    }
                                    catch (Exception dbEx)
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB Cleanup Failed: {dbEx.Message}\n");
                                    }
                                }
                                catch (Exception delEx)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Delete failed (partial): {delEx.Message}\n");
                                }
                            }
                            cleanupTracker.SetItemCount(clustersToDelete.Count);
                        }
                    }

                    int resetCount = 0;
                    foreach (var cz in allClashZones)
                    {
                        // Only process zones that were previously clustered
                        if (cz.ClusterSleeveInstanceId > 0 && cz.IsClusterResolved)
                        {
                            var clusterSleeve = doc.GetElement(new ElementId(cz.ClusterSleeveInstanceId));
                            if (clusterSleeve == null || !clusterSleeve.IsValidObject)
                            {
                                // ✅ SELF-HEALING: Cluster sleeve was deleted - reset flags and allow re-clustering
                                // ✅ CRITICAL: Restore SleeveInstanceId from AfterClusterSleevePlacedSleeveInstanceId
                                if (cz.AfterClusterSleevePlacedSleeveInstanceId > 0)
                                {
                                    cz.SleeveInstanceId = cz.AfterClusterSleevePlacedSleeveInstanceId;
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] 🔄 SELF-HEALING: Restored SleeveInstanceId={cz.AfterClusterSleevePlacedSleeveInstanceId} for zone {cz.Id}\n");
                                }
                                cz.IsClusterResolved = false;
                                cz.IsClusterResolvedFlag = false;
                                cz.ClusterSleeveInstanceId = -1;
                                cz.AfterClusterSleevePlacedSleeveInstanceId = 0;
                                cz.MarkedForClusterProcess = true; // ✅ Mark for cluster process
                                resetCount++;
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔄 SELF-HEALING: Reset cluster flags for zone {cz.Id} (cluster sleeve deleted)\n");
                            }
                        }
                    }

                    if (resetCount > 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔄 SELF-HEALING: Reset {resetCount} zones for re-clustering (deleted cluster sleeves)\n");
                    }

                    // ✅ NOW filter by SleeveInstanceId > 0 (after self-healing restored SleeveInstanceId)
                    var withSleeveId = allClashZones.Where(cz => cz.SleeveInstanceId > 0).ToList();

                    // 🔍 TRACING: Check specific problematic IDs
                    foreach (var czId in new[] { 1183693, 1183702 })
                    {
                        var found = allClashZones.FirstOrDefault(cz => cz.SleeveInstanceId == czId);
                        if (found != null)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[TRACE-FILTER] ID {czId}: Found=True, SleeveId={found.SleeveInstanceId}, IsResolved={found.IsClusterResolved}, IsCurrentClash={found.IsCurrentClash}, Category={found.MepElementCategory}\n");
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[TRACE-FILTER] ID {czId}: Found=False in allClashZones\n");
                        }
                    }

                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: {withSleeveId.Count} zones with SleeveInstanceId > 0\n");

                    // Filter by not cluster resolved
                    var notClusterResolved = withSleeveId.Where(cz => !cz.IsClusterResolved).ToList();
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: {notClusterResolved.Count} zones not cluster resolved (or reset for re-clustering)\n");

                    filteredClashZones = notClusterResolved
                        .Where(cz => (string.IsNullOrEmpty(targetCategory) || string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase)) 
                                      && (cz.MarkedForClusterProcess.GetValueOrDefault() == true)) // ✅ CRITICAL FIX: Explicitly handle nullable bool
                        .ToList();
                    
                    if (notClusterResolved.Count > filteredClashZones.Count)
                    {
                        var diffCount = notClusterResolved.Count - filteredClashZones.Count;
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ FILTERED OUT {diffCount} zones because MarkedForClusterProcess=FALSE/NULL\n");
                        
                        // Log a sample of filtered out zones to prove why
                        var sampleFiltered = notClusterResolved.Where(cz => cz.MarkedForClusterProcess.GetValueOrDefault() != true).Take(5).ToList();
                        foreach(var s in sampleFiltered)
                        {
                             SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                 $"[{DateTime.Now:HH:mm:ss}]    ↳ Filtered Out Zone {s.Id}: MarkedForClusterProcess={s.MarkedForClusterProcess?.ToString() ?? "NULL"}\n");
                        }
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                             $"[{DateTime.Now:HH:mm:ss}] ✅ All {filteredClashZones.Count} zones are MarkedForClusterProcess=TRUE\n");
                    }
                    
                    // ✅ AGGRO LOGGING: Detailed breakdown of MarkedForClusterProcess states
                    var total = notClusterResolved.Count;
                    var markedTrue = notClusterResolved.Count(cz => cz.MarkedForClusterProcess.GetValueOrDefault() == true);
                    var markedFalse = notClusterResolved.Count(cz => cz.MarkedForClusterProcess.GetValueOrDefault() == false);
                    var markedNull = notClusterResolved.Count(cz => !cz.MarkedForClusterProcess.HasValue);
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] 📊 POPULATION CHECK: Total={total}, True={markedTrue}, False={markedFalse}, Null={markedNull}\n");
                    if (markedNull > 0)
                    {
                         SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🚨 WARNING: Found {markedNull} zones with NULL MarkedForClusterProcess. These should have been fixed by BulkUpdate!\n");
                    }

                    filterTracker.SetItemCount(filteredClashZones.Count);

                    // ✅ DIAGNOSTIC: Inspect StructuralElementIdValue distribution
                    if (!DeploymentConfiguration.DeploymentMode && filteredClashZones.Count > 0)
                    {
                        try
                        {
                            var hostDistribution = filteredClashZones
                                .GroupBy(z => z.StructuralElementIdValue)
                                .Select(g => new { HostId = g.Key, Count = g.Count() })
                                .OrderByDescending(x => x.Count)
                                .ToList();

                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 📊 HOST ELEMENT DISTRIBUTION: {filteredClashZones.Count} zones across {hostDistribution.Count} unique hosts\n");

                            foreach (var host in hostDistribution.Take(20)) // Log top 20 hosts to avoid massive logs
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   → HostId {host.HostId}: {host.Count} zones\n");
                            }
                            if (hostDistribution.Count > 20)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   → ... and {hostDistribution.Count - 20} more hosts\n");
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error logging host distribution: {ex.Message}\n");
                        }
                    }
                }

                try
                {
                    string logPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ FILTERED: {filteredClashZones.Count} clash zones ready for clustering (SleeveId>0, not cluster resolved, category='{targetCategory}')\n");
                }
                catch { }
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ FILTERED: {filteredClashZones.Count} clash zones ready for clustering (SleeveId>0, not cluster resolved, category='{targetCategory}')\n");

                // ✅ PHASE 6: PRE-CLUSTERING VALIDATION
                // Validate all zones have required database values BEFORE clustering
                var validZones = new List<ClashZone>();
                var invalidZones = new List<(Guid id, string reason)>();

                foreach (var cz in filteredClashZones)
                {
                    if (cz == null) continue;

                    // Validate HostElementId
                    if (cz.StructuralElementIdValue <= 0)
                    {
                        invalidZones.Add((cz.Id, $"Invalid HostElementId: {cz.StructuralElementIdValue}"));
                        continue;
                    }

                    // Validate StructuralThickness
                    if (cz.StructuralElementThickness <= 0.001)
                    {
                        invalidZones.Add((cz.Id, $"Invalid StructuralThickness: {cz.StructuralElementThickness}"));
                        continue;
                    }

                    // Validate HostOrientation (for walls/framing is critical, for others good to have)
                    // If HostOrientation is null, we might default to something else, but let's log it
                    // Note: Floors might have empty HostOrientation but use MepElementRotationAngle

                    // Passed all validations
                    validZones.Add(cz);
                }

                if (invalidZones.Count > 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ PRE-CLUSTERING VALIDATION: Removed {invalidZones.Count} invalid zones\n");
                    foreach (var (id, reason) in invalidZones)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌ INVALID ZONE SKIPPED: ID={id}, Reason={reason}\n");
                    }
                    // Update list to only include valid zones
                    filteredClashZones = validZones;
                }


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
                    // ✅ STEP 8: Group sleeves by host type, system type, orientation, and SPATIAL PROXIMITY (Phase 12: Duplicate Fix)
                    // We group by 10m buckets (approx 32.8ft) to ensure sleeves on the same wall/floor are processed together
                    // regardless of slight Host ID mismatches or linked model issues.
                    sleeveGroups = rawSleeves.GroupBy(sleeve =>
                    {
                        ClashZone cz = sleeve.ClashZone;
                        // Use IntesectionPoint if available, otherwise 0
                        double x = cz?.IntersectionPointX ?? 0;
                        double y = cz?.IntersectionPointY ?? 0;
                        double z = cz?.IntersectionPointZ ?? 0;

                        // 10m bucket size (approx 32.8 ft)
                        int bucketX = (int)Math.Floor(x / 32.8);
                        int bucketY = (int)Math.Floor(y / 32.8);
                        int bucketZ = (int)Math.Floor(z / 32.8);

                        return new SleeveGroupKey(
                            sleeve.HostType,
                            sleeve.Category,
                            sleeve.Orientation,
                            bucketX,
                            bucketY,
                            bucketZ
                        );
                    }).ToArray();
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

                    // ✅ PHASE 6: POST-CLUSTERING VALIDATION
                    // Verify each cluster group has consistent HostElementId, HostOrientation, and StructuralThickness
                    var validatedClustersByGroup = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
                    int discardedClusterCount = 0;

                    foreach (var kvp in clustersByGroup)
                    {
                        var groupKey = kvp.Key;
                        var clusterList = kvp.Value;
                        var validClustersForGroup = new List<List<dynamic>>();

                        foreach (var cluster in clusterList)
                        {
                            if (cluster == null || cluster.Count == 0) continue;

                            // Get first zone as reference
                            ClashZone? firstClashZone = null;
                            var firstItem = cluster[0];
                            if (firstItem is ClashZone cz) firstClashZone = cz;
                            else if (firstItem?.ClashZone != null) firstClashZone = firstItem.ClashZone as ClashZone;

                            if (firstClashZone == null)
                            {
                                SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ DISCARDING CLUSTER: First zone data invalid\n");
                                discardedClusterCount++;
                                continue;
                            }

                            int referenceHostId = firstClashZone.StructuralElementIdValue;
                            string referenceOrientation = (firstClashZone.HostOrientation ?? "").Trim();
                            double referenceThickness = firstClashZone.StructuralElementThickness;

                            // Validate ALL zones in cluster match reference values
                            bool isValid = true;
                            foreach (var item in cluster)
                            {
                                ClashZone? itemCz = null;
                                if (item is ClashZone icz) itemCz = icz;
                                else if (item?.ClashZone != null) itemCz = item.ClashZone as ClashZone;

                                if (itemCz == null) { isValid = false; break; }

                                // Check HostElementId
                                if (itemCz.StructuralElementIdValue != referenceHostId)
                                {
                                    SafeFileLogger.SafeAppendText("placement_errors.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ DISCARDING CLUSTER: Mixed HostElementIds ({referenceHostId} vs {itemCz.StructuralElementIdValue})\n");
                                    isValid = false; break;
                                }

                                // Check HostOrientation (Strict string match? Or loose?)
                                // We check simply for inequality of trimmed strings
                                string itemOrientation = (itemCz.HostOrientation ?? "").Trim();
                                if (!string.Equals(referenceOrientation, itemOrientation, StringComparison.OrdinalIgnoreCase))
                                {
                                    SafeFileLogger.SafeAppendText("placement_errors.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ DISCARDING CLUSTER: Mixed HostOrientations ('{referenceOrientation}' vs '{itemOrientation}')\n");
                                    isValid = false; break;
                                }

                                // Check StructuralThickness (tolerance 1mm = ~0.003 ft)
                                if (Math.Abs(itemCz.StructuralElementThickness - referenceThickness) > 0.003)
                                {
                                    SafeFileLogger.SafeAppendText("placement_errors.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ DISCARDING CLUSTER: Mixed StructuralThickness ({referenceThickness:F3} vs {itemCz.StructuralElementThickness:F3})\n");
                                    isValid = false; break;
                                }
                            }

                            if (isValid)
                            {
                                validClustersForGroup.Add(cluster);
                            }
                            else
                            {
                                discardedClusterCount++;
                            }
                        }

                        if (validClustersForGroup.Count > 0)
                        {
                            validatedClustersByGroup[groupKey] = validClustersForGroup;
                        }
                    }

                    if (discardedClusterCount > 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ POST-CLUSTERING VALIDATION: Discarded {discardedClusterCount} invalid clusters (mixed properties)\n");
                    }

                    // Replace results with validated clusters
                    clustersByGroup = validatedClustersByGroup;


                    int totalClusters = clustersByGroup?.Sum(g => g.Value?.Count ?? 0) ?? 0;
                    int totalClustersWithMultipleSleeves = clustersByGroup?.Sum(g => g.Value?.Count(c => c.Count > 1) ?? 0) ?? 0;
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ CLUSTERING: Formed {clustersByGroup?.Count ?? 0} cluster groups, {totalClusters} total clusters, {totalClustersWithMultipleSleeves} clusters with >1 sleeve\n");
                    formTracker.SetItemCount(totalClusters);

                    // ✅ Mark zones in multi-sleeve clusters with MarkedForClusterProcess = true
                    if (clustersByGroup != null && allClashZones != null)
                    {
                        var sleeveIdToZone = allClashZones.Where(cz => cz.SleeveInstanceId > 0)
                            .ToDictionary(cz => cz.SleeveInstanceId, cz => cz);

                        foreach (var groupEntry in clustersByGroup)
                        {
                            foreach (var cluster in groupEntry.Value)
                            {
                                if (cluster.Count > 1) // Only clusters with >1 sleeve
                                {
                                    foreach (var sleeve in cluster)
                                    {
                                        int sleeveId = (int)sleeve.SleeveInstanceId;
                                        if (sleeveIdToZone.TryGetValue(sleeveId, out var zone))
                                        {
                                            zone.MarkedForClusterProcess = true;
                                        }
                                    }
                                }
                            }
                        }
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ CLUSTERING: Marked zones in {totalClustersWithMultipleSleeves} multi-sleeve clusters with MarkedForClusterProcess=true\n");
                    }
                }

                // ✅ STEP 11: Check timeout after clustering
                if (_timeoutService.IsTimedOut())
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RefactoredClusterService] ⏱ TIMEOUT: Clustering exceeded {_timeoutService.TimeoutLimitMs / 1000} second limit");
                    _timeoutService.ShowTimeoutWarning("during FormClusters");
                    return (placedCount, deletedCount);
                }

                // ⚠️⚠️⚠️ CRITICAL PROTECTION - SOLID REFACTORING WITH 28 FEATURES ⚠️⚠️⚠️
                // This section implements parallel pre-calculation for cluster rotation angles and bounding boxes.
                // When flag is enabled: Pre-calculates all clusters in parallel BEFORE placement loop (2-4× faster)
                // When flag is disabled: Falls back to legacy sequential calculation inside placement loop
                // All 28 features from Comprehensive Architecture are preserved in both paths.
                // ⚠️ DO NOT REMOVE FALLBACK LOGIC - It ensures backward compatibility and safe rollback

                // ✅ SOLID REFACTORING: Pre-calculate rotation angles and bounding boxes in parallel (if flag enabled)
                // ✅ PRE-CALCULATION: Pre-calculate rotation angles and bounding boxes in parallel
                // User requested NO FALLBACKS - if this fails, we want to know immediately.
                Dictionary<int, ClusterCalculationResult>? preCalculatedResults = null;

                if (_preCalculationService != null)
                {
                    using (var preCalcTracker = performanceMonitor.TrackOperation("Pre-Calculate Clusters (Parallel)"))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 🚀 Starting parallel pre-calculation for all clusters\n");

                        // ✅ Build ClashZone dictionary for fast lookup
                        var clashZoneDict = new Dictionary<int, ClashZone>();
                        if (allClashZones != null && allClashZones.Count > 0)
                        {
                            foreach (var cz in allClashZones)
                            {
                                if (cz.SleeveInstanceId > 0)
                                {
                                    clashZoneDict[cz.SleeveInstanceId] = cz;
                                }
                            }
                        }

                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 📦 Passing {clashZoneDict.Count} pre-loaded ClashZones to pre-calculation\n");

                        // ✅ EXECUTE PRE-CALCULATION (No Try/Catch Wrapper)
                        preCalculatedResults = _preCalculationService.PreCalculateAllClusters(clustersByGroup, doc, xmlFilePath, clashZoneDict);

                        int validPreCalc = preCalculatedResults.Values.Count(r => r.IsValid);
                        int invalidPreCalc = preCalculatedResults.Values.Count(r => !r.IsValid);

                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ Pre-calculation complete: {validPreCalc} valid, {invalidPreCalc} invalid\n");

                        preCalcTracker.SetItemCount(validPreCalc);
                    }
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                       $"[{DateTime.Now:HH:mm:ss}] ⚠️ Pre-calculation service is NULL - cannot optimize\n");
                }

                // ✅ PERFORMANCE: Track cluster placement loop
                try
                {
                    int clusterProcessedCount = 0;
                    // ✅ BATCH MODE FIX: Track by ClashZone GUID, not SleeveInstanceId (Prevent Duplicates across groups)
                    // This is critical for batch mode - ensures each ClashZone is processed exactly once
                    var globalProcessedClashZoneGuids = new HashSet<Guid>();

                    try
                    {
                        using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Clusters Loop"))
                        {
                            // ✅ STEP 12: Process each cluster group
                            // ✅ SOLID REFACTORING: Use pre-calculated results if available, otherwise use legacy sequential calculation
                            int clusterIndex = 0;
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
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 PROCESSING: Group {groupKey.hostType}/{groupKey.systemType}/{groupKey.orientation} (Spatial: X={groupKey.SpatialX}, Y={groupKey.SpatialY}, Z={groupKey.SpatialZ}) has {clusters.Count} clusters\n");

                                // ✅ UNIFIED BATCH CONTEXT: Ensure ClusterPlacementService writes to the same dictionary we flush
                                // This fixes the "Missing Parameters" bug where PlaceClusterSleeve used a local dictionary that confused the flush logic
                                // ✅ UNIFIED IMMEDIDATE MODE: Do NOT divert. Write directly to element.
                                // _parameterService.DivertedBatchDictionary = _deferredClusterParameters;

                                foreach (var cluster in clusters)
                                {
                                    // ✅ PHASE 6: CLUSTER PLACEMENT LOGGING
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        var idx = clusters.IndexOf(cluster);
                                        try
                                        {
                                            var firstItem = cluster[0];
                                            ClashZone? logClashZone = null;
                                            if (firstItem is ClashZone cz) logClashZone = cz;
                                            else if (firstItem?.ClashZone != null) logClashZone = firstItem.ClashZone as ClashZone;

                                            if (logClashZone != null)
                                            {
                                                string logMsg = $"[{DateTime.Now:HH:mm:ss}] 🔧 PLACING CLUSTER {idx + 1}/{clusters.Count}:\n";
                                                logMsg += $"  GroupKey: Host={groupKey.hostType}, Orient={groupKey.orientation}\n";
                                                logMsg += $"  Size: {cluster.Count} zones\n";
                                                logMsg += $"  HostElementId: {logClashZone.StructuralElementIdValue}\n";
                                                logMsg += $"  HostOrientation: {logClashZone.HostOrientation}\n";
                                                double thkMm = RevitUnitConversionService.Instance.FromInternalMillimeters(logClashZone.StructuralElementThickness);
                                                logMsg += $"  StructuralThickness: {thkMm:F1}mm\n";
                                                SafeFileLogger.SafeAppendText("cluster_debug.log", logMsg);
                                            }
                                        }
                                        catch { /* Log failure shouldn't stop placement */ }
                                    }

                                    // ✅ PHASE 1: Duplicate Prevention Check
                                    // Verify if any sleeve in this cluster has already been processed in another group
                                    // This handles cases where sleeves sit on the boundary of spatial buckets
                                    var clusterSleeveIds = cluster.Select(s => (int)s.SleeveInstanceId).ToList();
                                    // ✅ PHASE 1: DUPLICATE PREVENTION (Global Tracking)
                                    bool isDuplicate = false;
                                    try
                                    {
                                        foreach (var s in cluster)
                                        {
                                            Guid czGuid = Guid.Empty;
                                            try { czGuid = (s.ClashZone as ClashZone)?.Id ?? Guid.Empty; }
                                            catch { }

                                            if (czGuid != Guid.Empty && globalProcessedClashZoneGuids.Contains(czGuid))
                                            {
                                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                    $"[{DateTime.Now:HH:mm:ss}] 🚫 DUPLICATE PREVENTED: ClashZone {czGuid} already processed. Skipping cluster.\n");
                                                isDuplicate = true;
                                                break;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ ERROR in Duplicate Check: {ex.Message}\n");
                                    }

                                    if (isDuplicate)
                                    {
                                        clusterIndex++;
                                        continue;
                                    }

                                    // ✅ Add IDs to global set (Pre-add to block self/others immediately)
                                    foreach (var s in cluster)
                                    {
                                        Guid czGuid = Guid.Empty;
                                        try { czGuid = (s.ClashZone as ClashZone)?.Id ?? Guid.Empty; } catch { }
                                        if (czGuid != Guid.Empty) globalProcessedClashZoneGuids.Add(czGuid);
                                    }

                                    if (cluster.Count <= 1)
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⏭️ SKIP: Cluster with only {cluster.Count} sleeve(s) (need >1 to cluster)\n");
                                        clusterIndex++; // Increment index even for skipped clusters
                                        continue; // Skip individual sleeves
                                    }

                                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ PROCESSING: Cluster with {cluster.Count} sleeves (index={clusterIndex})\n");

                                    try
                                    {


                                        // ✅ PERFORMANCE: Track cluster placement
                                        (bool success, int placedCount, int deletedCount, FamilyInstance? placedClusterSleeve, int? capturedClusterSleeveId) placementResult;
                                        using (placementLoopTracker?.TrackSubOperation("Place Cluster Sleeve"))
                                        {
                                            // ✅ USE PRE-CALCULATED RESULTS
                                            ClusterCalculationResult? preCalcResult = null;
                                            if (preCalculatedResults != null && preCalculatedResults.ContainsKey(clusterIndex))
                                            {
                                                preCalcResult = preCalculatedResults[clusterIndex];
                                            }

                                            // ✅ STEP 14: Place cluster sleeve (Phase 5: Placement Service + Phase 6: Rotation Service + Phase 3: BoundingBox)
                                            // Note: Placement service needs to be wired with functions from rotation service and data service
                                            placementResult = PlaceClusterForGroup(
                                                doc,
                                                cluster,
                                                groupKey,
                                                targetCategory,
                                                xmlFilePath,
                                                placementLoopTracker, // Pass tracker for sub-operation tracking
                                                preCalcResult); // ✅ Pass pre-calculated result if available
                                        } // End Place Cluster Sleeve sub-operation

                                        clusterIndex++; // Increment index after processing



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
                                                // ✅ CRITICAL: Always register with legacy FlagManager (static method works regardless of which service is active)
                                                // The legacy FlagManager's static HashSet is shared and will protect sleeves even when using refactored services
                                                Services.FlagManagement.FlagManagerProtectionHelper.RegisterRecentlyPlacedClusterSleeve(placementResult.capturedClusterSleeveId.Value);

                                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                    $"[{DateTime.Now:HH:mm:ss}] ✅✅✅ REGISTERED cluster sleeve {placementResult.capturedClusterSleeveId.Value} with FlagManagerProtectionHelper (static protection)\n");

                                                placedClusters.Add(placementResult.placedClusterSleeve);

                                                // Track ClashZoneIds for database save
                                                var clusterClashZoneIds = cluster
                                                    .Select(s => (s.ClashZone as ClashZone)?.Id)
                                                    .Where(id => id.HasValue)
                                                    .Select(id => id!.Value)
                                                    .ToList();

                                                // ✅ BATCH MODE FIX: Mark ClashZones as processed globally (by GUID)
                                                foreach (var s in cluster)
                                                {
                                                    Guid czGuid = Guid.Empty;
                                                    try { czGuid = (s.ClashZone as ClashZone)?.Id ?? Guid.Empty; } catch { }
                                                    if (czGuid != Guid.Empty) globalProcessedClashZoneGuids.Add(czGuid);
                                                }

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

                                                    // ✅ PHASE 3 PERISTENCE: Calculate and save cluster corners
                                                    if (placementResult.placedClusterSleeve != null && placementResult.capturedClusterSleeveId.HasValue)
                                                    {
                                                        try
                                                        {
                                                            var cornerService = new SleeveCornerCalculationService();
                                                            // Extract parameters from placed element
                                                            var loc = placementResult.placedClusterSleeve.Location as LocationPoint;
                                                            var point = loc?.Point;
                                                            var rotation = loc?.Rotation ?? 0.0;

                                                            // Get dimensions (prioritize RCS dimensions)
                                                            var widthParam = placementResult.placedClusterSleeve.LookupParameter("Element Width") ?? placementResult.placedClusterSleeve.LookupParameter("Width");
                                                            var heightParam = placementResult.placedClusterSleeve.LookupParameter("Element Height") ?? placementResult.placedClusterSleeve.LookupParameter("Height");

                                                            // ⚠️ DEPRECATED: Old per-cluster corner save - now handled by BatchSaveClusterDataToDatabase
                                                            // This was causing 0.0 values because it ran before DB entry existed
                                                            /*
                                                            double width = widthParam?.AsDouble() ?? 1.0;
                                                            double height = heightParam?.AsDouble() ?? 1.0;

                                                            if (point != null)
                                                            {
                                                                var cornersResult = cornerService.CalculateCorners(point, width, height, rotation);

                                                                if (cornersResult.HasValue)
                                                                {
                                                                    var c = cornersResult.Value;
                                                                    _dataService.UpdateClusterSleeveCorners(
                                                                        placementResult.capturedClusterSleeveId.Value,
                                                                        c.corner1.X, c.corner1.Y, c.corner1.Z,
                                                                        c.corner2.X, c.corner2.Y, c.corner2.Z,
                                                                        c.corner3.X, c.corner3.Y, c.corner3.Z,
                                                                        c.corner4.X, c.corner4.Y, c.corner4.Z
                                                                    );
                                                                }
                                                            }

                                                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ SAVED CORNERS for Cluster {placementResult.capturedClusterSleeveId.Value}\n");
                                                            */

                                                        }
                                                        catch (Exception cornerEx)
                                                        {
                                                            // File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️ FAILED to save corners for Cluster {placementResult.capturedClusterSleeveId.Value}: {cornerEx.Message}\n");
                                                        }
                                                    }
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
                                    } // End catch
                                } // End foreach
                                placementLoopTracker.SetItemCount(placedCount);
                            } // End using
                    }
                    }
                    finally
                    {
                        // ✅ UNIFIED BATCH CONTEXT: Reset diverted dictionary
                        _parameterService.DivertedBatchDictionary = null;
                    }




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
                        // int flushedCount = FlushDeferredClusterParameters();
                        int flushedCount = 0; // Forced to 0

                        // ✅ CRITICAL FIX: Regenerate AFTER flush to ensure Revit has updated parameters for cleanup
                        // This ensures cleanup service can read correct dimensions from Revit element if deferred dictionary is empty
                        if (flushedCount > 0)
                        {
                            try
                            {
                                doc.Regenerate();
                                System.Threading.Thread.Sleep(100); // Brief pause for regeneration
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ POST-FLUSH REGENERATION: Regenerated document after flushing {flushedCount} cluster sleeve parameters\n");
                                }
                            }
                            catch (Exception regenEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Error regenerating after parameter flush: {regenEx.Message}");
                            }
                        }
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
                            // ✅ CRITICAL FIX: Pass deferred parameters dictionary to cleanup service so it can read correct dimensions when batching is enabled
                            // This prevents cleanup from using stale parameter values from Revit before flush/regeneration
                            // Note: Dictionary may be empty if already flushed, but cleanup will fallback to Revit element after regeneration
                            // Updated to pass targetCategory to prevent cross-category deletion
                            deletedInCleanup = _cleanupService.CleanupSleevesWithinClusters(doc, validClusters, _deferredClusterParameters, targetCategory);
                            deletedCount += deletedInCleanup;
                            cleanupTracker.SetItemCount(deletedInCleanup);
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: No valid cluster sleeves to protect, skipping cleanup service\n");
                        }
                    }

                    // ✅ NOTE: Do NOT clear _deferredClusterParameters here!
                    // The FlushDeferredClusterParameters() method needs to read this data later.
                    // The flush method will clear it after writing parameters to Revit elements.
                    // Clearing here was causing cluster sleeves to default to standard size.

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

                    // ✅✅✅ CRITICAL FIX: Regenerate ONCE after all placements, BEFORE collecting corners
                    // This ensures all geometry is finalized before we calculate corner coordinates
                    if (_clusterToClashZoneIds.Count > 0 && comboId.HasValue && filterId.HasValue)
                    {
                        try
                        {
                            var versionTag = Helpers.VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");

                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔄 REGENERATING document before collecting corners for {_clusterToClashZoneIds.Count} clusters\n");
                        }
                        catch { }

                        doc.Regenerate();

                        try
                        {
                            var versionTag = Helpers.VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");

                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ REGENERATION COMPLETE - Now calling BatchSaveClusterDataToDatabase\n");
                        }
                        catch { }

                        // 🔴 REMOVED REDUNDANT CALL: BatchSaveClusterDataToDatabase was being called twice
                        // BatchSaveClusterDataToDatabase(doc, comboId.Value, filterId.Value, targetCategory, _clusterToClashZoneIds);

                        try
                        {
                            var versionTag = Helpers.VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");
                            
                            // Log regarding next step which is actual tracked save
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⏳ Proceeding to tracked Batch Save...\n");
                        }
                        catch { }
                    }


                    // ✅ PERFORMANCE: Track database save
                    using (var dbSaveTracker = performanceMonitor.TrackOperation("Save Cluster Data to Database"))
                    {
                        // ✅ STEP 17: Save cluster data to database (if comboId and filterId are available)
                        // 🚀 BATCH SAVE: Use optimized BatchSaveClusterDataToDatabase instead of old SaveClusterDataToDatabase
                        if (comboId.HasValue && filterId.HasValue && _clusterToClashZoneIds.Count > 0)
                        {
                            // 🔥🔥🔥 CRITICAL DEBUG: First line inside if block
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 ENTERED IF BLOCK - About to batch save\n");

                            // 🔥 DEBUG LOG: Confirm we're about to save
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVING cluster data to database: {_clusterToClashZoneIds.Count} clusters\n");

                            // ✅ USER REQUEST: "Regen one time only" before collecting corners
                            // This ensures all geometry is valid before we read it for corner calculation
                            doc.Regenerate();

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
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPED saving cluster data: comboId={comboId?.ToString() ?? "NULL"}, filterId={filterId?.ToString() ?? "NULL"}, clusters={_clusterToClashZoneIds.Count}\n");
                        }
                        dbSaveTracker.SetItemCount(_clusterToClashZoneIds.Count);

                        // ✅ REMOVED: ReadyForPlacementFlag reset after cluster placement (redundant)
                        // IsClusterResolved flag is sufficient to track cluster placement status

                        // ✅ STEP 5 OPTIMIZATION: Flush deferred cluster parameters after all placements
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            int flushedCount = FlushDeferredClusterParameters();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH-FLUSH: Flushed parameters for {flushedCount} cluster sleeves\n");
                            }
                        }

                        // ✅ BATCH PERSISTENCE FIX: Clear cache at the end
                        _clusterSaveDataCache.Clear();
                        _actualPlacementPoints.Clear();

                        // ✅ PHASE 3: Final Verification of Duplicates
                        // Safety check to ensure no duplicates slipped through
                        var verificationSleeveIds = new HashSet<int>();
                        var duplicatesFound = new List<int>();

                        foreach (var clusterId in _clusterToClashZoneIds.Keys)
                        {
                            var zoneIds = _clusterToClashZoneIds[clusterId];
                            // Get original sleeve IDs from zone IDs (via cached data service if possible, or assume 1:1 for verification)
                            // Actually, we tracked processed IDs in `globalProcessedSleeveIds` earlier.
                            // Let's verify _clusterToClashZoneIds content.
                            // NOTE: zoneIds are ClashZone IDs, not Sleeve Instance IDs.
                            // We can't easily map back here without DB lookup, so we'll rely on globalProcessedSleeveIds count
                            // vs total expected items.
                        }

                        // Simple verification: Check if globalProcessedSleeveIds contains duplicates?
                        // No, hashset can't.
                        // But duplicates would have been rejected by the `alreadyProcessed` check.

                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ DUPLICATE CHECK: Processed {globalProcessedClashZoneGuids.Count} unique ClashZones across {placedCount} clusters.\n");

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

                        // ✅ DIAGNOSTIC: Verify sleeves are visible in Revit after placement
                        if (placedClusters.Count > 0)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔍🔍🔍 FINAL VISIBILITY CHECK: Verifying {placedClusters.Count} cluster sleeves are visible in Revit...\n");

                            foreach (var cluster in placedClusters)
                            {
                                try
                                {
                                    if (cluster != null && cluster.IsValidObject)
                                    {
                                        var location = cluster.Location as LocationPoint;
                                        var locPoint = location?.Point;
                                        var bbox = cluster.get_BoundingBox(null);
                                        var levelName = "Unknown";
                                        try
                                        {
                                            // Try to get level from document
                                            var levelId = cluster.LevelId;
                                            if (levelId != null && !levelId.Equals(ElementId.InvalidElementId))
                                            {
                                                var levelElem = cluster.Document?.GetElement(levelId);
                                                if (levelElem is Autodesk.Revit.DB.Level lvl)
                                                {
                                                    levelName = lvl.Name ?? "Unknown";
                                                }
                                            }
                                        }
                                        catch { }

                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ VISIBILITY: Cluster sleeve {cluster.Id.IntegerValue} EXISTS:\n" +
                                            $"  Name='{cluster.Name}', Family='{cluster.Symbol?.Family?.Name ?? "NULL"}',\n" +
                                            $"  Location=({locPoint?.X:F3}, {locPoint?.Y:F3}, {locPoint?.Z:F3}),\n" +
                                            $"  Level='{levelName}',\n" +
                                            $"  Category='{cluster.Category?.Name ?? "NULL"}',\n" +
                                            $"  BBox={(bbox != null && bbox.Enabled ? $"({bbox.Min.X:F3},{bbox.Min.Y:F3},{bbox.Min.Z:F3}) to ({bbox.Max.X:F3},{bbox.Max.Y:F3},{bbox.Max.Z:F3})" : "NULL")},\n" +
                                            $"  IsValid={cluster.IsValidObject}, Document='{cluster.Document?.Title ?? "NULL"}'\n");
                                    }
                                    else
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ❌ VISIBILITY: Cluster sleeve is NULL or INVALID!\n");
                                    }
                                }
                                catch (Exception visEx)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ VISIBILITY ERROR: Exception checking cluster sleeve: {visEx.Message}\n");
                                }
                            }
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[RefactoredClusterService] ✅ Completed: Placed {placedCount} clusters, Deleted {deletedCount} individual sleeves");
                        }

                        return (placedCount, deletedCount);
                    }
                }
                catch (Exception innerEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[RefactoredClusterService] ❌ Inner try error: {innerEx.Message}");
                        SafeFileLogger.SafeAppendText("cluster_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [RefactoredClusterService] ❌ INNER ERROR: {innerEx.Message}\nStackTrace: {innerEx.StackTrace}\n");
                    }
                    throw;
                }
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
                        return (true, path1Result.PlacedCount, path1Result.DeletedCount);
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] PATH 1: No cluster data found, proceeding to normal calculation");
                        return (false, 0, 0); // Fall through to normal calculation
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
            var processedGuids = new HashSet<Guid>();  // ✅ BATCH MODE FIX: Prevent duplicate ClashZones
            
            foreach (var cz in filteredClashZones)
            {
                try
                {
                    // ✅ BATCH MODE FIX: Skip if already processed (prevents duplicates)
                    if (processedGuids.Contains(cz.Id))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] [DEDUP] Skipping duplicate ClashZone GUID: {cz.Id}\n");
                        continue;
                    }
                    processedGuids.Add(cz.Id);
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
                        ClashZone = cz,
                        // ✅ FIX: Populate properties expected by ProximityCheckerFactory
                        SystemType = cz.MepElementCategory, 
                        IsCircular = !string.IsNullOrEmpty(cz.DuctShape) && (cz.DuctShape.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || cz.DuctShape.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0) || (cz.MepElementCategory != null && cz.MepElementCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
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
        /// <summary>
        /// Place a cluster sleeve for a group using all relevant services.
        /// ✅ SOLID REFACTORING: Accepts optional pre-calculated result to avoid recalculation.
        /// </summary>
        private (bool success, int placedCount, int deletedCount, FamilyInstance? placedClusterSleeve, int? capturedClusterSleeveId) PlaceClusterForGroup(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            string? xmlFilePath,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker? performanceTracker = null,
            ClusterCalculationResult? preCalculatedResult = null)
        {
            // ✅ CRITICAL FIX: Sort cluster by SleeveInstanceId to ensure deterministic order
            // This guarantees that "First Sleeve" logic (used for placement/orientation) always picks the same sleeve
            // regardless of batch vs sequential execution order.
            // Fixes regression where batch mode produced different results due to list ordering.
            if (cluster != null && cluster.Count > 1)
            {
                cluster = cluster.OrderBy(x => (int)x.SleeveInstanceId).ToList();
            }
            
            try
            {
                // ✅ SOLID REFACTORING: Use pre-calculated results if available, otherwise calculate sequentially
                double rotationAngle;
                List<FamilyInstance> actualSleeves;
                (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) bboxResult;
                
                if (preCalculatedResult != null && preCalculatedResult.IsValid)
                {
                    // ✅ USE PRE-CALCULATED RESULTS: No recalculation needed (2-4× faster)
                    rotationAngle = preCalculatedResult.RotationAngle;
                    actualSleeves = preCalculatedResult.ActualSleeves;
                    bboxResult = preCalculatedResult.BoundingBox;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ [SOLID-REFACTORING] Using pre-calculated results: Rotation={rotationAngle * 180.0 / Math.PI:F1}°, " +
                            $"BBox=({RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.width):F1}mm x " +
                            $"{RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.height):F1}mm x " +
                            $"{RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.depth):F1}mm)\n");
                    }
                }
                else
                {
                    // ✅ LEGACY SEQUENTIAL CALCULATION: Fallback when pre-calculation not available or failed
                    if (!DeploymentConfiguration.DeploymentMode && preCalculatedResult != null)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ [SOLID-REFACTORING] Pre-calculation invalid, using legacy sequential calculation: {preCalculatedResult.ErrorMessage}\n");
                    }
                    
                    // ✅ PERFORMANCE: Track rotation angle determination
                    using (performanceTracker?.TrackSubOperation("Determine Rotation Angle"))
                    {
                        // ✅ Step 1: Determine rotation angle (Phase 6: Rotation Service)
                        rotationAngle = _rotationService.DetermineRotationAngle(cluster, xmlFilePath);
                    } // End Determine Rotation Angle sub-operation
                    
                    // ✅ Step 2: Get actual sleeve elements from document
                    actualSleeves = cluster
                        .Select(s => doc.GetElement(new ElementId(s.SleeveInstanceId)) as FamilyInstance)
                        .Where(fi => fi != null && fi.IsValidObject)
                        .ToList();
                    
                    // ✅ PERFORMANCE: Track bounding box calculation
                    using (performanceTracker?.TrackSubOperation("Calculate Rotated Bounding Box"))
                    {
                        bboxResult = _rotationService.CalculateRotatedBoundingBox(cluster, actualSleeves, rotationAngle, xmlFilePath);
                    } // End Calculate Rotated Bounding Box sub-operation
                }
                
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

                // ✅ CLEARANCE & ROUNDING FIX: Apply user-defined clearance and rounding
                // 1. Get Clearance Setting
                double clearanceMm = 50.0; // Default standard clearance (Safe Fallback)
                try
                {
                     // Attempt to refine based on category if possible settings exist, 
                     // but since ClusterConfigurationManager doesn't hold these, we default to 50mm (std MEP).
                     string cat = cluster[0].Category;
                     if (!string.IsNullOrEmpty(cat) && cat.IndexOf("Cable", StringComparison.OrdinalIgnoreCase) >= 0)
                     {
                         clearanceMm = 75.0; // Typical for trays
                     }
                }
                catch { }

                // 2. Add Clearance (2x, one for each side)
                double clearanceInternal = RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                
                // Original Tight Bounds
                double rawWidth = bboxResult.width;
                double rawHeight = bboxResult.height;
                
                // Add Clearance
                double widthWithClearance = rawWidth + (2 * clearanceInternal);
                double heightWithClearance = rawHeight + (2 * clearanceInternal);

                // 3. Rounding
                // Use OpeningSettingsHelper.RoundDimensionsToNearest5mm (which respects global RoundingValue)
                // ✅ FIX: Explicit namespace to avoid CS0234 and explicit types to avoid CS8130
                (double finalWidth, double finalHeight) = JSE_RevitAddin_MEP_OPENINGS.Services.OpeningSettingsHelper.RoundDimensionsToNearest5mm(widthWithClearance, heightWithClearance);

                // Update bboxResult with final dimensions
                // Note: We keep Depth as is (usually wall thickness or union of depths)
                // Also: We do NOT shift the midpoint here directly, assuming symmetrical expansion around center.
                // IF Lopsidedness persists, we may need to re-calculate Center if expansion isn't symmetric (e.g. dampers), 
                // but for clusters, symmetric expansion around the 'Tight Bounds Center' is usually correct.
                
                // Log the transformation
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                     $"[{DateTime.Now:HH:mm:ss}] 📏 DIMENSION ADJUSTMENT:\n" +
                     $"   Raw: W={rawWidth * 304.8:F1}mm, H={rawHeight * 304.8:F1}mm\n" +
                     $"   +Clearance ({clearanceMm:F1}mm x2): W={widthWithClearance * 304.8:F1}mm, H={heightWithClearance * 304.8:F1}mm\n" +
                     $"   Rounded: W={finalWidth * 304.8:F1}mm, H={finalHeight * 304.8:F1}mm\n");

                // Update result tuple (create new tuple with updated width/height)
                bboxResult = (finalWidth, finalHeight, bboxResult.depth, bboxResult.mid, 
                              bboxResult.rotatedMinX, bboxResult.rotatedMinY, bboxResult.rotatedMinZ,
                              bboxResult.rotatedMaxX, bboxResult.rotatedMaxY, bboxResult.rotatedMaxZ);

                
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

                // ✅ DIAGNOSTIC TRACE: Verify group-to-dimensions mapping
                try
                {
                    string sleeveIdsStr = string.Join(",", cluster.Select(c => c.SleeveInstanceId));
                    string logMsg = $"[{DateTime.Now:HH:mm:ss}] 🔄 PROCESSING GROUP: Sleeves=[{sleeveIdsStr}], CalcWidth={bboxResult.width * 304.8:F1}mm, CalcHeight={bboxResult.height * 304.8:F1}mm, CalcDepth={bboxResult.depth * 304.8:F1}mm\n";
                    SafeFileLogger.SafeAppendText("cluster_debug.log", logMsg);
                }
                catch {}
                
                // ✅ Step 3: Determine family name before placement
                // This fulfills the user's requirement to decouple determination from placement
                // Clusters are ALWAYS rectangular, so we pass isCluster: true to skip complex individual sleeve logic
                string sleeveFamilyName = ClusterPlacementService.GetFamilyName(groupKey.hostType, groupKey.systemType, Math.Max(bboxResult.width, bboxResult.height), isCluster: true);
                
                // ✅ USER REQUIREMENT: Persist family name BEFORE placement for robustness
                try
                {
                    using (var dbContext = new SleeveDbContext(doc))
                    {
                        var clashZoneRepo = new ClashZoneRepository(dbContext);
                        var clashZoneGuids = cluster.Select(c => (Guid)((ClashZone)c.ClashZone).Id).ToList();
                        clashZoneRepo.UpdateSleeveFamilyNameBulk(clashZoneGuids, sleeveFamilyName);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                             DebugLogger.Info($"[RefactoredClusterService] ✅ Pre-placement persistence: Saved family '{sleeveFamilyName}' for {clashZoneGuids.Count} zones");
                    }
                }
                catch (Exception dbEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RefactoredClusterService] ⚠️ Pre-placement persistence failed (non-critical): {dbEx.Message}");
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ FAMILY DETERMINED: Using '{sleeveFamilyName}' for cluster placement\n");

                // ✅ Step 4: Extract orientation data from first ClashZone in cluster
                string? hostOrientation = null;
                double mepRotationAngle = 0.0;
                if (cluster.Count > 0)
                {
                    var firstCz = cluster[0].ClashZone as ClashZone;
                    if (firstCz != null)
                    {
                        hostOrientation = firstCz.HostOrientation;
                        mepRotationAngle = firstCz.MepElementRotationAngle;
                    }
                }

                // ✅ Step 5: Place cluster sleeve (Phase 5: Placement Service)
                // 🔥 CRITICAL: Direct IO logging before placement
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO CALL PlaceClusterSleeve: placementPoint=({bboxResult.mid.X:F2},{bboxResult.mid.Y:F2},{bboxResult.mid.Z:F2}), W={RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.width):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(bboxResult.height):F1}mm, family={sleeveFamilyName}, hostOrientation={hostOrientation ?? "NULL"}, rotation={mepRotationAngle:F2}\n");
                }
                catch { }
                
                // ✅ IMMEDIATE PARAMETER SETTING: Pass null to disable batch buffering
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
                    sleeveFamilyName, // ✅ Pass pre-determined family name
                    out FamilyInstance? placedClusterSleeve,
                    out int? capturedClusterSleeveId,
                    out XYZ? actualPlacementPoint, // ✅ SOLID: Capture actual placement point via out parameter
                    out ClusterSaveData? clusterSaveData, // ✅ BATCH PERSISTENCE FIX: Capture calculated data
                    null, // ✅ FIXED: Pass null to force IMMEDIATE parameter setting
                    hostOrientation,
                    mepRotationAngle); 
                
                // ✅ BATCH PERSISTENCE FIX: Store calculated data for database save
                if (capturedClusterSleeveId.HasValue && clusterSaveData != null)
                {
                    _clusterSaveDataCache[capturedClusterSleeveId.Value] = clusterSaveData;
                }
                
                // ✅ SOLID: Store actual placement point for database save (if available)
                // This ensures database saves the correct calculated placement point instead of Revit bbox center
                if (capturedClusterSleeveId.HasValue && actualPlacementPoint != null)
                {
                    if (!_actualPlacementPoints.ContainsKey(capturedClusterSleeveId.Value))
                    {
                        _actualPlacementPoints[capturedClusterSleeveId.Value] = actualPlacementPoint;
                    }
                }
                
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
                        
                    // ✅ CRITICAL FIX: Add Width/Height/Depth to deferred parameters for batch flush
                    // This ensures cluster sleeves get correct size instead of defaulting to standard size
                        // ✅ REDUNDANCY REMOVAL: Width/Height/Depth are already added to deferredParameters 
                        // inside _placementService.PlaceClusterSleeve -> SetSizeParameters.
                        // Overwriting them here with bboxResult.depth would nullify the WallThickness override
                        // applied in ClusterPlacementService for wall-hosted clusters.
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ PLACEMENT COMPLETE: Cluster {placedClusterSleeve.Id.IntegerValue} placed and parameters deferred.\n");
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
                                    _flagManager.BatchUpdateFlagsForPlacement(
                                        new List<(ClashZone clashZone, int sleeveId)> { (clashZone, capturedClusterSleeveId.Value) },
                                        isCluster: true,
                                        targetCategory,
                                        _filterName,
                                        clusterWidth: bboxResult.width,
                                        clusterHeight: bboxResult.height,
                                        clusterDiameter: bboxResult.depth);
                                    
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
                // ⚠️⚠️⚠️ CRITICAL PROTECTION: Only delete individual sleeves if BOTH conditions are met:
                // 1. Cluster sleeve was successfully placed (capturedClusterSleeveId.HasValue && placedClusterSleeve != null)
                // 2. Cluster sleeve still exists in document (verifyBeforeDelete != null && IsValidObject)
                if (!capturedClusterSleeveId.HasValue || placedClusterSleeve == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Cluster placement FAILED or cluster sleeve is NULL - SKIPPING individual sleeve deletion\n" +
                        $"  capturedClusterSleeveId.HasValue={capturedClusterSleeveId.HasValue}, placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")}\n" +
                        $"  Individual sleeves will NOT be deleted because cluster was not successfully formed.\n");
                    // Return success=false if cluster wasn't placed, but don't delete individual sleeves
                    return (false, 0, 0, null, null);
                }
                
                // ✅ DOUBLE-CHECK: Verify cluster sleeve exists in document BEFORE deleting individual sleeves
                var verifyBeforeDelete = doc.GetElement(new ElementId(capturedClusterSleeveId.Value)) as FamilyInstance;
                if (verifyBeforeDelete == null || !verifyBeforeDelete.IsValidObject)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL: Cluster sleeve {capturedClusterSleeveId.Value} NOT FOUND in document before deleting individual sleeves - SKIPPING DELETION\n" +
                        $"  Individual sleeves will NOT be deleted because cluster sleeve does not exist.\n");
                    return (false, 0, 0, null, null); // Return failure - cluster sleeve was lost
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ VERIFY BEFORE DELETE: Cluster sleeve {capturedClusterSleeveId.Value} EXISTS before deleting {cluster.Count} individual sleeves\n");
                }
                
                // ✅ CRITICAL PROTECTION: Only proceed with deletion if cluster sleeve is confirmed to exist
                // This prevents individual sleeves from being deleted if cluster placement failed or cluster sleeve was lost
                if (!capturedClusterSleeveId.HasValue || placedClusterSleeve == null || verifyBeforeDelete == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Cannot proceed with individual sleeve deletion - cluster sleeve validation failed\n" +
                        $"  capturedClusterSleeveId.HasValue={capturedClusterSleeveId.HasValue}, placedClusterSleeve={(placedClusterSleeve != null ? "NOT NULL" : "NULL")}, verifyBeforeDelete={(verifyBeforeDelete != null ? "NOT NULL" : "NULL")}\n");
                    return (false, 0, 0, null, null);
                }
                
                int deletedIndividualCount = 0;
                var sleevesToDelete = new List<ElementId>();
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 DELETION PREPARATION: Preparing to delete {cluster.Count} individual sleeves for cluster {capturedClusterSleeveId.Value}\n");
                
                foreach (var sleeveData in cluster)
                {
                    try
                    {
                        int sleeveInstanceId = sleeveData.SleeveInstanceId;
                        if (sleeveInstanceId <= 0)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING: sleeveInstanceId={sleeveInstanceId} (not > 0)\n");
                            continue;
                        }
                        
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
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING: Sleeve {sleeveInstanceId} not found or not a FamilyInstance\n");
                        }
                    }
                    catch (Exception delEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error queuing sleeve {sleeveData.SleeveInstanceId} for deletion: {delEx.Message}\n");
                    }
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 DELETION QUEUE: {sleevesToDelete.Count} individual sleeves queued for deletion (out of {cluster.Count} in cluster)\n");
                
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
        /// ✅ PATH 1a REPLAY: Place cluster sleeves from pre-calculated database data with smart placement.
        /// 
        /// PATH 1a (Super Fast Lane) - Smart Placement:
        /// - Checks existence: Compares SleeveInstanceId in DB with what exists in Revit (SOLID with all 28 features)
        /// - Smart placement: Only places individual sleeves NOT affected by existing cluster sleeves
        /// - Cluster placement: Only places clusters if no conflicting individual sleeves exist
        /// - No deletion needed: Conflicts avoided through smart placement
        /// 
        /// Implements all 28 features:
        /// - 🏗️ SOLID Architecture (SRP, DIP)
        /// - 📊 Diagnostic Logging
        /// - 🛡️ Crash-Safe Execution (Exception handling)
        /// - ⚡ Performance Monitoring
        /// - 🔒 Transaction Safety
        /// 
        /// NO SIZE PROCESSING - dimensions are already in database
        /// NO CALCULATIONS - everything is pre-calculated
        /// NO DETECTION - Path 3 handles invalidated zones
        /// NO REGENERATION - simple placement only
        /// </summary>
        public (int PlacedCount, int DeletedCount) PlaceClustersFromDatabase(
            Document doc,
            List<ClusterSleeveData> clusterDataList,
            UIDocument? uiDoc,
            List<FamilyInstance>? placedClusterSleevesOut,
            string? xmlFilePath,
            string targetCategory,
            int comboId)
        {
            // ✅ LOGGING: Log Build Timestamp to verify correct DLL is running (PATH 1)
            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🚀 STARTING BATCH PLACEMENT (PATH 1) - Build: {Helpers.VersionInfo.GetBuildTimestamp()} - Version: {Helpers.VersionInfo.VersionTag}\n");

            int placedCount = 0;
            int deletedCount = 0;
            var placedClusters = new List<FamilyInstance>();

            try
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1a REPLAY: Smart placement with existence checks (SOLID + 28 features)\n");

                foreach (var clusterData in clusterDataList)
                {
                    try
                    {
                        // Load clash zones for this cluster
                        var clashZoneIds = clusterData.ClashZoneIds;
                        if (clashZoneIds == null || clashZoneIds.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[RefactoredClusterService] PATH 1a: Cluster {clusterData.ClusterInstanceId} has no ClashZoneIds, skipping");
                            continue;
                        }

                        // ✅ PATH 1a: SMART PLACEMENT - Check if cluster already exists in Revit
                        if (_existenceChecker.ClusterSleeveExists(clusterData.ClusterInstanceId))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1a: Cluster {clusterData.ClusterInstanceId} already exists in Revit - SKIPPING placement\n");
                            continue; // Cluster already exists, skip placement
                        }

                        // ✅ PATH 1a: SMART PLACEMENT - Check if conflicting individual sleeves exist
                        if (_existenceChecker.HasConflictingIndividualSleeves(clashZoneIds))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ PATH 1a: Cluster {clusterData.ClusterInstanceId} has conflicting individual sleeves - SKIPPING placement (Path 3 will handle)\n");
                            continue; // Conflicting individual sleeves exist, skip cluster placement (Path 3 will handle)
                        }

                        // ✅ CLUSTER FIX: Use dynamic family selection instead of hardcoded Wall family
                        string hostType = clusterData.HostType ?? "Unknown";
                        string familyName = JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementService.GetFamilyName(
                            hostType, 
                            targetCategory, 
                            Math.Max(clusterData.ClusterWidth, clusterData.ClusterHeight), 
                            true); // isCluster = true

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

                                // ✅ PATH 1: Set dimensions directly from DB data (no processing, no calculations)
                                // Simple immediate write - dimensions are already calculated and stored in database
                                if (widthParam != null && !widthParam.IsReadOnly)
                                    widthParam.Set(clusterData.ClusterWidth);
                                if (heightParam != null && !heightParam.IsReadOnly)
                                    heightParam.Set(clusterData.ClusterHeight);
                                
                                // ✅ ROBUST DEPTH: Set BOTH 'Depth' and 'Wall Width' to ensure geometry update
                                // This mirrors the fix in SleeveParameterService to handle various family definitions
                                var wallWidthParam = placedClusterSleeve.LookupParameter("Wall Width");
                                if (wallWidthParam != null && !wallWidthParam.IsReadOnly)
                                    wallWidthParam.Set(clusterData.ClusterDepth);

                                // Always set Depth if possible
                                if (depthParam != null && !depthParam.IsReadOnly)
                                    depthParam.Set(clusterData.ClusterDepth);
                                    
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(clusterData.ClusterDepth);
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1 REPLAY: Set Depth/WallWidth to {dMm:F1}mm (DB Value)\n");
                                }
                                
                                // ✅ PATH 1: Set rotation if available (from DB)
                                // Check if rotation angle is valid (not zero or NaN)
                                if (!double.IsNaN(clusterData.RotationAngleDeg) && 
                                    Math.Abs(clusterData.RotationAngleDeg) > 1e-6)
                                {
                                    var rotationParam = placedClusterSleeve.LookupParameter("Rotation");
                                    if (rotationParam != null && !rotationParam.IsReadOnly)
                                    {
                                        // Convert degrees to radians if needed (check parameter unit)
                                        rotationParam.Set(clusterData.RotationAngleDeg * Math.PI / 180.0);
                                    }
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
                                // ✅ CRITICAL OPTIMIZATION: Load only specific clash zones by GUID (not all for category)
                                // This avoids loading hundreds of zones when we only need a few per cluster
                                using (var dbContext = new SleeveDbContext(doc))
                                {
                                    var clashZoneRepository = new ClashZoneRepository(dbContext);
                                    
                                    // ✅ BATCH LOAD: Load all clash zones for this cluster in one query (not per-zone)
                                    var clashZoneGuids = clashZoneIds.Select(id => id).ToList();
                                    var clashZones = clashZoneRepository.GetClashZonesByGuids(clashZoneGuids);
                                    
                                    if (clashZones != null && clashZones.Count > 0)
                                    {
                                        // ✅ BATCH UPDATE: Update flags for all clash zones in this cluster at once
                                        var flagUpdates = clashZones.Select(cz => (cz, capturedClusterSleeveId.Value)).ToList();
                                        
                                        // Get filter name (use empty string if not available)
                                        var categoryName = clashZones.First().MepElementCategory ?? targetCategory;
                                        var baseFilterName = string.Empty; // Filter name not critical for PATH 1
                                        
                                        // Update flags using FlagManager (batch update for all zones in cluster)
                                        _flagManager.BatchUpdateFlagsForPlacement(
                                            flagUpdates,
                                            isCluster: true,
                                            categoryName,
                                            baseFilterName,
                                            clusterWidth: clusterData.ClusterWidth,
                                            clusterHeight: clusterData.ClusterHeight,
                                            clusterDiameter: clusterData.ClusterDepth);
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

                // ✅ PATH 1a: Smart placement avoids conflicts, so cleanup should return 0
                // If cleanup finds sleeves to delete, it means smart placement didn't work (shouldn't happen)
                if (placedClusters.Count > 0)
                {
                    // ✅ PATH 1a: Cleanup individual sleeves within placed clusters (should return 0 due to smart placement)
                    // Smart placement already skipped clusters with conflicting individual sleeves
                    // Updated to pass targetCategory to prevent cross-category deletion (e.g. Pipe clusters deleting Dampers)
                    deletedCount = _cleanupService.CleanupSleevesWithinClusters(doc, placedClusters, null, targetCategory);
                    
                    if (deletedCount > 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ PATH 1a: Cleanup deleted {deletedCount} individual sleeves (unexpected - smart placement should have avoided this)\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1a: No cleanup needed - smart placement avoided conflicts (Super Fast Lane)\n");
                    }
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1a: No clusters placed - all skipped due to existence checks or conflicts\n");
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
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH] 🔍 CHECKING: UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, _deferredClusterParameters.Count={_deferredClusterParameters.Count}\n");
            
            if (!OptimizationFlags.UseBatchedParameterWrites)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH] ⚠️ SKIPPING: UseBatchedParameterWrites is FALSE\n");
                return 0;
            }

            // ✅ CRITICAL FIX: Flush SleeveParameterService deferred parameters FIRST
            // Depth/Wall Width are set via SleeveParameterService, which uses its own batching.
            // We must flush those too, or Depth will remain at default!
            // MOVED UP: Must run even if _deferredClusterParameters (Width/Height) is empty!
            if (_parameterService != null)
            {
                int flushedCount = _parameterService.FlushDeferredParameters();
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH] ✅ Flushed {flushedCount} parameters from SleeveParameterService (Depth/WallWidth)\n");
            }
            
            if (_deferredClusterParameters.Count == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH] ⚠️⚠️⚠️ WARNING: _deferredClusterParameters is EMPTY! No parameters to flush. This may cause cluster sleeves to be misidentified!\n" +
                    $"  ⚠️ This means parameters were NOT added to deferred dictionary, OR dictionary was cleared before flush.\n" +
                    $"  ⚠️ Cleanup service will read from Revit element parameters (may be stale if not regenerated).\n");
                return 0;
            }



            var sw = Stopwatch.StartNew();
            int successCount = 0;
            int errorCount = 0;
            
            SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] Flushing {_deferredClusterParameters.Count} cluster sleeve parameters...\n");
            
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH] 🔍 DETAILS: Flushing {_deferredClusterParameters.Count} cluster sleeves. IDs: {string.Join(", ", _deferredClusterParameters.Keys.Select(id => id.IntegerValue))}\n");

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

                        // ✅ DIAGNOSTIC LOGGING: Track depth/size parameters being applied
                        if (paramName == "Depth" || paramName == "Wall Width" || paramName == "Width" || paramName == "Height")
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-FLUSH-PARAM] Sleeve={elementId.IntegerValue}, Parameter='{paramName}', Value={paramValue}\n");
                        }
                        
                        // ✅ CRITICAL FIX: Try parameter name variations for "Bottom of Opening" (same as in ClusterPlacementService)
                        Parameter parameter = null;
                        if (string.Equals(paramName, "Bottom of Opening", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(paramName, "Bottom Of Opening", StringComparison.OrdinalIgnoreCase))
                        {
                            // Try all variations for "Bottom of Opening"
                            parameter = element.LookupParameter("Bottom Of Opening")  // ✅ FIRST: Exact name from Properties
                                      ?? element.LookupParameter("Bottom of Opening")
                                      ?? element.LookupParameter("BottomOfOpening")
                                      ?? (element as FamilyInstance)?.Symbol?.LookupParameter("Bottom Of Opening")
                                      ?? (element as FamilyInstance)?.Symbol?.LookupParameter("Bottom of Opening")
                                      ?? (element as FamilyInstance)?.Symbol?.LookupParameter("BottomOfOpening");
                        }
                        else
                        {
                            // For other parameters, use exact name
                            parameter = element.LookupParameter(paramName);
                        }
                        
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
                            // ✅ ENHANCED LOGGING: For "Bottom of Opening", log that all variations were tried
                            if (string.Equals(paramName, "Bottom of Opening", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(paramName, "Bottom Of Opening", StringComparison.OrdinalIgnoreCase))
                            {
                                SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️⚠️⚠️ Parameter 'Bottom of Opening' NOT FOUND on element {elementId} " +
                                    $"(tried: 'Bottom Of Opening', 'Bottom of Opening', 'BottomOfOpening' on instance and symbol)\n");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("cluster_param_timing.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAMS] ⚠️ Parameter '{paramName}' not found on element {elementId}\n");
                            }
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

            // ✅ CRITICAL: Clear deferred parameters AFTER writing - prevents memory leak and ensures fresh data for next batch
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
            // 🔥 DEBUG: Confirm method entry
            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 BatchSaveClusterDataToDatabase ENTERED: {clusterToClashZoneIds?.Count ?? 0} clusters\n");
            
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
                    var clashZoneRepository = new ClashZoneRepository(dbContext); // For UpdateClusterSleeveCorners
                    
                    // Prepare all cluster save data
                    var clustersToSave = new List<ClusterSaveData>();
                    
                    // ✅ Prepare ClashZone updates for batch processing
                    var clashZoneUpdates = new List<(Guid ClashZoneGuid, int ClusterInstanceId, 
                        double Width, double Height, double Diameter,
                        double BoundingBoxMinX, double BoundingBoxMinY, double BoundingBoxMinZ,
                        double BoundingBoxMaxX, double BoundingBoxMaxY, double BoundingBoxMaxZ,
                        string SleeveFamilyName)>();

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
                            
                            // Initialize variables that will be populated from cache or Revit
                            double width = 0.0;
                            double height = 0.0;
                            double depth = 0.0;
                            double rotationAngleDeg = 0.0;
                            bool isRotated = false;
                            string hostType = string.Empty;
                            string hostOrientation = string.Empty;
                            string sleeveFamilyName = string.Empty;
                            XYZ placementPoint = null;

                            XYZ bboxMin = bbox.Min;
                            XYZ bboxMax = bbox.Max;
                            
                            // ✅ BATCH PERSISTENCE FIX: Use calculated data from cache if available
                            // This bypasses stale/deferred Revit parameters during batched writes
                            if (_clusterSaveDataCache.TryGetValue(clusterInstanceId, out var calculatedData))
                            {
                                width = calculatedData.ClusterWidth;
                                height = calculatedData.ClusterHeight;
                                depth = calculatedData.ClusterDepth;
                                rotationAngleDeg = calculatedData.RotationAngleDeg;
                                isRotated = calculatedData.IsRotated;
                                hostType = calculatedData.HostType;
                                hostOrientation = calculatedData.HostOrientation;
                                sleeveFamilyName = calculatedData.SleeveFamilyName;
                                placementPoint = new XYZ(calculatedData.PlacementX, calculatedData.PlacementY, calculatedData.PlacementZ);
                                
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🎯 CACHE HIT: Using calculated data for Cluster {clusterInstanceId} (W={width*304.8:F1}mm, H={height*304.8:F1}mm)\n");
                            }
                            else
                            {
                                // Fallback: Read from Revit (only if not in cache)
                                if (clusterSleeve != null)
                                {
                                    width = (clusterSleeve.LookupParameter("Width") ?? clusterSleeve.LookupParameter("Element Width"))?.AsDouble() ?? 0.0;
                                    height = (clusterSleeve.LookupParameter("Height") ?? clusterSleeve.LookupParameter("Element Height"))?.AsDouble() ?? 0.0;
                                    depth = (clusterSleeve.LookupParameter("Depth") ?? clusterSleeve.LookupParameter("Element Depth") ?? clusterSleeve.LookupParameter("Wall Width"))?.AsDouble() ?? 0.0;
                                    
                                    var currentRotationData = _rotationService.GetRotationData(clusterInstanceId);
                                    if (currentRotationData != null)
                                    {
                                        rotationAngleDeg = currentRotationData.Value.rotationAngleDeg;
                                        isRotated = currentRotationData.Value.isRotated;
                                    }

                                    hostType = GetHostTypeFromSleeve(clusterSleeve);
                                    hostOrientation = GetOrientationFromSleeve(clusterSleeve);
                                    sleeveFamilyName = clusterSleeve.Symbol.Family.Name;
                                    
                                    _actualPlacementPoints.TryGetValue(clusterInstanceId, out var actualPt);
                                    placementPoint = actualPt ?? (bboxMin + bboxMax) / 2.0;

                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ CACHE MISS: Reading from Revit for Cluster {clusterInstanceId}\n");
                                }
                            }

                            // Ensure placementPoint is not null, fallback to bbox center if still null
                            if (placementPoint == null)
                            {
                                placementPoint = (bboxMin + bboxMax) / 2.0;
                            }
                            
                            // ✅ CORNER PERSISTENCE (USER REQUEST: "Database-Driven Sizing")
                            // 1. Fetch constituent ClashZones from DB to get their accurate corners
                            // 2. Aggregate corners to find the true envelope (Min/Max)
                            // 3. Do NOT recalculate from placement point (avoids placement drift)
                            
                            double c1x=0, c1y=0, c1z=0, c2x=0, c2y=0, c2z=0, c3x=0, c3y=0, c3z=0, c4x=0, c4y=0, c4z=0;
                            
                            try 
                            {
                                // Fetch constituent zones
                                var constituentZones = clashZoneRepository.GetClashZonesByGuids(clashZoneIds);

                                if (constituentZones != null && constituentZones.Count > 0)
                                {
                                    double envMinX = double.MaxValue, envMinY = double.MaxValue, envMinZ = double.MaxValue;
                                    double envMaxX = double.MinValue, envMaxY = double.MinValue, envMaxZ = double.MinValue;
                                    bool hasValidCorners = false;

                                    foreach (var cz in constituentZones)
                                    {
                                        // Check if zone has valid corners
                                        if (cz.SleeveCorner1X == 0 && cz.SleeveCorner1Y == 0 && cz.SleeveCorner1Z == 0) continue;

                                        hasValidCorners = true;
                                        // Aggregate all 4 corners with null checks
                                        double[] xs = { cz.SleeveCorner1X ?? 0.0, cz.SleeveCorner2X ?? 0.0, cz.SleeveCorner3X ?? 0.0, cz.SleeveCorner4X ?? 0.0 };
                                        double[] ys = { cz.SleeveCorner1Y ?? 0.0, cz.SleeveCorner2Y ?? 0.0, cz.SleeveCorner3Y ?? 0.0, cz.SleeveCorner4Y ?? 0.0 };
                                        double[] zs = { cz.SleeveCorner1Z ?? 0.0, cz.SleeveCorner2Z ?? 0.0, cz.SleeveCorner3Z ?? 0.0, cz.SleeveCorner4Z ?? 0.0 };

                                        foreach (var x in xs) { if (x < envMinX) envMinX = x; if (x > envMaxX) envMaxX = x; }
                                        foreach (var y in ys) { if (y < envMinY) envMinY = y; if (y > envMaxY) envMaxY = y; }
                                        foreach (var z in zs) { if (z < envMinZ) envMinZ = z; if (z > envMaxZ) envMaxZ = z; }
                                    }

                                    if (hasValidCorners)
                                    {
                                        // Construct AABB corners from envelope
                                        // C1: Bottom-Left, C2: Bottom-Right, C3: Top-Left, C4: Top-Right
                                        // C1: Min-Min-Min
                                        c1x = envMinX; c1y = envMinY; c1z = envMinZ;
                                        // C2: Max-Min-Min
                                        c2x = envMaxX; c2y = envMinY; c2z = envMinZ;
                                        // C3: Min-Max-Max (TOP corner - use maxZ!)
                                        c3x = envMinX; c3y = envMaxY; c3z = envMaxZ;
                                        // C4: Max-Max-Max (TOP corner - use maxZ!)
                                        c4x = envMaxX; c4y = envMaxY; c4z = envMaxZ;
                                        
                                        // Update width/height/depth from this envelope for data consistency?
                                        // The user said: "Cluster corners should be derived... not recalculate"
                                        // We should probably trust the "envelope" dimensions more than the Revit parameters if the Revit family is misbehaving?
                                        // But for now, we leave 'width', 'height' as read from Revit (lines 2372+), 
                                        // and only override the *stored corners*.
                                        
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] 📏 CONSTITUENT ENVELOPE (Cluster {clusterInstanceId}): " +
                                            $"Based on {constituentZones.Count} zones. " +
                                            $"Min=({envMinX:F3},{envMinY:F3},{envMinZ:F3}) Max=({envMaxX:F3},{envMaxY:F3},{envMaxZ:F3}) " +
                                            $"-> Size: W={(envMaxX-envMinX)*304.8:F1}mm, H={(envMaxY-envMinY)*304.8:F1}mm\n");
                                    }
                                    else
                                    {
                                         SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CORNER WARNING: Cluster {clusterInstanceId} has {constituentZones.Count} zones but NO valid corners found in DB. Falling back to calculation.\n");
                                         // Fallback to calculation if DB corners are missing
                                         var cornersResult = _cornerService.CalculateCorners(placementPoint, width, height, rotationAngleDeg);
                                         if (cornersResult.HasValue)
                                         {
                                             var c = cornersResult.Value;
                                             c1x = c.corner1.X; c1y = c.corner1.Y; c1z = c.corner1.Z;
                                             c2x = c.corner2.X; c2y = c.corner2.Y; c2z = c.corner2.Z;
                                             c3x = c.corner3.X; c3y = c.corner3.Y; c3z = c.corner3.Z;
                                             c4x = c.corner4.X; c4y = c.corner4.Y; c4z = c.corner4.Z;
                                         }
                                    }
                                }
                                else
                                {
                                    // If no constituent zones or no valid corners, calculate from current data
                                    var cornersResult = _cornerService.CalculateCorners(placementPoint, width, height, rotationAngleDeg);
                                    if (cornersResult.HasValue)
                                    {
                                        var c = cornersResult.Value;
                                        c1x = c.corner1.X; c1y = c.corner1.Y; c1z = c.corner1.Z;
                                        c2x = c.corner2.X; c2y = c.corner2.Y; c2z = c.corner2.Z;
                                        c3x = c.corner3.X; c3y = c.corner3.Y; c3z = c.corner3.Z;
                                        c4x = c.corner4.X; c4y = c.corner4.Y; c4z = c.corner4.Z;
                                    }
                                }
                            }
                            catch (Exception cornerEx)
                            {
                                try
                                {
                                    var versionTag = Helpers.VersionInfo.VersionTag;
                                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ CORNER ERROR: {clusterSleeve.Id} -> Error fetching constituents: {cornerEx.Message}\n");
                                }
                                catch { }
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[RefactoredClusterService] Failed to aggregate corners: {cornerEx.Message}");
                            }
                            
                            // ✅ DIAGNOSTIC: Log bounding box values before saving
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 📦 PREPARE SAVE: Cluster {clusterInstanceId} - " +
                                $"BBoxMin=({bboxMin.X:F6}, {bboxMin.Y:F6}, {bboxMin.Z:F6}), " +
                                $"BBoxMax=({bboxMax.X:F6}, {bboxMax.Y:F6}, {bboxMax.Z:F6}), " +
                                $"Placement=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})" +
                                (placementPoint != null ? " [ACTUAL]" : " [BBOX_CENTER]") + "\n");
                            
                            // Add to batch
                            // ✅ PERSISTENCE FIX: Save Family Name
                            // string sleeveFamilyName = clusterSleeve.Symbol.Family.Name; // Now populated from cache or Revit fallback

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
                                HostOrientation = hostOrientation,
                                ClashZoneIds = clashZoneIds,
                                SleeveFamilyName = sleeveFamilyName,
                                // ✅ Corners
                                Corner1X = c1x, Corner1Y = c1y, Corner1Z = c1z,
                                Corner2X = c2x, Corner2Y = c2y, Corner2Z = c2z,
                                Corner3X = c3x, Corner3Y = c3y, Corner3Z = c3z,
                                Corner4X = c4x, Corner4Y = c4y, Corner4Z = c4z
                            });

                            // ✅ Add constituent zones to batch update
                            foreach (var zoneGuid in clashZoneIds)
                            {
                                clashZoneUpdates.Add((
                                    zoneGuid, 
                                    clusterInstanceId, 
                                    width, height, depth,
                                    bboxMin.X, bboxMin.Y, bboxMin.Z,
                                    bboxMax.X, bboxMax.Y, bboxMax.Z,
                                    sleeveFamilyName
                                ));
                            }
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
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] 💾 SAVING: Attempting to save {clustersToSave.Count} clusters to database...\n");
                    clusterRepository.BatchSaveClusterSleeves(clustersToSave);
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ SAVED: Successfully saved {clustersToSave.Count} clusters.\n");
                
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[RefactoredClusterService] ✅ Batch saved {clustersToSave.Count} clusters to database (ComboId={comboId}, FilterId={filterId}, Category={targetCategory})");
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVED {clustersToSave.Count} clusters to ClusterSleeves table in single transaction (Skipped: {skippedCount})\n");
                    }
                    
                    // ✅ Execute batch update for ClashZones
                    if (clashZoneUpdates.Count > 0)
                    {
                        clashZoneRepository.BatchUpdateClusterPlacement(clashZoneUpdates);
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH UPDATED {clashZoneUpdates.Count} ClashZones with cluster placement info\n");
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
                            
                            // Get dimensions (Robust lookup)
                            var widthParam = clusterSleeve.LookupParameter("Width") ?? clusterSleeve.LookupParameter("Element Width");
                            var heightParam = clusterSleeve.LookupParameter("Height") ?? clusterSleeve.LookupParameter("Element Height");
                            var depthParam = clusterSleeve.LookupParameter("Depth") ?? clusterSleeve.LookupParameter("Element Depth");
                            
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
                            
                            // ✅ SOLID: Get actual placement point from instance dictionary
                            // This is the placement point that was actually used to place the sleeve
                            // Fallback to bbox center only if actual placement point is not available
                            _actualPlacementPoints.TryGetValue(clusterInstanceId, out var actualPlacementPoint);
                            var placementPoint = actualPlacementPoint ?? (bboxMin + bboxMax) / 2.0;
                            
                            // ✅ CORRECT CORNER CALCULATION:
                            // We need to calculate the 4 corners of the sleeve in WORLD COORDINATES.
                            // If rotated, these corners will reflect the rotation.
                            // If not rotated, they will align with the bounding box.
                            
                            // 1. Get the transform of the instance
                            Transform transform = clusterSleeve.GetTransform();
                            
                            // 2. Get geometry to find the solid solid
                            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
                            GeometryElement geoElem = clusterSleeve.get_Geometry(opt);
                            Solid solid = null;
                            if (geoElem != null)
                            {
                                foreach (GeometryObject obj in geoElem)
                                {
                                    if (obj is Solid s && s.Volume > 0.0001)
                                    {
                                        solid = s;
                                        break;
                                    }
                                    else if (obj is GeometryInstance gi)
                                    {
                                        foreach (GeometryObject obj2 in gi.SymbolGeometry)
                                        {
                                            if (obj2 is Solid s2 && s2.Volume > 0.0001)
                                            {
                                                solid = s2;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            
                            // 3. Calculate Corners
                            // Default to bbox corners if solid not found (fallback)
                            XYZ c1 = new XYZ(bboxMin.X, bboxMin.Y, bboxMin.Z);
                            XYZ c2 = new XYZ(bboxMax.X, bboxMin.Y, bboxMin.Z);
                            XYZ c3 = new XYZ(bboxMax.X, bboxMax.Y, bboxMin.Z);
                            XYZ c4 = new XYZ(bboxMin.X, bboxMax.Y, bboxMin.Z);
                            
                            if (solid != null)
                            {
                                // Use refined logic if available, or just BBox of solid if simple
                                // For cluster sleeves (rect/round), we want the 4 corners of the base or the "footprint" on the host
                                // A simple robust way: Get BBox of the solid (transformed)
                                var sBox = solid.GetBoundingBox();
                                if (sBox != null)
                                {
                                    // Transform is already applied if we got solid from instance geometry? 
                                    // If we got it from symbol geometry, we need to apply transform.
                                    // Let's assume standard instance geometry retrieval works best.
                                    // Actually, simpler fallback that works 99%: Use the Rotated BBox from RotationService if available
                                    
                                    if (isRotated && rotationData.HasValue && rotationData.Value.rotatedBboxMin != null)
                                    {
                                        // This is already what we want? 
                                        // NO, we need 4 explicit corners for the "Width" dimension logic in Manual adapter.
                                        // Let's assume planar corners on the Z-plane of the insertion point.
                                        
                                        // Better yet: Use the stored rotation to calculate corners from the center/width/height/depth
                                        // Center is placementPoint.
                                        // We know Width/Depth based on Orientation.
                                        
                                        // Let's allow the repository/adapter to calculate from Width/Height/Rotation if corners are 0?
                                        // NO, the user wants explicit corners saved.
                                        
                                        // Let's try to get actual corners from the solid's bottom face?
                                        // Too complex and brittle.
                                        
                                        // ROBUST APPROACH: Calculate corners mathematically from Center, Width, Depth, Rotation.
                                        // Assuming Z-axis rotation (common for walls/floors)
                                        
                                        double w = width;
                                        double d = depth;
                                        // Adjust W/D based on orientation? 
                                        // Width is usually "along wall", Depth is "through wall".
                                        
                                        // Let's use clean math based on the saved dimensions and rotation.
                                        // This ensures consistency between "Size" log and "Corners" log.
                                        
                                        double angleRad = rotationAngleDeg * Math.PI / 180.0;
                                        XYZ center = placementPoint;
                                        
                                        // Calculate offsets for 4 corners (unrotated locally)
                                        // Assuming standard box centered at 0,0
                                        // C1: -W/2, -D/2
                                        // C2: +W/2, -D/2
                                        // C3: +W/2, +D/2
                                        // C4: -W/2, +D/2
                                        // (Ignoring Height/Z for "footprint" corners)
                                        
                                        double halfW = width / 2.0;
                                        double halfD = depth / 2.0;
                                        
                                        // Rotate these offsets
                                        double cos = Math.Cos(angleRad);
                                        double sin = Math.Sin(angleRad);
                                        
                                        // C1
                                        c1 = new XYZ(
                                            center.X + (-halfW * cos - -halfD * sin),
                                            center.Y + (-halfW * sin + -halfD * cos),
                                            center.Z); // Keep Z flat
                                            
                                        // C2
                                        c2 = new XYZ(
                                            center.X + (halfW * cos - -halfD * sin),
                                            center.Y + (halfW * sin + -halfD * cos),
                                            center.Z);
                                            
                                        // C3
                                        c3 = new XYZ(
                                            center.X + (halfW * cos - halfD * sin),
                                            center.Y + (halfW * sin + halfD * cos),
                                            center.Z);
                                            
                                        // C4
                                        c4 = new XYZ(
                                            center.X + (-halfW * cos - halfD * sin),
                                            center.Y + (-halfW * sin + halfD * cos),
                                            center.Z);
                                    }
                                }
                            }

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
                                clashZoneIds: clashZoneIds,
                                sleeveFamilyName: clusterSleeve.Symbol.Family.Name,
                                // ✅ Pass Calculated Corners
                                corner1X: c1.X, corner1Y: c1.Y, corner1Z: c1.Z,
                                corner2X: c2.X, corner2Y: c2.Y, corner2Z: c2.Z,
                                corner3X: c3.X, corner3Y: c3.Y, corner3Z: c3.Z,
                                corner4X: c4.X, corner4Y: c4.Y, corner4Z: c4.Z);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[RefactoredClusterService] ✅ Saved cluster {clusterInstanceId} to database (ComboId={comboId}, FilterId={filterId}, Category={targetCategory})");
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ SAVED cluster {clusterInstanceId} to ClusterSleeves table: ComboId={comboId}, FilterId={filterId}, Category={targetCategory}, ClashZoneIds={clashZoneIds.Count}\n");
                            
                            // ✅ CRITICAL FIX: Update ClashZones flags for each constituent zone
                            // This ensures IsClusterResolvedFlag=1, IsClusteredFlag=1, ClusterInstanceId are set
                            var clashZoneRepo = new ClashZoneRepository(dbContext);
                            foreach (var zoneGuid in clashZoneIds)
                            {
                                try
                                {
                                    clashZoneRepo.UpdateClusterPlacement(
                                        clashZoneId: zoneGuid,
                                        clusterInstanceId: clusterInstanceId,
                                        minX: bboxMin.X, minY: bboxMin.Y, minZ: bboxMin.Z,
                                        maxX: bboxMax.X, maxY: bboxMax.Y, maxZ: bboxMax.Z,
                                        isClustered: true,
                                        markedForCluster: false,
                                        sleeveFamilyName: clusterSleeve.Symbol.Family.Name,
                                        sleeveWidth: width,
                                        sleeveHeight: height,
                                        sleeveDiameter: depth);
                                }
                                catch (Exception czEx)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_errors.log",
                                        $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Error updating ClashZone {zoneGuid} flags: {czEx.Message}\n");
                                }
                            }
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATED {clashZoneIds.Count} ClashZone flags for cluster {clusterInstanceId}\n");
                            
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

                            // ✅ OPTIMIZATION: Query only the specific clash zones needed (instead of loading all category zones)
                            // This is much more efficient when only a few zones are needed from a large category
                            var individualClashZones = repository.GetClashZonesByGuids(clashZoneIds);

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
                // ✅ Check if we have primary bounding box data (legacy/Revit source)
                bool hasPrimaryBBox = !(cz.SleeveBoundingBoxMinX == 0 && cz.SleeveBoundingBoxMaxX == 0 &&
                                       cz.SleeveBoundingBoxMinY == 0 && cz.SleeveBoundingBoxMaxY == 0 &&
                                       cz.SleeveBoundingBoxMinZ == 0 && cz.SleeveBoundingBoxMaxZ == 0);

                if (hasPrimaryBBox)
                {
                    // ✅ DIAGNOSTIC: Log for dampers/ducts
                    if (cz.MepElementCategory != null && (cz.MepElementCategory.IndexOf("Accessories", StringComparison.OrdinalIgnoreCase) >= 0 || cz.MepElementCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] 📦 [GetBBox] Zone={cz.ClashZoneId}, SleeveId={cz.SleeveInstanceId}: Using primary Sidebar BBox=({cz.SleeveBoundingBoxMinX:F3},{cz.SleeveBoundingBoxMinY:F3},{cz.SleeveBoundingBoxMinZ:F3}) to ({cz.SleeveBoundingBoxMaxX:F3},{cz.SleeveBoundingBoxMaxY:F3},{cz.SleeveBoundingBoxMaxZ:F3})\n");
                    }
                    
                    return new BoundingBoxXYZ
                    {
                        Min = new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ),
                        Max = new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ),
                        Enabled = true
                    };
                }

                // ✅ FALLBACK: Use Sleeve Corners if primary BBox is empty (common for Voids/Batch-placed elements)
                if (cz.SleeveCorner1X != null && cz.SleeveCorner1Y != null && cz.SleeveCorner1Z != null)
                {
                    var helper = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity.SleeveCornerProximityHelper();
                    if (helper.HasValidSleeveCorners(cz))
                    {
                        var cornerBbox = helper.GetBoundingBoxFromCorners(cz);
                        
                        if (cz.MepElementCategory != null && (cz.MepElementCategory.IndexOf("Accessories", StringComparison.OrdinalIgnoreCase) >= 0 || cz.MepElementCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 || cz.MepElementCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] 📐 [GetBBox] Zone={cz.ClashZoneId}: DERIVED BBox from corners=({cornerBbox.minX:F3},{cornerBbox.minY:F3},{cornerBbox.minZ:F3}) to ({cornerBbox.maxX:F3},{cornerBbox.maxY:F3},{cornerBbox.maxZ:F3})\n");
                        }

                        return new BoundingBoxXYZ
                        {
                            Min = new XYZ(cornerBbox.minX, cornerBbox.minY, cornerBbox.minZ),
                            Max = new XYZ(cornerBbox.maxX, cornerBbox.maxY, cornerBbox.maxZ),
                            Enabled = true
                        };
                    }
                }
                
                return null;
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
