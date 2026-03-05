using System;
using System.Collections.Generic;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// Service for Phase 2: Sequential Placement from Database - V2 Architecture
    /// Responsibilities:
    /// 1. Query 'Pending' clusters from ClusterSleeves_v2
    /// 2. Place Family Instances (Sequential, Single Thread)
    /// 3. Update DB (Status='Placed', ClusterInstanceId)
    /// </summary>
    public class BatchClusterPlacementService
    {
        private readonly string _databasePath;
        private readonly IClashZoneRepository _repository;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor _performanceMonitor;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService _parameterService;
        private readonly IClusterCleanupService _cleanupService;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IParameterSnapshotTransferService _snapshotTransferService;

        public BatchClusterPlacementService(
            string databasePath, 
            IClashZoneRepository repository,
            JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService parameterService,
            IClusterCleanupService cleanupService = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor performanceMonitor = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IParameterSnapshotTransferService snapshotTransferService = null)
        {
            _databasePath = databasePath;
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _parameterService = parameterService ?? throw new ArgumentNullException(nameof(parameterService));
            _cleanupService = cleanupService ?? new ClusterCleanupService(); // Default if not injected
            _performanceMonitor = performanceMonitor;
            _snapshotTransferService = snapshotTransferService;
        }

        /// <summary>
        /// ✅ CONSOLIDATED PLACEMENT: Place all pending clusters from all categories in one stroke
        /// Then run cleanup once for all placed clusters
        /// </summary>
        public (int placed, int failed, int cleanedUp) PlaceAllCategoriesAndCleanup(Document doc, IClusterCleanupService cleanupService, bool useSingleTransaction = true)
        {
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🚀 CONSOLIDATED PLACEMENT: Starting placement for ALL categories\n");
            
            // Step 1: Place all pending clusters (all categories)
            var (placed, failed) = PlaceFromDatabase(doc, batchId: null, useSingleTransaction);
            
            // Step 2: Cleanup once for all placed clusters
            int cleanedUp = 0;
            if (cleanupService != null && placed > 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CONSOLIDATED CLEANUP: Starting cleanup for ALL categories after placement\n");
                
                try
                {
                    // Get all cluster instance IDs that were just placed (from all categories)
                    // Query ClusterSleeves_v2 for all clusters with Status='Placed' and ClusterInstanceId > 0
                    var clusterInstanceIds = new List<int>();
                    using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                    {
                        conn.Open();
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = @"
                                SELECT DISTINCT ClusterInstanceId 
                                FROM ClusterSleeves_v2 
                                WHERE Status = 'Placed' AND ClusterInstanceId > 0";
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    clusterInstanceIds.Add(reader.GetInt32(0));
                                }
                            }
                        }
                    }
                    
                    if (clusterInstanceIds.Count > 0)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CONSOLIDATED CLEANUP: Found {clusterInstanceIds.Count} placed clusters — running Stage 2 cleanup\n");
                        
                        // Bbox/corner prep is best-effort; cleanup always runs (uses DB backfill if bbox was not updated)
                        try
                        {
                            var instancesForBbox = new List<FamilyInstance>();
                            foreach (int id in clusterInstanceIds)
                            {
#if REVIT2023
                                var elem = doc.GetElement(new Autodesk.Revit.DB.ElementId(id)) as FamilyInstance;
#else
                                var elem = doc.GetElement(new Autodesk.Revit.DB.ElementId((long)id)) as FamilyInstance;
#endif
                                if (elem != null && elem.IsValidObject) instancesForBbox.Add(elem);
                            }
                            if (instancesForBbox.Count > 0)
                            {
                                try
                                {
                                    UpdateClusterBoundingBoxesAfterPlacement(doc, instancesForBbox);
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2: Ensured {instancesForBbox.Count} cluster bboxes in DB before cleanup\n");
                                    try
                                    {
                                        var cornerExtractor = new JSE_RevitAddin_MEP_OPENINGS.Services.Calculation.BatchSleeveCornerExtractor(_repository);
                                        int cornersExtracted = cornerExtractor.ExtractAndSaveCornersForClusters(doc);
                                        SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 📐 STAGE 2: Extracted corners for {cornersExtracted} cluster sleeves.\n");
                                    }
                                    catch (Exception cornerEx)
                                    {
                                        SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 corner extraction failed (continuing): {cornerEx.Message}\n");
                                    }
                                }
                                catch (Exception bboxEx)
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 bbox update failed (cleanup will still run): {bboxEx.Message}\n");
                                }
                            }
                        }
                        catch (Exception prepEx)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 prep failed (cleanup will still run): {prepEx.Message}\n");
                        }
                        
                        // Always run cleanup when we have placed clusters (cleanup service backfills bbox from placement+dimensions if needed)
                        cleanedUp = cleanupService.CleanupSleevesWithinClustersFromDatabase(
                            doc, 
                            targetCategory: null,
                            clusterInstanceIds: clusterInstanceIds);
                        
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CONSOLIDATED CLEANUP COMPLETE: Deleted {cleanedUp} individual sleeves (all categories)\n");
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ CONSOLIDATED CLEANUP FAILED: {ex.Message}\n{ex.StackTrace}\n");
                }
            }
            
            return (placed, failed, cleanedUp);
        }

        public (int placed, int failed) PlaceFromDatabase(Document doc, string batchId, bool useSingleTransaction = true)
        {
            // 🔥 DIAGNOSTIC: Log entry and transaction state at method start
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🟢 PlaceFromDatabase ENTRY: batchId={batchId}, useSingleTransaction={useSingleTransaction}, doc.IsModifiable={doc.IsModifiable}\n");
            
            // ✅ FIX: Query pending clusters FIRST before starting performance tracker
            List<BatchClusterData> pendingClusters;
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: About to call GetPendingClusters with batchId={(batchId ?? "NULL")}\n");
            using (var step1Tracker = _performanceMonitor?.TrackOperation("Step 1: LOAD CLUSTERS FROM DB"))
            {
                pendingClusters = GetPendingClusters(doc, batchId);
                step1Tracker?.SetItemCount(pendingClusters.Count);
            }
            
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: GetPendingClusters returned {pendingClusters.Count} clusters\n");
            _performanceMonitor?.AddZonesProcessed(pendingClusters.Count);
            
            // ✅ FIX: Early return BEFORE starting the main tracker if no clusters to place
            if (pendingClusters.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: No pending clusters found for batch {batchId}, returning early.\n");
                return (0, 0);
            }
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📂 LOAD: Found {pendingClusters.Count} pending clusters to place for batch {batchId}\n");
            
            // ✅ FIX: Only start the main tracker if there are clusters to place
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ STARTING PLACEMENT V2: Pending={pendingClusters.Count}, Mode={(useSingleTransaction ? "BULK" : "SEQUENTIAL")}\n");
            
            using (var tracker = _performanceMonitor?.TrackOperation("Step 2: CLUSTER PLACEMENT MAIN LOOP"))
            {
                tracker?.SetItemCount(pendingClusters.Count);
                int placedCount = 0;
                int failedCount = 0;
                var placedInstances = new List<FamilyInstance>(); // Track placed instances for 2nd cleanup

                // ✅ OPTIMIZATION: Pre-fetch and pre-activate all required Family Symbols (like individual bulk path)
                var symbolCache = new Dictionary<string, FamilySymbol>();
                using (_performanceMonitor?.TrackOperation("Step 1a: CACHE CLUSTER SYMBOLS"))
                {
                    var uniqueFamilyNames = pendingClusters
                        .Select(c => c.FamilyName)
                        .Where(n => !string.IsNullOrWhiteSpace(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    // ✅ CRITICAL FIX: Ensure all required cluster families are loaded into the document
                    // Previously, bulk placement assumed families were already loaded, which caused
                    // "Placed=0" when clusters existed in DB but their families were not yet in the model.
                    // This pre-loads any missing cluster families (Rectangular/Circular On Wall/Slab, etc.)
                    ClusterPlacementService.PreLoadClusterFamilies(doc, uniqueFamilyNames);

                    var allSymbols = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .ToList();

                    foreach (var name in uniqueFamilyNames)
                    {
                        var symbol = allSymbols.FirstOrDefault(s =>
                            s.Name.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                            s.Family.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                        if (symbol != null)
                        {
                            symbolCache[name] = symbol;
                            if (!symbol.IsActive) symbol.Activate();
                        }
                        else if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLUSTER-FAMILY-MISS: No FamilySymbol found for cluster family '{name}' after pre-load.\n");
                        }
                    }
                }

                // ✅ OPTIMIZATION: Pre-fetch all constituent ClashZones to avoid per-cluster DB queries
                var zoneCache = new Dictionary<Guid, ClashZone>();
                using (_performanceMonitor?.TrackOperation("Step 1b: PRE-FETCH CLUSTER CONTEXT"))
                {
                    var allConstituentGuids = pendingClusters
                        .SelectMany(c => (c.ConstituentZoneGuids ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                        .Select(s => Guid.TryParse(s.Trim(), out var g) ? g : Guid.Empty)
                        .Where(g => g != Guid.Empty)
                        .Distinct()
                        .ToList();

                    if (allConstituentGuids.Count > 0)
                    {
                        var fetchedZones = _repository.GetClashZonesByGuids(allConstituentGuids);
                        foreach (var z in fetchedZones) zoneCache[z.Id] = z;
                    }
                }

                // ✅ RESOLVE LEVELS: Populate MepElementLevelName for each cluster from its constituent zones
                foreach (var cluster in pendingClusters)
                {
                    if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                    {
                        var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                        if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid g))
                        {
                            if (zoneCache.TryGetValue(g, out var z))
                            {
                                cluster.MepElementLevelName = z.MepElementLevelName;
                            }
                        }
                    }
                }

                // ✅ LEVEL CACHING: Pre-fetch all Level objects to avoid slow spatial lookups during placement
                var levelMap = new Dictionary<string, ElementId>(StringComparer.OrdinalIgnoreCase);
                using (_performanceMonitor?.TrackOperation("Step 1c: CACHE LEVELS"))
                {
                    var distinctLevelNames = pendingClusters
                        .Select(c => c.MepElementLevelName)
                        .Where(l => !string.IsNullOrEmpty(l))
                        .Distinct()
                        .ToList();

                    if (distinctLevelNames.Any())
                    {
                        var levels = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .Where(l => distinctLevelNames.Contains(l.Name, StringComparer.OrdinalIgnoreCase))
                            .ToList();

                        foreach (var l in levels) levelMap[l.Name] = l.Id;
                    }
                }
                
                // ✅ CLUSTER OPTIMIZATION START (SINGLE TRANSACTION FLOW)
                // We manage the transaction if one doesn't exist.
                // If one DOES exist, we use it, but assume the caller will commit.
                // WE ALWAYS USE PlaceBulkClusters (NewFamilyInstances2) for speed.

                bool manageTransaction = !doc.IsModifiable;
                Transaction mainTransaction = null;
                bool transactionCommitted = false;
                ClusterPersistenceData clusterPersistenceData = null;

                try
                {
                    // PERF DIAGNOSTIC: Tick-precision timing for cluster transaction phases
                    var clusterSubSw = new System.Diagnostics.Stopwatch();
                    double clusterTicksPerMs = System.Diagnostics.Stopwatch.Frequency / 1000.0;
                    long bulkPlaceTicks = 0, snapshotTransferTicks = 0, paramFlushTicks = 0;

                    SwallowWarningsPreprocessor clusterWarningHandler = null;

                    // 1. PRE-PROCESS: Prepare data and identify sleeves to delete
                    PreparedPlacementData preparedData = null;
                    using (var prepTracker = _performanceMonitor?.TrackOperation("Step 2a: PREPARE BULK PLACEMENT DATA"))
                    {
                        preparedData = PrepareBulkPlacementData(doc, pendingClusters, symbolCache, zoneCache, prepTracker, levelMap);
                    }

                    // 2. PREPARED DATA (Identify sleeves to delete later)
                    // We no longer delete here. We delete after the main placement transaction commits.
                    if (preparedData != null && preparedData.IndividualSleevesToDelete.Count > 0)
                    {
                        // Store these for deletion later
                    }

                    // 3. START MAIN PLACEMENT TRANSACTION
                    if (manageTransaction)
                    {
                        mainTransaction = new Transaction(doc, "Place Batch Clusters V2 (Unified)");
                        clusterWarningHandler = new SwallowWarningsPreprocessor();
                        var options = mainTransaction.GetFailureHandlingOptions();
                        options.SetFailuresPreprocessor(clusterWarningHandler);
                        mainTransaction.SetFailureHandlingOptions(options);

                        mainTransaction.Start();
                    }

                    // 4. PLACE CLUSTERS (NewFamilyInstances2)
                    using (var placementTracker = _performanceMonitor?.TrackOperation("Step 2b: EXECUTE BULK CLUSTER PLACEMENT"))
                    {
                        if (OptimizationFlags.UseBulkClusterSleevePlacement && preparedData != null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}] USING BULK PLACEMENT PATH (calling PlaceBulkClusters)\n");
                            clusterSubSw.Restart();
                            var (actualPlaced, persistData) = PlaceBulkClusters(doc, preparedData, zoneCache, placementTracker, placedInstances);
                            bulkPlaceTicks = clusterSubSw.ElapsedTicks;
                            placedCount = actualPlaced;
                            clusterPersistenceData = persistData;
                            placementTracker?.SetItemCount(actualPlaced);
                        }
                        else if (preparedData != null)
                        {
                            // Fallback to sequential Loop (only if flag is specifically disabled)
                            // This path should ideally be deprecated
                            const int batchLogIntervalSeq = 100;
                            int idx = 0;
                            foreach (var cluster in pendingClusters)
                            {
                                bool logDetail = (idx % batchLogIntervalSeq == 0 || idx == 0);
                                if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances, logDetail, levelMap)) placedCount++;
                                else failedCount++;
                                idx++;
                            }
                            if (pendingClusters.Count > 0)
                                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ Sequential: Set parameters for all {placedCount} cluster instances\n");
                        }
                    }

                    // 2. SNAPSHOT PARAMETER TRANSFER + FLUSH PARAMETERS (Inside the transaction!)
                    // ✅ NOTE: For Bulk placement, parameter handling is now handled INSIDE PlaceBulkClusters 
                    // to ensure dimensions are current for the persistence phase.
                    // This call remains as a safety for the sequential path or missed flushes.
                    if (OptimizationFlags.UseBatchedParameterWrites && placedCount > 0)
                    {
                        // ✅ SNAPSHOT PARAMETER TRANSFER: Transfer MEP parameters from snapshots to placed cluster sleeves
                        if (OptimizationFlags.EnableSnapshotParameterTransfer && _snapshotTransferService != null && placedInstances.Count > 0)
                        {
                            try
                            {
                                using (var snapTracker = _performanceMonitor?.TrackOperation("Step 3b: SNAPSHOT PARAMETER TRANSFER"))
                                {
                                    var clusterElementIds = placedInstances.Select(fi => fi.Id).ToList();
                                    clusterSubSw.Restart();
                                    int deferredCount = _snapshotTransferService.TransferSnapshotParameters(doc, clusterElementIds, _repository, _parameterService);
                                    snapshotTransferTicks = clusterSubSw.ElapsedTicks;
                                    snapTracker?.SetItemCount(deferredCount);
                                }
                            }
                            catch (Exception snapEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] SNAPSHOT TRANSFER FAILED: {snapEx.Message}\n");
                            }
                        }
                        
                        using (var flushTracker = _performanceMonitor?.TrackOperation("Step 4: FLUSH CLUSTER PARAMETERS (Safety)"))
                        {
                            clusterSubSw.Restart();
                            int flushedCount = _parameterService.FlushDeferredParameters(clearList: true, context: "Cluster-Safety");
                            paramFlushTicks += clusterSubSw.ElapsedTicks;
                            flushTracker?.SetItemCount(flushedCount);
                        }
                    }

                    // ✅ PERF: Fire DB persistence on background thread BEFORE Revit commit
                    // DB work (284ms) runs in parallel with Revit commit (377ms) → saves ~284ms
                    Task dbTask = null;
                    if (clusterPersistenceData != null)
                    {
                        var capturedData = clusterPersistenceData;
                        dbTask = Task.Run(() =>
                        {
                            try
                            {
                                PersistClusterPlacementToDb(capturedData);
                            }
                            catch (Exception dbEx)
                            {
                                SafeFileLogger.SafeAppendText("placement_errors.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ BACKGROUND DB PERSISTENCE FAILED: {dbEx.Message}\n{dbEx.StackTrace}\n");
                            }
                        });
                    }

                    // PERF DIAGNOSTIC: Log pre-commit timing summary
                    _performanceMonitor?.LogMetric("CLUSTER PRE-COMMIT TIMING",
                        $"BulkPlace(total)={bulkPlaceTicks / clusterTicksPerMs:F1}ms, " +
                        $"SnapshotTransfer={snapshotTransferTicks / clusterTicksPerMs:F1}ms, " +
                        $"SafetyFlush={paramFlushTicks / clusterTicksPerMs:F1}ms, " +
                        $"Clusters={placedCount}");

                    if (manageTransaction)
                    {
                        // 🔬 DIAGNOSTIC: Split commit cost into (a) Regeneration vs (b) Undo-buffer write
                        // doc.Regenerate() inside the TX forces Revit to compute geometry NOW.
                        // If commit is fast after this, regen was the dominant cost.
                        // If commit is still slow, undo-buffer write / constraint resolution is the cost.
                        using (_performanceMonitor?.TrackOperation("Revit API: Global Refresh (Regenerate)"))
                        {
                            var regenSw = System.Diagnostics.Stopwatch.StartNew();
                            doc.Regenerate();
                            regenSw.Stop();
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔬 PRE-COMMIT REGEN: {regenSw.ElapsedMilliseconds}ms (clusters={placedCount})\n");
                        }

                        using (var commitTracker = _performanceMonitor?.TrackOperation("Transaction Commit (Cluster)"))
                        {
                            mainTransaction.Commit();
                        }
                        transactionCommitted = true;
                        if (clusterWarningHandler != null)
                            _performanceMonitor?.LogMetric("CLUSTER TX WARNINGS",
                                $"Deleted={clusterWarningHandler.WarningsDeleted}, Errors={clusterWarningHandler.ErrorsFound}");
                    }

                    // 3. UPDATE BOUNDING BOXES (After commit — Revit auto-regens on commit; then we read bbox)
                    if (placedInstances.Count > 0)
                    {
                        using (var bboxTracker = _performanceMonitor?.TrackOperation("Step 5: RETRIEVE CLUSTER DATA FROM MODEL"))
                        {
                            UpdateClusterBoundingBoxesAfterPlacement(doc, placedInstances);
                            bboxTracker?.SetItemCount(placedInstances.Count);
                        }
                    }

                    // 4. STAGE 1 DELETE & STAGE 2 CLEANUP (Single post-placement transaction)
                    // Now that clusters are placed and committed, we can delete the individual sleeves
                    if (preparedData != null && preparedData.IndividualSleevesToDelete.Count > 0)
                    {
                        using (var delTracker = _performanceMonitor?.TrackOperation("Step 6: DELETE INDIVIDUAL SLEEVES"))
                        {
                            ExecuteStage1PreDelete(doc, preparedData.IndividualSleevesToDelete);
                            delTracker?.SetItemCount(preparedData.IndividualSleevesToDelete.Count);
                        }
                    }

                    // 4. Wait for background DB task to complete (needed before cleanup reads DB)
                    if (dbTask != null)
                    {
                        using (var waitTracker = _performanceMonitor?.TrackOperation("Step 6: WAIT FOR CLUSTER DB PERSISTENCE (Parallel)"))
                        {
                            dbTask.Wait();
                            waitTracker?.SetItemCount(clusterPersistenceData?.PlacedIdInts?.Count ?? 0);
                        }
                    }


                }
                catch (Exception ex)
                {
                     SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ BATCH TRANS ERROR: {ex.Message}\n{ex.StackTrace}\n");
                     failedCount = pendingClusters.Count; 
                     if (manageTransaction && mainTransaction != null && mainTransaction.GetStatus() == TransactionStatus.Started) 
                     {
                        mainTransaction.RollBack();
                     }
                     // If nested transaction, we rethrow so caller knows to rollback
                     throw;
                }
                finally
                {
                    if (mainTransaction != null) mainTransaction.Dispose();
                }

                tracker?.SetItemCount(placedCount);
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH COMPLETED: Placed={placedCount}, Failed={failedCount}, TransactionManaged={manageTransaction}\n");
                return (placedCount, failedCount);
            }
        }





        private List<BatchClusterData> GetPendingClusters(Document doc, string batchId)
        {
            var list = new List<BatchClusterData>();
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: GetPendingClusters called with batchId={(batchId ?? "NULL")}\n");
            
            // ✅ PATH 1 FIX: Use existence checker to verify clusters actually exist in Revit
            var existenceChecker = new Services.Clustering.Existence.SleeveExistenceChecker(doc);
            
            using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
            {
                conn.Open();
                
                // ✅ DIAGNOSTIC: First check total count of clusters in ClusterSleeves table (legacy)
                using (var countCmd = conn.CreateCommand())
                {
                    countCmd.CommandText = "SELECT COUNT(*) FROM ClusterSleeves";
                    var rawTotal = countCmd.ExecuteScalar();
                    var totalCount = (rawTotal != null && rawTotal != DBNull.Value) ? Convert.ToInt32(rawTotal) : 0;
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Total clusters in ClusterSleeves (legacy): {totalCount}\n");
                }
                
                // ✅ DIAGNOSTIC: Check clusters needing placement count from ClusterSleeves
                using (var pendingCmd = conn.CreateCommand())
                {
                    pendingCmd.CommandText = @"SELECT COUNT(*) FROM ClusterSleeves 
                        WHERE ClusterInstanceId IS NOT NULL";
                    var rawPending = pendingCmd.ExecuteScalar();
                    var totalWithInstanceId = (rawPending != null && rawPending != DBNull.Value) ? Convert.ToInt32(rawPending) : 0;
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Clusters in ClusterSleeves with ClusterInstanceId: {totalWithInstanceId}\n");
                }
                
            using (var cmd = conn.CreateCommand())
            {
                // ✅ FIX: Query from ClusterSleeves_v2 table (New Batch V2)
                // PENDING clusters: (ClusterInstanceId IS NULL OR <= 0) - NEW clusters needing placement
                // RECOVERY clusters: ClusterInstanceId > 0 - check existence, re-place if deleted from Revit
                // Previously only queried ClusterInstanceId > 0, which excluded all new clusters!
                cmd.CommandText = @"
                    SELECT
                        cs.ClusterGUID,
                        cs.PlacementX, cs.PlacementY, cs.PlacementZ,
                        cs.ClusterWidth, cs.ClusterHeight, cs.ClusterDepth,
                        cs.RotationAngleRad,
                        cs.HostElementId, cs.HostType, cs.HostOrientation,
                        cs.FamilyName, cs.ConstituentZoneGuids,
                        cs.ClusterInstanceId,
                        cs.IsCrossCategory
                    FROM ClusterSleeves_v2 cs
                    WHERE cs.Status IN ('Pending', 'Placed')";
                
                // Add optional batch filtering if provided
                if (!string.IsNullOrEmpty(batchId))
                {
                    cmd.CommandText += " AND cs.ClusterBatchId = @BatchId";
                    cmd.Parameters.AddWithValue("@BatchId", batchId);
                }

                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 Batch V2: Querying ClusterSleeves_v2 for Pending/Placed clusters (new + recovery)\n");
                
                int totalInTable = 0;
                    int needsPlacementCount = 0;
                    int alreadyExistsCount = 0;
                    
                    int batchLogIntervalClusters = 20;
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            totalInTable++;
                            var clusterGuid = reader["ClusterGUID"]?.ToString() ?? "";
                            var clusterInstanceId = (int)ReadLong(reader, "ClusterInstanceId"); // ✅ FIX: Use ReadLong to handle DBNull safely
                            
                            if (clusterInstanceId <= 0)
                            {
                                // ✅ NEW CLUSTER: Never placed - needs placement
                                needsPlacementCount++;
                                // ✅ BATCH LOGGING: Log every 20th only
                                if (needsPlacementCount % batchLogIntervalClusters == 1 || needsPlacementCount <= 1)
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 NEW: Cluster {clusterGuid} - needs placement (count={needsPlacementCount})\n");
                                
                                list.Add(new BatchClusterData
                                {
                                    ClusterGUID = clusterGuid,
                                    PlacementX = ReadDouble(reader, "PlacementX"),
                                    PlacementY = ReadDouble(reader, "PlacementY"),
                                    PlacementZ = ReadDouble(reader, "PlacementZ"),
                                    ClusterWidth = ReadDouble(reader, "ClusterWidth"),
                                    ClusterHeight = ReadDouble(reader, "ClusterHeight"),
                                    ClusterDepth = ReadDouble(reader, "ClusterDepth"),
                                    RotationAngleRad = ReadDouble(reader, "RotationAngleRad"),
                                    HostElementId = ReadLong(reader, "HostElementId"),
                                    HostType = reader["HostType"]?.ToString(),
                                    HostOrientation = reader["HostOrientation"]?.ToString(),
                                    FamilyName = reader["FamilyName"]?.ToString() ?? "",
                                    ConstituentZoneGuids = reader["ConstituentZoneGuids"]?.ToString() ?? "",
                                    IsCrossCategory = ReadLong(reader, "IsCrossCategory") == 1
                                });
                            }
                            else
                            {
                                // ✅ RECOVERY: ClusterInstanceId > 0 - check if element still exists in Revit
                                bool existsInRevit = existenceChecker.ClusterSleeveExists(clusterInstanceId);
                                if (!existsInRevit)
                                {
                                    needsPlacementCount++;
                                    // ✅ BATCH LOGGING: Log every 20th only
                                    if (needsPlacementCount % batchLogIntervalClusters == 1 || needsPlacementCount <= 1)
                                        SafeFileLogger.SafeAppendText("batch_v2.log",
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 PATH 1 RECOVERY: Cluster {clusterGuid} [ID={clusterInstanceId}] deleted from Revit - will re-place (count={needsPlacementCount})\n");

                                    list.Add(new BatchClusterData
                                    {
                                        ClusterGUID = clusterGuid,
                                        PlacementX = ReadDouble(reader, "PlacementX"),
                                        PlacementY = ReadDouble(reader, "PlacementY"),
                                        PlacementZ = ReadDouble(reader, "PlacementZ"),
                                        ClusterWidth = ReadDouble(reader, "ClusterWidth"),
                                        ClusterHeight = ReadDouble(reader, "ClusterHeight"),
                                        ClusterDepth = ReadDouble(reader, "ClusterDepth"),
                                        RotationAngleRad = ReadDouble(reader, "RotationAngleRad"),
                                        HostElementId = ReadLong(reader, "HostElementId"),
                                        HostType = reader["HostType"]?.ToString(),
                                        HostOrientation = reader["HostOrientation"]?.ToString(),
                                        FamilyName = reader["FamilyName"]?.ToString() ?? "",
                                        ConstituentZoneGuids = reader["ConstituentZoneGuids"]?.ToString() ?? "",
                                        IsCrossCategory = ReadLong(reader, "IsCrossCategory") == 1
                                    });
                                }
                                else
                                {
                                    alreadyExistsCount++;
                                    // ✅ BATCH LOGGING: Log every 20th only
                                    if (alreadyExistsCount % batchLogIntervalClusters == 1 || alreadyExistsCount <= 1)
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1: Cluster {clusterGuid} [ID={clusterInstanceId}] exists in Revit - skipping (count={alreadyExistsCount})\n");
                                }
                            }
                        }
                    }
                    
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Cluster breakdown from ClusterSleeves - Total={totalInTable}, NeedsPlacement={needsPlacementCount}, AlreadyExists={alreadyExistsCount}\n");
                }
            }
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: GetPendingClusters returning {list.Count} clusters needing placement\n");
            return list;
        }

        private const int StatusBatchChunkSize = 100;

        private void UpdateStatusBatch(Dictionary<string, int> guidToIdMap, string status, SQLiteTransaction? externalTrans = null)
        {
            if (guidToIdMap == null || guidToIdMap.Count == 0) return;
            var list = guidToIdMap.ToList();

            if (externalTrans != null)
            {
                ExecuteUpdateStatusBatch(list, status, externalTrans.Connection, externalTrans);
            }
            else
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        ExecuteUpdateStatusBatch(list, status, conn, trans);
                        trans.Commit();
                    }
                }
            }
        }

        private void ExecuteUpdateStatusBatch(List<KeyValuePair<string, int>> list, string status, SQLiteConnection conn, SQLiteTransaction trans)
        {
            for (int offset = 0; offset < list.Count; offset += StatusBatchChunkSize)
            {
                var chunk = list.Skip(offset).Take(StatusBatchChunkSize).ToList();
                if (chunk.Count == 0) continue;

                // Batch UPDATE ClusterSleeves_v2: one query per chunk (avoids N round-trips)
                var caseId = string.Join(" ", chunk.Select((kvp, i) => $"WHEN @g{i} THEN @id{i}"));
                var inList = string.Join(",", chunk.Select((_, i) => $"@g{i}"));
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = trans;
                    cmd.CommandText = $@"
                        UPDATE ClusterSleeves_v2 
                        SET Status = @status, 
                            ClusterInstanceId = CASE ClusterGUID {caseId} END,
                            PlacedAt = CURRENT_TIMESTAMP
                        WHERE ClusterGUID IN ({inList})";
                    cmd.Parameters.AddWithValue("@status", status);
                    for (int i = 0; i < chunk.Count; i++)
                    {
                        cmd.Parameters.AddWithValue($"@g{i}", chunk[i].Key);
                        cmd.Parameters.AddWithValue($"@id{i}", chunk[i].Value);
                    }
                    cmd.ExecuteNonQuery();
                }

                // ✅ PERF: Legacy ClusterSleeves sync removed — all readers now use ClusterSleeves_v2
            }
        }

        private void UpdateStatus(string guid, string status, string msg = null, int instanceId = -1)
        {
             using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClusterSleeves_v2 
                        SET Status = @status, 
                            ValidationMessage = @msg, 
                            ClusterInstanceId = @id,
                            PlacedAt = CURRENT_TIMESTAMP
                        WHERE ClusterGUID = @guid";
                    cmd.Parameters.AddWithValue("@status", status);
                    cmd.Parameters.AddWithValue("@msg", msg ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@id", instanceId);
                    cmd.Parameters.AddWithValue("@guid", guid);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void PerformSwapDeletion(Document doc, BatchClusterData cluster, int clusterElementId, bool skipPlacementPointUpdate = false, bool skipDeletion = false)
        {
            // 🔥 DIAGNOSTIC: Log entry and transaction state
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🟡 PerformSwapDeletion ENTRY: ClusterGUID={cluster.ClusterGUID}, doc.IsModifiable={doc.IsModifiable}\n");
            
            if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) return;

            // ✅ FIX: Declare local lists used for batch operations
            var toDeleteIds = new List<ElementId>();
            var databaseUpdates = new List<(System.Guid ClashZoneId, bool IsResolved, bool? IsClusterResolved, bool? IsCombinedResolved, long SleeveInstanceId, long ClusterInstanceId, bool? IsClusteredFlag, bool? MarkedForClusterProcess, long AfterClusterSleeveId, bool IsClustered, double SleeveWidth, double SleeveHeight, double SleeveDiameter, double SleeveDepth, string SleeveFamilyName, double? ActivePlacementX, double? ActivePlacementY, double? ActivePlacementZ, double? BBoxMinX, double? BBoxMinY, double? BBoxMinZ, double? BBoxMaxX, double? BBoxMaxY, double? BBoxMaxZ)>();

            // ✅ CRITICAL FIX: Use TryParse instead of Parse to handle invalid GUIDs gracefully
            var guids = new List<Guid>();
            var guidStrings = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var guidStr in guidStrings)
            {
                var trimmed = guidStr.Trim();
                if (Guid.TryParse(trimmed, out Guid parsedGuid))
                {
                    guids.Add(parsedGuid);
                }
                else
                {
                    // Log invalid GUID but don't crash
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ WARNING: Invalid GUID in ConstituentZoneGuids: '{trimmed}' for cluster {cluster.ClusterGUID}\n");
                }
            }

            if (!guids.Any()) 
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ WARNING: No valid GUIDs found in ConstituentZoneGuids for cluster {cluster.ClusterGUID}. ConstituentZoneGuids='{cluster.ConstituentZoneGuids}'\n");
                return;
            }

            var zones = _repository.GetClashZonesByGuids(guids);

            // ✅ CRITICAL FIX: Identify and delete OLD cluster data from ClusterSleeves table
            // If these zones were previously part of a cluster, that cluster is now being replaced/invalidated.
            // We must remove the old rows to prevent duplicates and stale data.
            var oldClusterInstanceIds = new List<long>();

            foreach (var z in zones)
            {
                if (z.ClusterInstanceId > 0 && z.ClusterInstanceId != clusterElementId)
                {
                    oldClusterInstanceIds.Add(z.ClusterInstanceId);
                }
                // ✅ FALLBACK: If ClusterInstanceId is missing (e.g. lost in DB update), try to find it in V2 table
                else if (z.ClusterInstanceId <= 0)
                {
                    // Use ClashZoneGuid (string) if available, otherwise use Id (Guid) converted to string
                    string guidToSearch = !string.IsNullOrEmpty(z.ClashZoneGuid) ? z.ClashZoneGuid : z.Id.ToString();
                    if (Guid.TryParse(guidToSearch, out Guid parsedGuid))
                    {
                        long recoveredId = GetParentClusterIdFromV2(parsedGuid);
                        if (recoveredId > 0 && recoveredId != clusterElementId)
                        {
                            oldClusterInstanceIds.Add(recoveredId);
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ♻️ RECOVERED ClusterInstanceId={recoveredId} for Zone {guidToSearch} from ClusterSleeves_v2\n");
                        }
                    }
                }
            }
            
            oldClusterInstanceIds = oldClusterInstanceIds.Distinct().ToList();

            if (oldClusterInstanceIds.Any())
            {
                DeleteOldClusterSleeves(oldClusterInstanceIds.Select(x => (int)x).ToList());
            }

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 PerformSwapDeletion: Retrieved {zones.Count} zones for cluster {cluster.ClusterGUID}\n");


            foreach (var z in zones)
            {
                // ✅ ROBUST PERSISTENCE: ALWAYS try to recover from snapshot first
                // This decouples "History" (AfterClusterSleeveId) from "Current State" (SleeveInstanceId)
                // If snapshot exists, it is the undeniable truth of what was there.
                long effectiveSleeveId = _repository.TryGetSleeveInstanceIdFromSnapshot(z.Id);
                
                // Fallback to current state if snapshot missing (e.g. first run)
                if (effectiveSleeveId <= 0)
                {
                    effectiveSleeveId = z.SleeveInstanceId;
                }
                else
                {
                    // If we recovered a valid ID, ensure local object reflects it for deletion logic
                    z.SleeveInstanceId = effectiveSleeveId;
                }

                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}]   - Zone {z.Id}: EffectiveSleeveId={effectiveSleeveId} (FromSnapshot={effectiveSleeveId != z.SleeveInstanceId}), ClusterInstanceId={z.ClusterInstanceId}\n");
                
                // Delete from Revit
                if (effectiveSleeveId > 0)
                {
                    toDeleteIds.Add(ElementIdCompat.FromValue(effectiveSleeveId));
                }

                // ✅ FIX: Update DB with FULL cluster data (flags + calculated columns)
                // ✅ FLAG MANAGEMENT: MarkedForClusterProcess should already be set to true during calculation phase
                // Only zones in multi-sleeve clusters should have this flag = true
                // We keep it as true here because we're placing a cluster (which means it was marked during calculation)
                // z.Id is already a Guid, so use it directly
                databaseUpdates.Add((
                    ClashZoneId: z.Id,
                    IsResolved: true,
                    IsClusterResolved: true,
                    IsCombinedResolved: false,
                    SleeveInstanceId: -1,
                    ClusterInstanceId: clusterElementId,
                    IsClusteredFlag: true,
                    MarkedForClusterProcess: true, // ✅ Keep true - this zone is in a cluster
                    AfterClusterSleeveId: effectiveSleeveId,
                    IsClustered: true, // Legacy flag
                    SleeveWidth: 0.0,
                    SleeveHeight: 0.0,
                    SleeveDiameter: 0.0,
                    SleeveDepth: 0.0,
                    SleeveFamilyName: null,
                    ActivePlacementX: (double?)null,
                    ActivePlacementY: (double?)null,
                    ActivePlacementZ: (double?)null,
                    BBoxMinX: (double?)null,
                    BBoxMinY: (double?)null,
                    BBoxMinZ: (double?)null,
                    BBoxMaxX: (double?)null,
                    BBoxMaxY: (double?)null,
                    BBoxMaxZ: (double?)null
                ));
            }

            // ✅ CRITICAL FIX: Skip deletion if already done in bulk (before cluster placement)
            if (toDeleteIds.Any() && !skipDeletion)
            {
                try
                {
                    // 🔥 DIAGNOSTIC: Log before doc.Delete
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🟡 ABOUT TO DELETE {toDeleteIds.Count} sleeves, doc.IsModifiable={doc.IsModifiable}\n");
                    
                    doc.Delete(toDeleteIds); // Bulk delete
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🗑️ DELETED {toDeleteIds.Count} individual sleeves for cluster {clusterElementId}\n");
                }
                catch (Exception delEx)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Warning: Failed to delete some sleeves: {delEx.Message}\n");
                }
            }
            else if (toDeleteIds.Any() && skipDeletion)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⏭️ SKIPPING deletion of {toDeleteIds.Count} sleeves (already deleted before cluster placement)\n");
            }

            // ✅ FIX 1: Update Flags in DB (existing functionality)
            _repository.BatchUpdateFlags(databaseUpdates);
            
            // ✅ FIX 2: Update ClashZones with calculated cluster dimensions
            // ✅ CRITICAL FIX: Only update placement point if skipPlacementPointUpdate is false
            // When false, we'll update it later after rotation and regeneration
            if (!skipPlacementPointUpdate)
            {
                // Use actual Revit instance if available for final accuracy
                var element = doc.GetElement(ElementIdCompat.FromValue(clusterElementId)) as FamilyInstance;
                UpdateClashZonesCalculatedColumns(guids, cluster, clusterElementId, element);
            }
            else
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⏭️ SKIPPING placement point update (will update after rotation/regeneration)\n");
            }
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATED {guids.Count} zones with cluster data (ClusterInstanceId={clusterElementId})\n");
        }

        /// <summary>
        /// ✅ RECOVERY: Find parent cluster ID from V2 table using ConstituentZoneGuids
        /// Needed when ClashZone.ClusterInstanceId is lost/reset but V2 record persists
        /// </summary>
        private long GetParentClusterIdFromV2(Guid zoneGuid)
        {
            try
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        // Check if this GUID exists in the CSV string of constituents
                        // We filter for ClusterInstanceId > 0 to ensure we get a PLACED cluster
                        cmd.CommandText = @"
                            SELECT ClusterInstanceId 
                            FROM ClusterSleeves_v2 
                            WHERE ConstituentZoneGuids LIKE @guidPattern 
                              AND ClusterInstanceId > 0
                            LIMIT 1";
                        
                        cmd.Parameters.AddWithValue("@guidPattern", $"%{zoneGuid}%");
                        
                        var result = cmd.ExecuteScalar();
                        if (result != null && long.TryParse(result.ToString(), out long id))
                        {
                            return id;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error in GetParentClusterIdFromV2: {ex.Message}\n");
            }
            return -1;
        }

        /// <summary>
        /// ✅ PHASE A BATCH: Single GetClashZonesByGuids for all clusters, single BatchUpdateFlags, single DeleteOldClusterSleeves.
        /// Builds ClusterSaveData from cluster data + CalculateCorners (no per-cluster PerformSwapDeletion).
        /// </summary>
        private (List<JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSaveData> clusterSaveDataList, Dictionary<string, int> placedGuidMap) PrepareClusterSaveDataBatch(
            Document doc,
            List<int> placedIdInts,
            List<BatchClusterData> clusterMap,
            Dictionary<int, (int comboId, int filterId, string category, string hostType, string hostOrientation)> comboIdMap,
            SQLiteTransaction? externalTrans = null,
            Dictionary<Guid, ClashZone> zoneCache = null)
        {
            var clusterSaveDataList = new List<JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSaveData>();
            var placedGuidMap = new Dictionary<string, int>();

            // 1. Collect all constituent GUIDs from all clusters
            var allGuids = new HashSet<Guid>();
            var guidsByClusterIndex = new Dictionary<int, List<Guid>>();
            for (int i = 0; i < placedIdInts.Count; i++)
            {
                var cluster = clusterMap[i];
                if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                var guids = new List<Guid>();
                foreach (var part in cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var trimmed = part.Trim();
                    if (Guid.TryParse(trimmed, out Guid g))
                    {
                        guids.Add(g);
                        allGuids.Add(g);
                    }
                }
                if (guids.Count > 0)
                    guidsByClusterIndex[i] = guids;
            }
            if (allGuids.Count == 0)
                return (clusterSaveDataList, placedGuidMap);

            // 2. Use pre-fetched zoneCache if available (avoids duplicate DB query)
            List<ClashZone> allZones;
            if (zoneCache != null && zoneCache.Count > 0)
            {
                allZones = allGuids.Where(g => zoneCache.ContainsKey(g)).Select(g => zoneCache[g]).ToList();
            }
            else
            {
                allZones = _repository.GetClashZonesByGuids(allGuids.ToList());
            }

            // 3. Zones per cluster (from pre-fetched allZones)
            var zonesByClusterIndex = new Dictionary<int, List<ClashZone>>();
            foreach (var kvp in guidsByClusterIndex)
            {
                int idx = kvp.Key;
                var guidsSet = kvp.Value.ToHashSet();
                zonesByClusterIndex[idx] = allZones.Where(z => guidsSet.Contains(z.Id)).ToList();
            }

            // 4. Collect all old cluster instance IDs (for DeleteOldClusterSleeves)
            var currentClusterIds = new HashSet<int>(placedIdInts);
            var allOldIds = new HashSet<int>();
            foreach (var z in allZones)
            {
                if (z.ClusterInstanceId > 0 && !currentClusterIds.Contains((int)z.ClusterInstanceId))
                    allOldIds.Add((int)z.ClusterInstanceId);
                else if (z.ClusterInstanceId <= 0)
                {
                    int recovered = (int)GetParentClusterIdFromV2(z.Id);
                    if (recovered > 0 && !currentClusterIds.Contains(recovered))
                        allOldIds.Add(recovered);
                }
            }
            if (allOldIds.Any())
                DeleteOldClusterSleeves(allOldIds.ToList(), externalTrans);

            // 5. Build all databaseUpdates and call BatchUpdateFlags once
            // ✅ PERF: Batch-fetch all SleeveInstanceIds from snapshots in 1 query (was N+1)
            var snapshotSleeveIdMap = new Dictionary<Guid, int>();
            if (allGuids.Count > 0)
            {
                try
                {
                    using (var snapshotConn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                    {
                        snapshotConn.Open();
                        using (var cmd = snapshotConn.CreateCommand())
                        {
                            var guidParams = allGuids.Select((g, i) => $"UPPER(@g{i})").ToList();
                            cmd.CommandText = $@"
                                SELECT ClashZoneGuid, SleeveInstanceId
                                FROM SleeveSnapshots
                                WHERE UPPER(REPLACE(REPLACE(ClashZoneGuid, '{{{{', ''), '}}}}', '')) IN ({string.Join(",", guidParams)})
                                  AND SleeveInstanceId > 0";
                            int idx = 0;
                            foreach (var g in allGuids)
                            {
                                cmd.Parameters.AddWithValue($"@g{idx}", g.ToString().ToUpperInvariant());
                                idx++;
                            }
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var guidStr = reader.GetString(0);
                                    var cleanGuid = guidStr.Replace("{", "").Replace("}", "").Trim();
                                    if (Guid.TryParse(cleanGuid, out Guid parsedGuid))
                                    {
                                        snapshotSleeveIdMap[parsedGuid] = reader.GetInt32(1);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Batch snapshot lookup failed (falling back): {ex.Message}\n");
                }
            }

            var allDatabaseUpdates = new List<(Guid ClashZoneId, bool IsResolved, bool? IsClusterResolved, bool? IsCombinedResolved, long SleeveInstanceId, long ClusterInstanceId, bool? IsClusteredFlag, bool? MarkedForClusterProcess, long AfterClusterSleeveId, bool IsClustered, double SleeveWidth, double SleeveHeight, double SleeveDiameter, double SleeveDepth, string SleeveFamilyName, double? ActivePlacementX, double? ActivePlacementY, double? ActivePlacementZ, double? BBoxMinX, double? BBoxMinY, double? BBoxMinZ, double? BBoxMaxX, double? BBoxMaxY, double? BBoxMaxZ)>();
            for (int i = 0; i < placedIdInts.Count; i++)
            {
                if (!zonesByClusterIndex.TryGetValue(i, out var zones) || zones.Count == 0) continue;
                long clusterInstanceId = placedIdInts[i];
                foreach (var z in zones)
                {
                    // ✅ PERF: Use batch-fetched snapshot map instead of per-zone DB query
                    long effectiveSleeveId = snapshotSleeveIdMap.TryGetValue(z.Id, out int snapId) ? snapId : -1;
                    if (effectiveSleeveId <= 0)
                        effectiveSleeveId = z.SleeveInstanceId;
                    allDatabaseUpdates.Add((z.Id, true, true, false, -1, clusterInstanceId, true, true, effectiveSleeveId, true, 0.0, 0.0, 0.0, 0.0, null, null, null, null, null, null, null, null, null, null));
                }
            }
            if (allDatabaseUpdates.Count > 0)
                _repository.BatchUpdateFlags(allDatabaseUpdates, externalTrans);

            // 6. Build ClusterSaveData from cluster data + corners from math (no GetElement for geometry)
            var cornerService = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 📊 DIAGNOSTIC: comboIdMap contains {comboIdMap.Count} entries for {placedIdInts.Count} placed clusters\n");

            for (int i = 0; i < placedIdInts.Count; i++)
            {
                var cluster = clusterMap[i];
                if (!comboIdMap.TryGetValue(i, out var combo)) continue;
                var (comboId, filterId, category, hostType, hostOrientation) = combo;
                int clusterInstanceId = placedIdInts[i];
                if (comboId <= 0 || filterId <= 0)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING cluster save for instance {clusterInstanceId} due to missing IDs: ComboId={comboId}, FilterId={filterId}, GUID={cluster.ClusterGUID}\n");
                    continue;
                }

                var placement = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                var calcCorners = cornerService.CalculateCorners(placement, cluster.ClusterWidth, cluster.ClusterHeight, cluster.RotationAngleRad);
                
                // ✅ OPTIMIZED: Skip slow diagnostic corner comparison unless explicitly enabled
                if (OptimizationFlags.EnablePlacementDiagnostics && doc != null)
                {
                    try
                    {
                        Element elem = doc.GetElement(ElementIdCompat.FromValue(clusterInstanceId));
                        if (elem is FamilyInstance instance)
                        {
                            var extractedCorners = cornerService.CalculateCornersFromInstance(instance, hostOrientation, category);
                            if (calcCorners.HasValue && extractedCorners.HasValue)
                            {
                                var c = calcCorners.Value;
                                var e = extractedCorners.Value;
                                
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 📐 CORNER COMPARISON (Cluster {clusterInstanceId}):\n" +
                                    $"   CALC: ({c.corner1.X:F3}, {c.corner1.Y:F3}, {c.corner1.Z:F3}) to ({c.corner4.X:F3}, {c.corner4.Y:F3}, {c.corner4.Z:F3})\n" +
                                    $"   REVT: ({e.corner1.X:F3}, {e.corner1.Y:F3}, {e.corner1.Z:F3}) to ({e.corner4.X:F3}, {e.corner4.Y:F3}, {e.corner4.Z:F3})\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Corner comparison failed for {clusterInstanceId}: {ex.Message}\n");
                    }
                }

                if (!calcCorners.HasValue) continue;
                var corners = calcCorners.Value;

                List<Guid> zoneGuids = null;
                if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                    zoneGuids = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => Guid.Parse(s.Trim())).ToList();
                if (zoneGuids == null || zoneGuids.Count == 0) continue;

                var saveData = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSaveData
                {
                    ClusterInstanceId = clusterInstanceId,
                    ComboId = comboId,
                    FilterId = filterId,
                    Category = category,
                    BoundingBoxMinX = 0,
                    BoundingBoxMinY = 0,
                    BoundingBoxMinZ = 0,
                    BoundingBoxMaxX = 0,
                    BoundingBoxMaxY = 0,
                    BoundingBoxMaxZ = 0,
                    ClusterWidth = cluster.ClusterWidth,
                    ClusterHeight = cluster.ClusterHeight,
                    ClusterDepth = cluster.ClusterDepth,
                    RotationAngleDeg = cluster.RotationAngleRad * (180.0 / Math.PI),
                    IsRotated = Math.Abs(cluster.RotationAngleRad) > 1e-6,
                    PlacementX = cluster.PlacementX,
                    PlacementY = cluster.PlacementY,
                    PlacementZ = cluster.PlacementZ,
                    HostType = hostType,
                    HostOrientation = hostOrientation,
                    ClashZoneIds = zoneGuids,
                    SleeveFamilyName = cluster.FamilyName ?? "",
                    Corner1X = corners.corner1.X,
                    Corner1Y = corners.corner1.Y,
                    Corner1Z = corners.corner1.Z,
                    Corner2X = corners.corner2.X,
                    Corner2Y = corners.corner2.Y,
                    Corner2Z = corners.corner2.Z,
                    Corner3X = corners.corner3.X,
                    Corner3Y = corners.corner3.Y,
                    Corner3Z = corners.corner3.Z,
                    Corner4X = corners.corner4.X,
                    Corner4Y = corners.corner4.Y,
                    Corner4Z = corners.corner4.Z,
                    IsCrossCategory = cluster.IsCrossCategory
                };
                clusterSaveDataList.Add(saveData);
                if (!string.IsNullOrEmpty(cluster.ClusterGUID))
                    placedGuidMap[cluster.ClusterGUID] = clusterInstanceId;
            }

            return (clusterSaveDataList, placedGuidMap);
        }

        /// <summary>
        /// ✅ FIX: Update ClashZones table with calculated cluster dimensions
        /// Populates: CalculatedSleeveWidth, Height, Depth, Rotation, FamilyName, PlacedAt
        /// </summary>
        private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance = null)
        {
            UpdateClashZonesCalculatedColumnsBatch(new List<(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance)>
            {
                (zoneGuids, cluster, clusterInstanceId, actualInstance)
            });
        }

        /// <summary>
        /// ✅ OPTIMIZATION: High-performance Batched update for ClashZones calculated columns.
        /// Uses a single database connection and transaction for all cluster updates.
        /// </summary>
        private void UpdateClashZonesCalculatedColumnsBatch(List<(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance)> updates, SQLiteTransaction? externalTrans = null)
        {
            if (updates == null || !updates.Any()) return;

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 💾 BATCH SQL: Updating calculated columns for {updates.Count} clusters (Single Transaction)\n");

            if (externalTrans != null)
            {
                ExecuteUpdateClashZonesCalculatedColumnsBatch(updates, externalTrans.Connection, externalTrans);
            }
            else
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            ExecuteUpdateClashZonesCalculatedColumnsBatch(updates, conn, transaction);
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ❌ BATCH UPDATE FAILED: {ex.Message}\n{ex.StackTrace}\n");
                            throw;
                        }
                    }
                }
            }
        }

        private void ExecuteUpdateClashZonesCalculatedColumnsBatch(List<(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance)> updates, SQLiteConnection conn, SQLiteTransaction transaction)
        {
            foreach (var update in updates)
            {
                var (zoneGuids, cluster, clusterInstanceId, actualInstance) = update;

                double finalWidth = cluster.ClusterWidth;
                double finalHeight = cluster.ClusterHeight;
                double finalDepth = cluster.ClusterDepth;
                double finalRotation = cluster.RotationAngleRad;

                // ✅ HIGH ACCURACY: Read from actual Revit element if provided
                if (actualInstance != null && actualInstance.IsValidObject)
                {
                    try
                    {
                        finalWidth = (actualInstance.LookupParameter("Width") ?? actualInstance.LookupParameter("Element Width"))?.AsDouble() ?? finalWidth;
                        finalHeight = (actualInstance.LookupParameter("Height") ?? actualInstance.LookupParameter("Element Height"))?.AsDouble() ?? finalHeight;
                        finalDepth = (actualInstance.LookupParameter("Depth") ?? actualInstance.LookupParameter("Element Depth") ?? actualInstance.LookupParameter("Wall Width"))?.AsDouble() ?? finalDepth;
                        
                        var loc = actualInstance.Location as LocationPoint;
                        if (loc != null)
                        {
                            finalRotation = loc.Rotation;
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("debug_db.log", $"    ⚠️ Error reading actual values: {ex.Message}\n");
                    }
                }
                
                if (zoneGuids == null || !zoneGuids.Any()) continue;

                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    
                    // Build WHERE clause with all GUIDs
                    var guidParams = new List<string>();
                    for (int i = 0; i < zoneGuids.Count; i++)
                    {
                        guidParams.Add($"@Guid{i}");
                        cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString());
                    }
                    
                    // ✅ Corners
                    var cornerService = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
                    var corners = actualInstance != null ? cornerService.CalculateCornersFromInstance(actualInstance, cluster.HostOrientation, cluster.HostType) : null;

                    cmd.CommandText = $@"
                        UPDATE ClashZones 
                        SET 
                            CalculatedSleeveWidth = @Width,
                            CalculatedSleeveHeight = @Height,
                            CalculatedSleeveDepth = @Depth,
                            CalculatedRotation = @Rotation,
                            CalculatedFamilyName = @FamilyName,
                            SleeveCorner1X = @C1X, SleeveCorner1Y = @C1Y, SleeveCorner1Z = @C1Z,
                            SleeveCorner2X = @C2X, SleeveCorner2Y = @C2Y, SleeveCorner2Z = @C2Z,
                            SleeveCorner3X = @C3X, SleeveCorner3Y = @C3Y, SleeveCorner3Z = @C3Z,
                            SleeveCorner4X = @C4X, SleeveCorner4Y = @C4Y, SleeveCorner4Z = @C4Z,
                            PlacedAt = CURRENT_TIMESTAMP,
                            PlacementStatus = 'Placed',
                            SleeveInstanceId = -1,
                            ClusterInstanceId = @ClusterInstanceId,
                            IsClusterResolvedFlag = 1,
                            SleeveState = 2,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE UPPER(ClashZoneGuid) IN ({string.Join(", ", guidParams.Select((_, i) => $"UPPER(@Guid{i})"))})";
                    
                    cmd.Parameters.AddWithValue("@Width", finalWidth);
                    cmd.Parameters.AddWithValue("@Height", finalHeight);
                    cmd.Parameters.AddWithValue("@Depth", finalDepth);
                    cmd.Parameters.AddWithValue("@Rotation", finalRotation);
                    cmd.Parameters.AddWithValue("@FamilyName", cluster.FamilyName ?? "");
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                    
                    cmd.Parameters.AddWithValue("@C1X", corners?.corner1?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C1Y", corners?.corner1?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C1Z", corners?.corner1?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C2X", corners?.corner2?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C2Y", corners?.corner2?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C2Z", corners?.corner2?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C3X", corners?.corner3?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C3Y", corners?.corner3?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C3Z", corners?.corner3?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C4X", corners?.corner4?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C4Y", corners?.corner4?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C4Z", corners?.corner4?.Z ?? 0);
                    
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// ✅ PERF: Background-thread-safe overload using pre-read Revit params (no FamilyInstance access).
        /// Uses math-based corner calculation instead of CalculateCornersFromInstance.
        /// </summary>
        private void ExecuteUpdateClashZonesCalculatedColumnsBatchPreRead(
            List<(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, double finalWidth, double finalHeight, double finalDepth, double finalRotation)> updates,
            SQLiteConnection conn, SQLiteTransaction transaction)
        {
            var cornerService = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
            foreach (var update in updates)
            {
                var (zoneGuids, cluster, clusterInstanceId, finalWidth, finalHeight, finalDepth, finalRotation) = update;
                if (zoneGuids == null || !zoneGuids.Any()) continue;

                var placement = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                var corners = cornerService.CalculateCorners(placement, finalWidth, finalHeight, finalRotation);

                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    var guidParams = new List<string>();
                    for (int i = 0; i < zoneGuids.Count; i++)
                    {
                        guidParams.Add($"@Guid{i}");
                        cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString());
                    }

                    cmd.CommandText = $@"
                        UPDATE ClashZones
                        SET
                            CalculatedSleeveWidth = @Width,
                            CalculatedSleeveHeight = @Height,
                            CalculatedSleeveDepth = @Depth,
                            CalculatedRotation = @Rotation,
                            CalculatedFamilyName = @FamilyName,
                            SleeveCorner1X = @C1X, SleeveCorner1Y = @C1Y, SleeveCorner1Z = @C1Z,
                            SleeveCorner2X = @C2X, SleeveCorner2Y = @C2Y, SleeveCorner2Z = @C2Z,
                            SleeveCorner3X = @C3X, SleeveCorner3Y = @C3Y, SleeveCorner3Z = @C3Z,
                            SleeveCorner4X = @C4X, SleeveCorner4Y = @C4Y, SleeveCorner4Z = @C4Z,
                            PlacedAt = CURRENT_TIMESTAMP,
                            PlacementStatus = 'Placed',
                            SleeveInstanceId = -1,
                            ClusterInstanceId = @ClusterInstanceId,
                            IsClusterResolvedFlag = 1,
                            SleeveState = 2,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE UPPER(ClashZoneGuid) IN ({string.Join(", ", guidParams.Select((_, i) => $"UPPER(@Guid{i})"))})";

                    cmd.Parameters.AddWithValue("@Width", finalWidth);
                    cmd.Parameters.AddWithValue("@Height", finalHeight);
                    cmd.Parameters.AddWithValue("@Depth", finalDepth);
                    cmd.Parameters.AddWithValue("@Rotation", finalRotation);
                    cmd.Parameters.AddWithValue("@FamilyName", cluster.FamilyName ?? "");
                    cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                    cmd.Parameters.AddWithValue("@C1X", corners?.corner1?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C1Y", corners?.corner1?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C1Z", corners?.corner1?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C2X", corners?.corner2?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C2Y", corners?.corner2?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C2Z", corners?.corner2?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C3X", corners?.corner3?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C3Y", corners?.corner3?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C3Z", corners?.corner3?.Z ?? 0);
                    cmd.Parameters.AddWithValue("@C4X", corners?.corner4?.X ?? 0);
                    cmd.Parameters.AddWithValue("@C4Y", corners?.corner4?.Y ?? 0);
                    cmd.Parameters.AddWithValue("@C4Z", corners?.corner4?.Z ?? 0);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// ✅ FIX: Save cluster to ClusterSleeves (legacy) table for PATH 1 compatibility
        /// This enables cluster replay from database
        /// </summary>
        private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, long clusterInstanceId, FamilyInstance actualInstance = null)
        {
            // ✅ DIAGNOSTIC: Log method entry
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"\n[{DateTime.Now:HH:mm:ss}] 🔴 SaveToClusterSleevesLegacy CALLED!\n" +
                $"    📊 TABLE: Saving to ClusterSleeves table (legacy)\n" +
                $"    ClusterInstanceId: {clusterInstanceId}\n" +
                $"    ConstituentZoneGuids: {cluster.ConstituentZoneGuids}\n");
            
            try
            {
                // Get ComboId and FilterId from first constituent zone
                int comboId = -1;
                int filterId = -1;
                string category = "";
                string hostType = "";
                string hostOrientation = "";

                if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                {
                    var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                    if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid firstGuid))
                    {
                        var zones = _repository.GetClashZonesByGuids(new List<Guid> { firstGuid });
                        var firstZone = zones.FirstOrDefault();
                        if (firstZone != null)
                        {
                            comboId = (int)firstZone.ComboId;
                            category = firstZone.MepElementCategory ?? "";
                            hostType = firstZone.StructuralElementType ?? "";
                            hostOrientation = firstZone.HostOrientation ?? "";

                            // Parse constituent zone GUIDs
                            var zoneGuids = cluster.ConstituentZoneGuids
                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(g => Guid.Parse(g.Trim()))
                                .ToList();

                            // Calculate bounding box (approximate from placement + dimensions)
                            double halfWidth = cluster.ClusterWidth / 2.0;
                            double halfHeight = cluster.ClusterHeight / 2.0;
                            double halfDepth = cluster.ClusterDepth / 2.0;

                            // Save to database using SleeveDbContext
                            using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(doc))
                            {
                                // ✅ FIX: FilterId is in FileCombos table, not ClashZone model
                                if (comboId > 0 && filterId <= 0)
                                {
                                    using (var cmd = dbContext.Connection.CreateCommand())
                                    {
                                        cmd.CommandText = "SELECT FilterId FROM FileCombos WHERE ComboId = @ComboId LIMIT 1";
                                        cmd.Parameters.AddWithValue("@ComboId", comboId);
                                        var result = cmd.ExecuteScalar();
                                        if (result != null && result != DBNull.Value)
                                        {
                                            filterId = Convert.ToInt32(result);
                                        }
                                    }
                                }

                                // ✅ DIAGNOSTIC: Log IDs before validation
                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     ComboId: {comboId}, FilterId: {filterId}\n");

                                if (comboId <= 0 || filterId <= 0)
                                {
                                    string error = $"Cannot save to ClusterSleeves: Missing ComboId ({comboId}) or FilterId ({filterId})";
                                    SafeFileLogger.SafeAppendText("batch_v2.log",
                                        $"[{DateTime.Now:HH:mm:ss}]     ❌ EARLY RETURN: Invalid IDs\n");
                                    
                                    // ✅ CRITICAL FIX: Throw exception for missing IDs - database save is MANDATORY
                                    throw new InvalidOperationException(error);
                                }

                                var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveRepository(dbContext, msg => SafeFileLogger.SafeAppendText("batch_v2.log", msg));

                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     ✅ IDs valid, about to call SaveClusterSleeve\n");

                                repo.SaveClusterSleeve(
                                    clusterInstanceId,
                                    comboId,
                                    filterId,
                                    category,
                                    cluster.FamilyName,
                                    actualInstance,
                                    rotationAngleRadOverride: cluster.RotationAngleRad
                                );

                                // ✅ DIAGNOSTIC: Log after SaveClusterSleeve returns
                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     ✅ SaveClusterSleeve RETURNED\n");

                                // 🔍 LOG BLOCK 4: SaveToClusterSleevesLegacy Success
                                SafeFileLogger.SafeAppendText("debug_db.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 💾 TABLE: SAVED to ClusterSleeves table (legacy): ClusterInstanceId={clusterInstanceId}, ComboId={comboId}\n");

                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ TABLE: SAVED to ClusterSleeves table (legacy): ClusterInstanceId={clusterInstanceId}, ComboId={comboId}\n");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Failed to save to ClusterSleeves (legacy): {ex.Message}\n");
                
                // ✅ CRITICAL FIX: Rethrow exception - database save failure is FATAL
                throw new InvalidOperationException($"Failed to save cluster {cluster.ClusterGUID} to legacy database: {ex.Message}", ex);
            }
        }

        private bool PlaceSingleCluster(
            Document doc, 
            BatchClusterData cluster, 
            Dictionary<string, FamilySymbol> symbolCache = null, 
            Dictionary<Guid, ClashZone> zoneCache = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker parentTracker = null,
            List<FamilyInstance> placedInstances = null,
            bool logDetail = true,
            Dictionary<string, ElementId> levelMap = null)
        {
            // ✅ BATCH LOGGING: Only log every 20th cluster in sequential path
            if (logDetail)
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔵 PlaceSingleCluster ENTRY: ClusterGUID={cluster.ClusterGUID}, doc.IsModifiable={doc.IsModifiable}\n");
            
            try
            {
                // A. Load/Activate Symbol
                FamilySymbol symbol = null;
                if (symbolCache != null && symbolCache.TryGetValue(cluster.FamilyName, out var cachedSymbol))
                {
                    symbol = cachedSymbol;
                }
                else
                {
                    symbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(x => x.Name == cluster.FamilyName || x.Family.Name == cluster.FamilyName);
                }

                if (symbol == null)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ Missing Family Symbol: {cluster.FamilyName}\n");
                    UpdateStatus(cluster.ClusterGUID, "Failed", "Missing Family Symbol");
                    return false;
                }

                if (!symbol.IsActive)
                {
                    using (parentTracker != null ? parentTracker.TrackSubOperation("Activate Cluster Symbol") : _performanceMonitor?.TrackOperation("Activate Cluster Symbol"))
                    {
                        symbol.Activate();
                    }
                }

                // B. Determine Level (Not used for placement, but kept for cache/context if needed)
                Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

                // C. Create Instance
                XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                
                // ✅ BATCH LOGGING: Only log every 20th cluster
                if (logDetail)
                    SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                        $"\n[{DateTime.Now:HH:mm:ss}] 🎯 BEFORE PLACEMENT - ClusterGUID={cluster.ClusterGUID}:\n" +
                        $"    THEORETICAL PLACEMENT POINT: X={cluster.PlacementX:F6}, Y={cluster.PlacementY:F6}, Z={cluster.PlacementZ:F6}\n" +
                        $"    HostType: {cluster.HostType}, HostOrientation: {cluster.HostOrientation}\n" +
                        $"    Rotation: {cluster.RotationAngleRad:F6} rad ({cluster.RotationAngleRad * 180.0 / Math.PI:F2}°)\n");
                
                // ✅ DUPLICATE DETECTION: Check if cluster already exists at this location (when DB is not cleared)
                // This prevents the "duplicate found do you want to proceed" dialog
                double tolerance = 0.1; // 0.1 feet tolerance for location matching
                var existingClusters = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => 
                        fi.Symbol?.Family?.Name == cluster.FamilyName &&
                        fi.Location is LocationPoint lp &&
                        Math.Abs(lp.Point.X - location.X) < tolerance &&
                        Math.Abs(lp.Point.Y - location.Y) < tolerance &&
                        Math.Abs(lp.Point.Z - location.Z) < tolerance)
                    .ToList();
                
                if (existingClusters.Any())
                {
                    // Cluster already exists at this location - skip placement
                    var existingId = existingClusters.First().Id.GetIntegerValue();
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ DUPLICATE DETECTED: Cluster already exists at ({location.X:F3}, {location.Y:F3}, {location.Z:F3}), " +
                        $"ExistingId={existingId}, Skipping placement for ClusterGUID={cluster.ClusterGUID}\n");
                    
                    // Update status to indicate this cluster was skipped due to duplicate
                    UpdateStatus(cluster.ClusterGUID, "Skipped", "Duplicate location", existingId);
                    
                    // Still update flags and save to legacy table using existing cluster
                    try
                    {
                        PerformSwapDeletion(doc, cluster, existingId);
                        SaveToClusterSleevesLegacy(doc, cluster, existingId, existingClusters.First());
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to update flags for duplicate cluster: {ex.Message}\n");
                    }
                    
                    return false; // Skip this placement
                }
                
                FamilyInstance instance = null;
                Autodesk.Revit.DB.Structure.StructuralType structuralType = Autodesk.Revit.DB.Structure.StructuralType.NonStructural;

                using (parentTracker != null ? parentTracker.TrackSubOperation("Revit Create Cluster Instance") : _performanceMonitor?.TrackOperation("Revit Create Cluster Instance"))
                {
                    // ✅ PERF: Use explicit level placement to skip spatial search
                    ElementId levelId = null;
                    if (levelMap != null && !string.IsNullOrEmpty(cluster.MepElementLevelName) && levelMap.TryGetValue(cluster.MepElementLevelName, out var id))
                    {
                        levelId = id;
                    }

                    if (levelId != null)
                    {
                        var targetLevel = doc.GetElement(levelId) as Level;
                        if (targetLevel != null)
                        {
                            instance = doc.Create.NewFamilyInstance(location, symbol, targetLevel, structuralType);
                        }
                        else
                        {
                            instance = doc.Create.NewFamilyInstance(location, symbol, structuralType);
                        }
                    }
                    else
                    {
                        instance = doc.Create.NewFamilyInstance(location, symbol, structuralType);
                    }
                }

                if (instance != null)
                {
                    int clusterInstanceId = instance.Id.GetIntegerValue();

                    // D. Set Parameters using SleeveParameterService (CODE REUSE)
                    // 1. Prepare Template Zone (Proxy)
                    ClashZone templateZone = null;
                    if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                    {
                        var firstGuidStr = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim();
                        if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid g))
                        {
                            // Try cache first
                            if (zoneCache != null && zoneCache.TryGetValue(g, out var cachedZone))
                            {
                                templateZone = cachedZone;
                            }
                            else
                            {
                                templateZone = _repository.GetClashZonesByGuids(new List<Guid> { g }).FirstOrDefault();
                            }
                        }
                    }

                    if (templateZone == null)
                    {
                        templateZone = new ClashZone();
                    }
                    
                    // 3. Override Template with Cluster Geometry
                    templateZone.CalculatedSleeveWidth = cluster.ClusterWidth;
                    templateZone.CalculatedSleeveHeight = cluster.ClusterHeight;
                    templateZone.CalculatedSleeveDepth = cluster.ClusterDepth;
                    templateZone.CalculatedRotation = cluster.RotationAngleRad;
                    
                    try
                    {
                        using (parentTracker != null ? parentTracker.TrackSubOperation("Swap Individual Sleeves") : _performanceMonitor?.TrackOperation("Swap Individual Sleeves"))
                        {
                            // F. SWAP LOGIC (Delete individual sleeves, but DON'T save placement point yet)
                            PerformSwapDeletion(doc, cluster, clusterInstanceId, skipPlacementPointUpdate: true);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ SWAP OPERATION FAILED for Cluster {cluster.ClusterGUID}: {ex.Message}\n");
                        throw; 
                    }

                    // 4. Call Service
                    using (parentTracker != null ? parentTracker.TrackSubOperation("Set Cluster Instance Parameters") : _performanceMonitor?.TrackOperation("Set Cluster Instance Parameters"))
                    {
                        bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || cluster.FamilyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;
                        try
                        {
                            _parameterService.SetSleeveParameters(
                                instance, 
                                cluster.ClusterWidth, 
                                cluster.ClusterHeight, 
                                isCircular ? cluster.ClusterWidth : 0, 
                                isCircular, 
                                templateZone,
                                isCircular ? (double?)null : cluster.ClusterDepth);
                            
                            // ✅ NEW: Set Cluster Sleeve Instance ID (Explicit Parameter)
                            _parameterService.SetClusterSleeveInstanceId(instance, clusterInstanceId);
                            // ✅ Cluster sleeves: SleeveInstanceId parameter on Revit element = -1 (not an individual)
                            _parameterService.SetSleeveInstanceId(instance, -1);
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Parameter setting failed for cluster {clusterInstanceId}: {ex.Message}\n");
                        }
                    }

                    // E. Physical Rotation
                    // ✅ PERF: Skip rotation if family is pre-rotated (_X variant)
                    bool isPreRotatedClusterFamily = (cluster.FamilyName ?? "").EndsWith("_X", StringComparison.OrdinalIgnoreCase);
                    if (Math.Abs(cluster.RotationAngleRad) > 1e-6 && !isPreRotatedClusterFamily)
                    {
                        using (parentTracker != null ? parentTracker.TrackSubOperation("Revit Rotate Cluster") : _performanceMonitor?.TrackOperation("Revit Rotate Cluster"))
                        {
                            Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
                        }
                    }
                    
                    // ✅ COMMENTED OUT: Revit API commit auto-invokes regeneration. Roll back if needed.
                    // doc.Regenerate();
                    
                    // ✅ CRITICAL FIX: Refresh instance reference after regeneration
#if REVIT2023
                    instance = doc.GetElement(new ElementId(clusterInstanceId)) as FamilyInstance;
#else
                    instance = doc.GetElement(new ElementId((long)clusterInstanceId)) as FamilyInstance;
#endif
                    
                    // ✅ BATCH LOGGING: Only log every 20th cluster
                    if (logDetail && instance != null && instance.IsValidObject)
                    {
                        var finalLoc = instance.Location as LocationPoint;
                        if (finalLoc != null)
                        {
                            double deltaX = finalLoc.Point.X - cluster.PlacementX;
                            double deltaY = finalLoc.Point.Y - cluster.PlacementY;
                            double deltaZ = finalLoc.Point.Z - cluster.PlacementZ;
                            double totalDisplacement = Math.Sqrt(deltaX * deltaX + deltaY * deltaY + deltaZ * deltaZ);
                            
                            SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                $"\n[{DateTime.Now:HH:mm:ss}] ✅ AFTER ALL TRANSFORMATIONS - ClusterInstanceId={clusterInstanceId}:\n" +
                                $"    THEORETICAL (original): X={cluster.PlacementX:F6}, Y={cluster.PlacementY:F6}, Z={cluster.PlacementZ:F6}\n" +
                                $"    ACTUAL (from Revit):    X={finalLoc.Point.X:F6}, Y={finalLoc.Point.Y:F6}, Z={finalLoc.Point.Z:F6}\n" +
                                $"    DISPLACEMENT: ΔX={deltaX * 304.8:F2}mm, ΔY={deltaY * 304.8:F2}mm, ΔZ={deltaZ * 304.8:F2}mm\n" +
                                $"    TOTAL DISPLACEMENT: {totalDisplacement * 304.8:F2}mm\n");
                            
                            if (totalDisplacement > 0.01) // More than 1cm
                            {
                                SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                    $"    ⚠️ WARNING: Significant displacement detected! (>1cm)\n");
                            }
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                $"\n[{DateTime.Now:HH:mm:ss}] ⚠️ AFTER TRANSFORMATIONS - ClusterInstanceId={clusterInstanceId}: Location is not LocationPoint (type: {instance.Location?.GetType().Name ?? "NULL"})\n");
                        }
                    }
                    else if (logDetail && (instance == null || !instance.IsValidObject))
                    {
                        SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                            $"\n[{DateTime.Now:HH:mm:ss}] ⚠️ AFTER TRANSFORMATIONS - ClusterInstanceId={clusterInstanceId}: Instance is NULL or invalid\n");
                    }
                    
                    // ✅ CRITICAL FIX: NOW save placement point AFTER all transformations (rotation, parameters, regeneration)
                    try
                    {
                        using (parentTracker != null ? parentTracker.TrackSubOperation("Save Placement Point to DB") : _performanceMonitor?.TrackOperation("Save Placement Point to DB"))
                        {
                            // G. Update ClusterSleeves_v2 Status
                            UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
                            // H. Save to ClusterSleeves (legacy) table
                            SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId, instance);
                            // I. Update ClashZones with FINAL placement point (after rotation and regeneration)
                            if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                            {
                                var guids = cluster.ConstituentZoneGuids
                                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(g => g.Trim())
                                    .Where(g => !string.IsNullOrEmpty(g) && Guid.TryParse(g, out _))
                                    .Select(g => Guid.Parse(g))
                                    .ToList();
                                if (guids.Any())
                                {
                                    UpdateClashZonesCalculatedColumns(guids, cluster, clusterInstanceId, instance);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ DB OPERATION FAILED for Cluster {cluster.ClusterGUID}: {ex.Message}\n");
                        throw; 
                    }

                    // ✅ Track placed instance for 2nd cleanup pass
                    placedInstances?.Add(instance);
                    
                    return true;
                }
                else
                {
                    UpdateStatus(cluster.ClusterGUID, "Failed", "Creation returned null");
                    return false;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ Placement Exception: {ex.Message}\n{ex.StackTrace}\n");
                UpdateStatus(cluster.ClusterGUID, "Failed", ex.Message);
                return false;
            }
        }



        /// <summary>
        /// ✅ CRITICAL FIX: Deletes old cluster records from BOTH ClusterSleeves (legacy) AND ClusterSleeves_v2 tables.
        /// This prevents stale placement points from being used when clusters are recalculated.
        /// </summary>
        private void DeleteOldClusterSleeves(List<int> clusterInstanceIds, SQLiteTransaction? externalTrans = null)
        {
            if (clusterInstanceIds == null || !clusterInstanceIds.Any()) return;

            if (externalTrans != null)
            {
                ExecuteDeleteOldClusterSleeves(clusterInstanceIds, externalTrans.Connection, externalTrans);
            }
            else
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try 
                        {
                            ExecuteDeleteOldClusterSleeves(clusterInstanceIds, conn, transaction);
                            transaction.Commit();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            SafeFileLogger.SafeAppendText("placement_errors.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to delete old clusters: {ex.Message}\n");
                        }
                    }
                }
            }
        }

        private void ExecuteDeleteOldClusterSleeves(List<int> clusterInstanceIds, SQLiteConnection conn, SQLiteTransaction transaction)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                var ids = string.Join(",", clusterInstanceIds);
                
                // ✅ FIX: Delete from BOTH tables to prevent stale placement points
                // 1. Delete from legacy ClusterSleeves table
                cmd.CommandText = $"DELETE FROM ClusterSleeves WHERE ClusterInstanceId IN ({ids})";
                int rowsLegacy = cmd.ExecuteNonQuery();
                
                // 2. Delete from ClusterSleeves_v2 table (where GetPendingClusters reads from)
                // ClusterSleeves_v2 has ClusterInstanceId column, so we can delete directly
                cmd.CommandText = $"DELETE FROM ClusterSleeves_v2 WHERE ClusterInstanceId IN ({ids})";
                int rowsV2 = cmd.ExecuteNonQuery();
                
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🗑️ DELETED {rowsLegacy} old cluster rows from ClusterSleeves, {rowsV2} rows from ClusterSleeves_v2 (Ids: {ids}) to prevent stale placement points.\n");
            }
        }

        /// <summary>
        /// Data needed for deferred DB persistence (after Revit tx commit).
        /// All Revit API data is pre-extracted so persistence can run on a background thread.
        /// </summary>
        private class ClusterPersistenceData
        {
            /// <summary>Placed element IDs as plain ints (avoids ElementId on background thread)</summary>
            public List<int> PlacedIdInts;
            public List<BatchClusterData> ClusterMap;
            public Dictionary<int, List<int>> ClusterToSleeveIdsMap;
            public Dictionary<Guid, ClashZone> ZoneCache;
            public Dictionary<int, (int comboId, int filterId, string category, string hostType, string hostOrientation)> ComboIdMap;
            /// <summary>Pre-read Revit params per cluster index: (width, height, depth, rotation)</summary>
            public Dictionary<int, (double w, double h, double d, double rot)> PreReadRevitParams;
        }

        private class PreparedPlacementData
        {
            public List<Autodesk.Revit.Creation.FamilyInstanceCreationData> CreationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
            public List<BatchClusterData> ClusterMap = new List<BatchClusterData>();
            public List<ElementId> IndividualSleevesToDelete = new List<ElementId>();
            public Dictionary<int, List<ElementId>> ClusterToSleeveIdsMap = new Dictionary<int, List<ElementId>>();
            public int SkippedDuplicateCount;
        }

        private PreparedPlacementData PrepareBulkPlacementData(
            Document doc,
            List<BatchClusterData> clusters,
            Dictionary<string, FamilySymbol> symbolCache,
            Dictionary<Guid, ClashZone> zoneCache,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker placementTracker,
            Dictionary<string, ElementId> levelMap = null)
        {
            var data = new PreparedPlacementData();
            var placedLocations = new HashSet<string>();
            int skippedDuplicateCount = 0;

            // 1. Prepare Creation Data & Deduplication
            using (placementTracker?.TrackSubOperation("Prepare Creation Data & Deduplication"))
            {
                foreach (var cluster in clusters)
                {
                    if (!symbolCache.TryGetValue(cluster.FamilyName, out var symbol)) continue;

                    XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                    string locationKey = $"{cluster.PlacementX:F6}_{cluster.PlacementY:F6}_{cluster.PlacementZ:F6}";
                    
                    if (placedLocations.Contains(locationKey))
                    {
                        skippedDuplicateCount++;
                        continue;
                    }

                    // ✅ PERF: Use explicit LevelId if available to skip nearest-neighbor search
                    ElementId levelId = null;
                    if (levelMap != null && !string.IsNullOrEmpty(cluster.MepElementLevelName) && levelMap.TryGetValue(cluster.MepElementLevelName, out var id))
                    {
                        levelId = id;
                    }

                    Autodesk.Revit.Creation.FamilyInstanceCreationData creationData;
                    if (levelId != null)
                    {
                        var level = doc.GetElement(levelId) as Level;
                        if (level != null)
                        {
                            creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, level, StructuralType.NonStructural);
                        }
                        else
                        {
                            creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, StructuralType.NonStructural);
                        }
                    }
                    else
                    {
                        creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, StructuralType.NonStructural);
                    }

                    data.CreationDataList.Add(creationData);
                    data.ClusterMap.Add(cluster);
                    placedLocations.Add(locationKey);
                }
            }

            // 2. Collect Individual Sleeves to Delete (Stage 1)
            using (placementTracker?.TrackSubOperation("Collect Individual Sleeves to Delete (Stage 1)"))
            {
                foreach (var cluster in data.ClusterMap)
                {
                    if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                    
                    var guids = cluster.ConstituentZoneGuids
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(g => g.Trim())
                        .Where(g => !string.IsNullOrEmpty(g) && Guid.TryParse(g, out _))
                        .Select(g => Guid.Parse(g))
                        .ToList();
                    
                    var sleeveIdsForThisCluster = new List<ElementId>();
                    foreach (var guid in guids)
                    {
                        if (zoneCache.TryGetValue(guid, out var z))
                        {
                            long recoveredId = _repository.TryGetSleeveInstanceIdFromSnapshot(z.Id);
                            long effectiveSleeveId = recoveredId > 0 ? recoveredId : z.SleeveInstanceId;
                            
                            if (effectiveSleeveId > 0)
                            {
                                var eid = ElementIdCompat.FromLong(effectiveSleeveId);
                                if (doc.GetElement(eid) != null)
                                {
                                    data.IndividualSleevesToDelete.Add(eid);
                                    sleeveIdsForThisCluster.Add(eid);
                                }
                            }
                        }
                    }
                    if (sleeveIdsForThisCluster.Any())
                    {
                        data.ClusterToSleeveIdsMap[data.ClusterMap.IndexOf(cluster)] = sleeveIdsForThisCluster;
                    }
                }
                data.IndividualSleevesToDelete = data.IndividualSleevesToDelete.Distinct().ToList();
            }

            data.SkippedDuplicateCount = skippedDuplicateCount;
            return data;
        }

        private void ExecuteStage1PreDelete(Document doc, List<ElementId> deleteIds)
        {
            if (deleteIds == null || deleteIds.Count == 0) return;

            using (var tx = new Transaction(doc, "Pre-Delete Clustered Individuals (Stage 1)"))
            {
                tx.Start();
                try
                {
                    doc.Delete(deleteIds);
                    tx.Commit();
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🗑️ PRE-DELETE: Deleted {deleteIds.Count} Stage 1 individual sleeves\n");
                }
                catch (Exception ex)
                {
                    tx.RollBack();
                    SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ PRE-DELETE FAILED: {ex.Message}\n");
                }
            }
        }

        private (int totalPlaced, ClusterPersistenceData persistenceData) PlaceBulkClusters(
            Document doc,
            PreparedPlacementData preparedData,
            Dictionary<Guid, ClashZone> zoneCache,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker placementTracker,
            List<FamilyInstance> placedInstances)
        {
            int totalPlaced = 0;
            ICollection<ElementId> placedIds;

            // 1. Execute Bulk Placement
            using (var step2aTracker = placementTracker?.TrackSubOperation("Step 2a: Revit NewFamilyInstances2"))
            {
                placedIds = doc.Create.NewFamilyInstances2(preparedData.CreationDataList);
                step2aTracker?.SetItemCount(placedIds.Count);
            }

            var placedIdList = placedIds.ToList();
            if (placedIdList.Count != preparedData.CreationDataList.Count)
            {
                throw new InvalidOperationException($"CRITICAL: Revit Batch Placement returned {placedIdList.Count} IDs for {preparedData.CreationDataList.Count} requests.");
            }

            // 2. Post-Placement Logic (Rotation, Parameters)
            var instanceLookup = new Dictionary<int, FamilyInstance>();
            int rotatedCount = 0;
            using (placementTracker?.TrackSubOperation("Step 3: APPLY ROTATION & PARAMS"))
            {
                for (int i = 0; i < placedIdList.Count; i++)
                {
                    var instance = doc.GetElement(placedIdList[i]) as FamilyInstance;
                    if (instance == null) continue;

                    instanceLookup[i] = instance;
                    placedInstances.Add(instance);
                    var cluster = preparedData.ClusterMap[i];

                    // A. Rotation
                    bool isPreRotated = (cluster.FamilyName ?? "").EndsWith("_X", StringComparison.OrdinalIgnoreCase);
                    if (Math.Abs(cluster.RotationAngleRad) > 1e-6 && !isPreRotated)
                    {
                        XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                        Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
                        rotatedCount++;
                    }

                    // B. Parameters
                    ClashZone templateZone = null;
                    if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                    {
                        var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                        if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid g))
                        {
                            zoneCache.TryGetValue(g, out templateZone);
                        }
                    }
                    if (templateZone == null) templateZone = new ClashZone();

                    _parameterService.QueueRectangularClusterParameters(instance.Id, cluster.ClusterWidth, cluster.ClusterHeight, cluster.ClusterDepth, instance.Id.GetIntegerValue(), templateZone);
                }
            }

            if (OptimizationFlags.UseBatchedParameterWrites && placedIdList.Count > 0)
            {
                _parameterService.FlushDeferredParameters(true, "Cluster-Bulk");
            }

            // 3. Prepare Persistence Data
            var comboIdMap = new Dictionary<int, (int comboId, int filterId, string category, string hostType, string hostOrientation)>();
            var preReadParams = new Dictionary<int, (double w, double h, double d, double rot)>();
            var clusterSleeveIdsMap = new Dictionary<int, List<int>>();

            for (int i = 0; i < placedIdList.Count; i++)
            {
                var cluster = preparedData.ClusterMap[i];
                preReadParams[i] = (cluster.ClusterWidth, cluster.ClusterHeight, cluster.ClusterDepth, cluster.RotationAngleRad);

                // Build Cluster-to-Sleeve mapping for DB
                if (preparedData.ClusterToSleeveIdsMap.TryGetValue(i, out var sIds))
                {
                    clusterSleeveIdsMap[i] = sIds.Select(eid => eid.GetIntegerValue()).ToList();
                }

                // Optimal ComboId lookup using the first constituent zone GUID
                if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                {
                    var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                    if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid firstGuid))
                    {
                        if (zoneCache.TryGetValue(firstGuid, out var fz) && fz.ComboId > 0)
                        {
                            comboIdMap[i] = ((int)fz.ComboId, -1, fz.MepElementCategory ?? "", fz.StructuralElementType ?? "", fz.HostOrientation ?? "");
                        }
                    }
                }
            }

            var persistenceData = new ClusterPersistenceData
            {
                PlacedIdInts = placedIdList.Select(eid => eid.GetIntegerValue()).ToList(),
                ClusterMap = preparedData.ClusterMap,
                ClusterToSleeveIdsMap = clusterSleeveIdsMap,
                ZoneCache = zoneCache,
                ComboIdMap = comboIdMap,
                PreReadRevitParams = preReadParams
            };

            return (placedIdList.Count, persistenceData);
        }


        /// <summary>
        /// ✅ PERF: Deferred DB persistence — fully Revit-API-free, safe for background thread.
        /// Uses pre-read params and _databasePath instead of Document/FamilyInstance.
        /// Runs in parallel with Revit tx commit to save ~284ms.
        /// </summary>
        private void PersistClusterPlacementToDb(ClusterPersistenceData data)
        {
            if (data == null || data.PlacedIdInts == null || data.PlacedIdInts.Count == 0) return;

            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] 💾 DEFERRED PERSISTENCE: Starting DB work for {data.PlacedIdInts.Count} clusters (background thread)\n");

            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Step 1: Fetch FilterIds using lightweight path-only context
            var comboIdMap = data.ComboIdMap;
            using (var dbContext = new SleeveDbContext(_databasePath))
            {
                var allComboIds = new HashSet<int>(comboIdMap.Values.Select(v => v.comboId).Where(c => c > 0));
                var filterIdMap = new Dictionary<int, int>();
                if (allComboIds.Any())
                {
                    var comboIdList = string.Join(",", allComboIds);
                    using (var cmd = dbContext.Connection.CreateCommand())
                    {
                        cmd.CommandText = $"SELECT ComboId, FilterId FROM FileCombos WHERE ComboId IN ({comboIdList})";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                filterIdMap[reader.GetInt32(0)] = reader.GetInt32(1);
                            }
                        }
                    }
                }
                var updatedMap = new Dictionary<int, (int, int, string, string, string)>();
                foreach (var kvp in comboIdMap)
                {
                    var (comboId, _, category, hostType, hostOrientation) = kvp.Value;
                    int filterId = filterIdMap.ContainsKey(comboId) ? filterIdMap[comboId] : -1;
                    updatedMap[kvp.Key] = (comboId, filterId, category, hostType, hostOrientation);
                }
                comboIdMap = updatedMap;
            }

            // Step 2: Consolidated DB transaction — all persistence in one hit
            using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
            {
                conn.Open();
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        // Pass doc=null — PrepareClusterSaveDataBatch only uses doc for diagnostics (guarded by null check)
                        var (clusterSaveDataList, placedGuidMap) = PrepareClusterSaveDataBatch(
                            null, data.PlacedIdInts, data.ClusterMap, comboIdMap, trans, data.ZoneCache);

                        // Save to ClusterSleeves_v2 using path-only context
                        if (clusterSaveDataList.Any())
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] 💾 PERSISTENCE: Saving {clusterSaveDataList.Count} clusters to v2 (shared transaction)\n");

                            using (var dbContext = new SleeveDbContext(_databasePath))
                            {
                                var clusterSleeveRepo = new ClusterSleeveRepository(dbContext, msg => SafeFileLogger.SafeAppendText("batch_v2.log", msg));
                                clusterSleeveRepo.BatchSaveClusterSleeves(clusterSaveDataList, conn, trans);
                            }
                        }

                        // Calc column updates using pre-read params (no FamilyInstance)
                        var calcColumnUpdates = new List<(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, double finalWidth, double finalHeight, double finalDepth, double finalRotation)>();
                        for (int i = 0; i < data.PlacedIdInts.Count; i++)
                        {
                            var cluster = data.ClusterMap[i];
                            if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                            var guids = cluster.ConstituentZoneGuids
                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(g => Guid.TryParse(g.Trim(), out Guid gu) ? gu : Guid.Empty)
                                .Where(gu => gu != Guid.Empty)
                                .ToList();
                            if (guids.Count == 0) continue;

                            // Use pre-read Revit values if available, fallback to cluster data
                            double w = cluster.ClusterWidth, h = cluster.ClusterHeight, d = cluster.ClusterDepth, rot = cluster.RotationAngleRad;
                            if (data.PreReadRevitParams != null && data.PreReadRevitParams.TryGetValue(i, out var preRead))
                            {
                                w = preRead.w; h = preRead.h; d = preRead.d; rot = preRead.rot;
                            }
                            calcColumnUpdates.Add((guids, cluster, data.PlacedIdInts[i], w, h, d, rot));
                        }

                        if (calcColumnUpdates.Any())
                        {
                            ExecuteUpdateClashZonesCalculatedColumnsBatchPreRead(calcColumnUpdates, conn, trans);
                        }

                        // ✅ FIX: Delete old calculation-phase rows (random GUIDs, ComboId=0)
                        // since BatchSaveClusterSleeves already created new rows with deterministic GUIDs,
                        // correct ComboId/FilterId, corners, and MEP data.
                        // Without this, both old (random GUID) and new (deterministic GUID) rows coexist → duplicates.
                        if (placedGuidMap.Any())
                        {
                            var oldGuids = placedGuidMap.Keys.ToList();
                            // Chunk delete to stay within SQLite parameter limits
                            for (int offset = 0; offset < oldGuids.Count; offset += 100)
                            {
                                var chunk = oldGuids.Skip(offset).Take(100).ToList();
                                using (var delCmd = conn.CreateCommand())
                                {
                                    delCmd.Transaction = trans;
                                    var inList = string.Join(",", chunk.Select((_, i) => $"@dg{i}"));
                                    delCmd.CommandText = $"DELETE FROM ClusterSleeves_v2 WHERE ClusterGUID IN ({inList})";
                                    for (int i = 0; i < chunk.Count; i++)
                                        delCmd.Parameters.AddWithValue($"@dg{i}", chunk[i]);
                                    int deleted = delCmd.ExecuteNonQuery();
                                    SafeFileLogger.SafeAppendText("batch_v2.log",
                                        $"[{DateTime.Now:HH:mm:ss}] 🗑️ CLEANUP: Deleted {deleted} old calculation-phase rows from ClusterSleeves_v2 (replaced by deterministic-GUID rows)\n");
                                }
                            }
                        }

                        // ✅ CRITICAL FIX: Update SleeveSnapshots with ClusterInstanceId for all constituent zones
                        // This ensures snapshots are linked to the placed cluster sleeves
                        UpdateSleeveSnapshotsWithClusterId(data, conn, trans);

                        trans.Commit();
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Deferred DB Batch failed: {ex.Message}\n{ex.StackTrace}\n");
                        throw;
                    }
                }
            }

            sw.Stop();
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] ✅ DEFERRED PERSISTENCE COMPLETE: {data.PlacedIdInts.Count} clusters persisted in {sw.ElapsedMilliseconds}ms (background thread)\n");
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Update SleeveSnapshots table with ClusterInstanceId for all constituent zones.
        /// This links the GUID-based refresh snapshots to the placed cluster sleeves.
        /// Called during PersistClusterPlacementToDb transaction.
        /// 
        /// ✅ BUG FIX (2025-03-05): Each constituent zone gets its OWN snapshot row with its SPECIFIC parameters.
        /// Previously, all zones in a cluster got the same aggregated parameters.
        /// </summary>
        private void UpdateSleeveSnapshotsWithClusterId(
            ClusterPersistenceData data, 
            SQLiteConnection conn, 
            SQLiteTransaction trans)
        {
            if (data?.PlacedIdInts == null || data.PlacedIdInts.Count == 0 || 
                data?.ClusterMap == null || data.ClusterMap.Count == 0)
                return;

            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] 📸 SNAPSHOT UPDATE: Updating SleeveSnapshots with ClusterInstanceId for {data.PlacedIdInts.Count} clusters\n");

            int updatedCount = 0;
            int insertedCount = 0;
            int skippedCount = 0;

            var snapshotUpdates = new List<(string CleanGuid, int ClusterId)>();

            for (int i = 0; i < data.PlacedIdInts.Count; i++)
            {
                int clusterInstanceId = data.PlacedIdInts[i];
                var cluster = data.ClusterMap[i];

                if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                {
                    skippedCount++;
                    continue;
                }

                // Parse constituent zone GUIDs
                var parts = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                bool hasValid = false;
                foreach (var part in parts)
                {
                    var trimmed = part.Trim();
                    if (Guid.TryParse(trimmed, out Guid g))
                    {
                        // Normalize to uppercase with hyphens, no braces
                        snapshotUpdates.Add((g.ToString().ToUpperInvariant(), clusterInstanceId));
                        hasValid = true;
                    }
                }
                if (!hasValid) skippedCount++;
            }

            if (snapshotUpdates.Count == 0) return;

            // Batch Update (Chunk size 40 to stay safe with complex query parameters)
            const int batchSize = 40;
            for (int i = 0; i < snapshotUpdates.Count; i += batchSize)
            {
                var batch = snapshotUpdates.Skip(i).Take(batchSize).ToList();
                try
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = trans;
                        
                        var caseBuilder = new System.Text.StringBuilder();
                        var inClauseBuilder = new System.Text.StringBuilder();
                        
                        caseBuilder.Append("CASE ");
                        
                        // Use normalized matching in SQL: UPPER(REPLACE(..., '{', ''), '}', ''))
                        string cleanCol = "UPPER(REPLACE(REPLACE(ClashZoneGuid, '{', ''), '}', ''))";

                        for (int j = 0; j < batch.Count; j++)
                        {
                            string pGuid = $"@g{j}";
                            string pId = $"@id{j}";
                            
                            // WHEN Clean(Col) = @g THEN @id
                            caseBuilder.Append($"WHEN {cleanCol} = {pGuid} THEN {pId} ");
                            
                            inClauseBuilder.Append(j == 0 ? pGuid : $", {pGuid}");
                            
                            cmd.Parameters.AddWithValue(pGuid, batch[j].CleanGuid);
                            cmd.Parameters.AddWithValue(pId, batch[j].ClusterId);
                        }
                        
                        caseBuilder.Append("END");

                        // 1. UPDATE existing snapshot rows (e.g. pipe zones placed individually first)
                        cmd.CommandText = $@"
                            UPDATE SleeveSnapshots
                            SET ClusterInstanceId = {caseBuilder},
                                SourceType = 'Cluster',
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE {cleanCol} IN ({inClauseBuilder})";

                        int rows = cmd.ExecuteNonQuery();
                        updatedCount += rows;

                        SafeFileLogger.SafeAppendText("placement_performance.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ✅ Updated {rows} existing snapshots in batch (targets: {batch.Count})\n");
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to update snapshot batch: {ex.Message}\n");
                }
            }

            // ✅ BUG FIX: INSERT one row PER zone with that zone's SPECIFIC parameters
            // This ensures pipe zones get pipe params, duct zones get duct params, etc.
            foreach (var (zoneGuid, clusterId) in snapshotUpdates)
            {
                try
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = trans;
                        string cleanCol = "UPPER(REPLACE(REPLACE(ClashZoneGuid, '{{', ''), '}}', ''))";
                        
                        // Check if this zone already has a snapshot
                        cmd.CommandText = $@"
                            SELECT COUNT(*) FROM SleeveSnapshots 
                            WHERE {cleanCol} = @Guid";
                        cmd.Parameters.AddWithValue("@Guid", zoneGuid);
                        
                        var existingCount = Convert.ToInt32(cmd.ExecuteScalar());
                        
                        if (existingCount > 0)
                        {
                            // Already exists - was updated in the batch UPDATE above
                            continue;
                        }
                        
                        // ✅ CRITICAL: INSERT one row for THIS specific zone with ITS specific parameters
                        cmd.CommandText = $@"
                            INSERT INTO SleeveSnapshots (
                                ClusterInstanceId, SourceType,
                                ClashZoneGuid, MepParametersJson, HostParametersJson,
                                CreatedAt, UpdatedAt
                            )
                            SELECT
                                @ClusterId,
                                'Cluster',
                                ClashZoneGuid,
                                MepParameterValuesJson,
                                HostParameterValuesJson,
                                CURRENT_TIMESTAMP,
                                CURRENT_TIMESTAMP
                            FROM ClashZones
                            WHERE {cleanCol} = @Guid";
                        
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@ClusterId", clusterId);
                        cmd.Parameters.AddWithValue("@Guid", zoneGuid);
                        
                        int rows = cmd.ExecuteNonQuery();
                        insertedCount += rows;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_performance.log",
                                $"[{DateTime.Now:HH:mm:ss}]   ✅ Inserted snapshot for zone {zoneGuid} in cluster {clusterId}\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to insert snapshot for zone {zoneGuid}: {ex.Message}\n");
                }
            }

            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] 📸 SNAPSHOT UPDATE COMPLETE: {updatedCount} updated, {insertedCount} inserted (cross-category), {skippedCount} skipped\n");
        }

        /// <summary>
        /// ✅ CRITICAL: Update cluster bounding boxes in database AFTER parameter flush and regeneration
        /// Updates both ClusterSleeves table and ClashZones.ClusterSleeveBoundingBox columns
        /// This ensures Stage 2 cleanup can find overlapping sleeves using DB-only queries
        /// </summary>
        private const int BboxBatchChunkSize = 100;

        private void UpdateClusterBoundingBoxesAfterPlacement(Document doc, List<FamilyInstance> placedInstances)
        {
            try
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 💾 BATCH BBOX UPDATE: Starting update for {placedInstances.Count} instances\n");

                var rows = new List<(int id, double minX, double minY, double minZ, double maxX, double maxY, double maxZ)>();
                foreach (var instance in placedInstances)
                {
                    if (instance == null || !instance.IsValidObject) continue;
                    var bbox = instance.get_BoundingBox(null);
                    if (bbox == null || !bbox.Enabled) continue;
                    rows.Add((instance.Id.GetIntegerValue(), bbox.Min.X, bbox.Min.Y, bbox.Min.Z, bbox.Max.X, bbox.Max.Y, bbox.Max.Z));
                }

                if (rows.Count == 0) return;

                using (var context = new SleeveDbContext(doc))
                {
                    // context.Connection.Open(); // ✅ CRASH FIX: SQLiteConnection already open in SleeveDbContext constructor
                    using (var trans = context.Connection.BeginTransaction())
                    {
                        for (int offset = 0; offset < rows.Count; offset += BboxBatchChunkSize)
                        {
                            var chunk = rows.Skip(offset).Take(BboxBatchChunkSize).ToList();
                            if (chunk.Count == 0) continue;

                            var inList = string.Join(",", chunk.Select((_, i) => $"@id{i}"));
                            var caseMinX = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @minx{i}"));
                            var caseMinY = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @miny{i}"));
                            var caseMinZ = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @minz{i}"));
                            var caseMaxX = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @maxx{i}"));
                            var caseMaxY = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @maxy{i}"));
                            var caseMaxZ = string.Join(" ", chunk.Select((_, i) => $"WHEN @id{i} THEN @maxz{i}"));

                            using (var cmd = context.Connection.CreateCommand())
                            {
                                cmd.Transaction = trans;
                                cmd.CommandText = $@"
                                    UPDATE ClusterSleeves_v2 SET
                                        BoundingBoxMinX = CASE ClusterInstanceId {caseMinX} END,
                                        BoundingBoxMinY = CASE ClusterInstanceId {caseMinY} END,
                                        BoundingBoxMinZ = CASE ClusterInstanceId {caseMinZ} END,
                                        BoundingBoxMaxX = CASE ClusterInstanceId {caseMaxX} END,
                                        BoundingBoxMaxY = CASE ClusterInstanceId {caseMaxY} END,
                                        BoundingBoxMaxZ = CASE ClusterInstanceId {caseMaxZ} END
                                    WHERE ClusterInstanceId IN ({inList});
                                    UPDATE ClusterSleeves SET
                                        BoundingBoxMinX = CASE ClusterInstanceId {caseMinX} END,
                                        BoundingBoxMinY = CASE ClusterInstanceId {caseMinY} END,
                                        BoundingBoxMinZ = CASE ClusterInstanceId {caseMinZ} END,
                                        BoundingBoxMaxX = CASE ClusterInstanceId {caseMaxX} END,
                                        BoundingBoxMaxY = CASE ClusterInstanceId {caseMaxY} END,
                                        BoundingBoxMaxZ = CASE ClusterInstanceId {caseMaxZ} END
                                    WHERE ClusterInstanceId IN ({inList})";
                                for (int i = 0; i < chunk.Count; i++)
                                {
                                    var r = chunk[i];
                                    cmd.Parameters.AddWithValue($"@id{i}", r.id);
                                    cmd.Parameters.AddWithValue($"@minx{i}", r.minX);
                                    cmd.Parameters.AddWithValue($"@miny{i}", r.minY);
                                    cmd.Parameters.AddWithValue($"@minz{i}", r.minZ);
                                    cmd.Parameters.AddWithValue($"@maxx{i}", r.maxX);
                                    cmd.Parameters.AddWithValue($"@maxy{i}", r.maxY);
                                    cmd.Parameters.AddWithValue($"@maxz{i}", r.maxZ);
                                }
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    }
                }
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH BBOX UPDATE COMPLETE: Updated {rows.Count} clusters in BOTH v2 and legacy tables\n");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ BATCH BBOX UPDATE ERROR: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        private double ReadDouble(SQLiteDataReader reader, string column)
        {
            try { return reader.IsDBNull(reader.GetOrdinal(column)) ? 0.0 : Convert.ToDouble(reader[column]); }
            catch { return 0.0; }
        }

        private long ReadLong(SQLiteDataReader reader, string column)
        {
            try { return reader.IsDBNull(reader.GetOrdinal(column)) ? 0 : Convert.ToInt64(reader[column]); }
            catch { return 0; }
        }
    }

    public class BatchClusterData
    {
        public string ClusterGUID { get; set; } = string.Empty;
        public double PlacementX { get; set; }
        public double PlacementY { get; set; }
        public double PlacementZ { get; set; }
        public double ClusterWidth { get; set; }
        public double ClusterHeight { get; set; }
        public double ClusterDepth { get; set; }
        public double RotationAngleRad { get; set; }
        public long HostElementId { get; set; }
        public string? HostType { get; set; }
        public string? HostOrientation { get; set; }
        public string? MepElementLevelName { get; set; } // ✅ Added for level-based placement
        public string FamilyName { get; set; } = string.Empty;
        public string ConstituentZoneGuids { get; set; } = string.Empty;
        public bool IsCrossCategory { get; set; }
    }
}
