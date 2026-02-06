using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
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

        public BatchClusterPlacementService(
            string databasePath, 
            IClashZoneRepository repository,
            JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService parameterService,
            IClusterCleanupService cleanupService = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor performanceMonitor = null)
        {
            _databasePath = databasePath;
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _parameterService = parameterService ?? throw new ArgumentNullException(nameof(parameterService));
            _cleanupService = cleanupService ?? new ClusterCleanupService(); // Default if not injected
            _performanceMonitor = performanceMonitor;
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
                    using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
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
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CONSOLIDATED CLEANUP: Found {clusterInstanceIds.Count} placed clusters from all categories\n");
                        
                        // ✅ STAGE 2 CLEANUP FIX: Ensure ClusterSleeves (legacy) has bounding boxes before point-in-box query
                        // Cleanup requires ClusterSleeves.BoundingBox* != 0; if in-placement bbox update was skipped/failed, do it now
                        var instancesForBbox = new List<FamilyInstance>();
                        foreach (int id in clusterInstanceIds)
                        {
                            var elem = doc.GetElement(new Autodesk.Revit.DB.ElementId(id)) as FamilyInstance;
                            if (elem != null && elem.IsValidObject) instancesForBbox.Add(elem);
                        }
                        if (instancesForBbox.Count > 0)
                        {
                            try
                            {
                                UpdateClusterBoundingBoxesAfterPlacement(doc, instancesForBbox);
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2: Ensured {instancesForBbox.Count} cluster bboxes in DB before cleanup\n");

                                // ✅ CRITICAL FIX: Extract and save corners for Cluster Sleeves
                                // This ensures that cluster sleeves have valid 3D geometry (corners) for the proximity check
                                try
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 📐 STAGE 2: Extracting corners for cluster sleeves...\n");
                                    var cornerExtractor = new JSE_RevitAddin_MEP_OPENINGS.Services.Calculation.BatchSleeveCornerExtractor(_repository);
                                    int cornersExtracted = cornerExtractor.ExtractAndSaveCornersForClusters(doc);
                                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 📐 STAGE 2: Extracted corners for {cornersExtracted} cluster sleeves.\n");
                                }
                                catch (Exception cornerEx)
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 cluster corner extraction failed (continuing): {cornerEx.Message}\n");
                                }
                            }
                            catch (Exception bboxEx)
                            {
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 bbox pre-cleanup failed (continuing): {bboxEx.Message}\n");
                            }
                        }
                        
                        // ✅ CLEANUP ALL CATEGORIES: No category filter - check all individual sleeves against all cluster bboxes
                        cleanedUp = cleanupService.CleanupSleevesWithinClustersFromDatabase(
                            doc, 
                            targetCategory: null, // null = all categories
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
            using (_performanceMonitor?.TrackOperation("Step 1: LOAD CLUSTERS FROM DB"))
            {
                pendingClusters = GetPendingClusters(doc, batchId);
            }
            
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: GetPendingClusters returned {pendingClusters.Count} clusters\n");
            
            // ✅ FIX: Early return BEFORE starting the main tracker if no clusters to place
            if (pendingClusters.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: No pending clusters found, returning early (ClusterSleeves table will not be populated)\n");
                return (0, 0);
            }
            
            // ✅ FIX: Only start the main tracker if there are clusters to place
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ STARTING PLACEMENT V2: Batch {batchId}, Pending={pendingClusters.Count}, Mode={(useSingleTransaction ? "BULK" : "SEQUENTIAL")}\n");
            
            using (var tracker = _performanceMonitor?.TrackOperation("Step 2: CLUSTER PLACEMENT MAIN LOOP"))
            {
                int placedCount = 0;
                int failedCount = 0;
                var placedInstances = new List<FamilyInstance>(); // Track placed instances for 2nd cleanup

                // ✅ OPTIMIZATION: Pre-fetch all required Family Symbols
                var symbolCache = new Dictionary<string, FamilySymbol>();
                using (_performanceMonitor?.TrackOperation("Step 1a: CACHE CLUSTER SYMBOLS"))
                {
                    var uniqueFamilyNames = pendingClusters.Select(c => c.FamilyName).Distinct().ToList();
                    var allSymbols = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .ToList();

                    foreach (var name in uniqueFamilyNames)
                    {
                        var symbol = allSymbols.FirstOrDefault(s => s.Name == name || s.Family.Name == name);
                        if (symbol != null) symbolCache[name] = symbol;
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
                
                // ✅ CLUSTER OPTIMIZATION START (SINGLE TRANSACTION FLOW)
                // We manage the transaction if one doesn't exist.
                // If one DOES exist, we use it, but assume the caller will commit.
                // WE ALWAYS USE PlaceBulkClusters (NewFamilyInstances2) for speed.

                bool manageTransaction = !doc.IsModifiable;
                Transaction mainTransaction = null;
                bool transactionCommitted = false;

                try 
                {
                    if (manageTransaction) 
                    {
                        mainTransaction = new Transaction(doc, "Place Batch Clusters V2 (Unified)");
                        mainTransaction.Start();
                    }

                    // 1. PLACE CLUSTERS (NewFamilyInstances2)
                    using (var placementTracker = _performanceMonitor?.TrackOperation("Step 2b: EXECUTE BULK CLUSTER PLACEMENT"))
                    {
                        // 🔍 DIAGNOSTIC 1: Log which code path is chosen
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC 1: UseBulkClusterSleevePlacement = {OptimizationFlags.UseBulkClusterSleevePlacement}\n");
                        
                        if (OptimizationFlags.UseBulkClusterSleevePlacement)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ USING BULK PLACEMENT PATH (calling PlaceBulkClusters)\n");
                            // Always use bulk placement logic, even inside an existing transaction
                            // NewFamilyInstances2 works fine inside existing transactions
                            placedCount = PlaceBulkClusters(doc, pendingClusters, symbolCache, zoneCache, placementTracker, placedInstances);
                        }
                        else
                        {
                            // Fallback to sequential Loop (only if flag is specifically disabled)
                            // This path should ideally be deprecated
                            foreach (var cluster in pendingClusters)
                            {
                                if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                else failedCount++;
                            }
                        }
                    }

                    // 2. FLUSH PARAMETERS (Inside the transaction!)
                    // This is the key to Single Transaction Flow.
                    // We flush parameters BEFORE committing the placement transaction.
                    // If we are in a nested transaction, we assume the caller will handle the final commit,
                    // but we still flush here to act on the freshly placed elements.
                    if (OptimizationFlags.UseBatchedParameterWrites && placedCount > 0)
                    {
                        using (var flushTracker = _performanceMonitor?.TrackOperation("Step 4: FLUSH CLUSTER PARAMETERS"))
                        {
                            var clusterFlushTimer = System.Diagnostics.Stopwatch.StartNew();
                            
                            // Flush logic:
                            int flushedCount = _parameterService.FlushDeferredParameters(clearList: true);
                            doc.Regenerate();
                            
                            clusterFlushTimer.Stop();
                            SafeFileLogger.SafeAppendText("performance.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER-FLUSH] Clusters={placedCount}, Parameters={flushedCount}, Time={clusterFlushTimer.ElapsedMilliseconds}ms, InExistingTx={!manageTransaction}\n");
                            
                            flushTracker?.SetItemCount(flushedCount);
                        }
                    }

                    // 3. UPDATE BOUNDING BOXES (After regen)
                    if (placedInstances.Count > 0)
                    {
                        using (var bboxTracker = _performanceMonitor?.TrackOperation("Step 5: RETRIEVE CLUSTER DATA FROM MODEL"))
                        {
                            UpdateClusterBoundingBoxesAfterPlacement(doc, placedInstances);
                            bboxTracker?.SetItemCount(placedInstances.Count);
                        }
                    }

                    if (manageTransaction) 
                    {
                        mainTransaction.Commit();
                        transactionCommitted = true;
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
            
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
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
                        cs.ClusterInstanceId
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
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            totalInTable++;
                            var clusterGuid = reader["ClusterGUID"]?.ToString() ?? "";
                            var clusterInstanceId = reader["ClusterInstanceId"] != DBNull.Value ? Convert.ToInt32(reader["ClusterInstanceId"]) : 0;
                            
                            if (clusterInstanceId <= 0)
                            {
                                // ✅ NEW CLUSTER: Never placed - needs placement
                                needsPlacementCount++;
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 NEW: Cluster {clusterGuid} - needs placement\n");
                                
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
                                    ConstituentZoneGuids = reader["ConstituentZoneGuids"]?.ToString() ?? ""
                                });
                            }
                            else
                            {
                                // ✅ RECOVERY: ClusterInstanceId > 0 - check if element still exists in Revit
                                bool existsInRevit = existenceChecker.ClusterSleeveExists(clusterInstanceId);
                                if (!existsInRevit)
                                {
                                    needsPlacementCount++;
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 PATH 1 RECOVERY: Cluster {clusterGuid} [ID={clusterInstanceId}] deleted from Revit - will re-place\n");
                                    
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
                                        ConstituentZoneGuids = reader["ConstituentZoneGuids"]?.ToString() ?? ""
                                    });
                                }
                                else
                                {
                                    alreadyExistsCount++;
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1: Cluster {clusterGuid} [ID={clusterInstanceId}] exists in Revit - skipping\n");
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

        private void UpdateStatusBatch(Dictionary<string, int> guidToIdMap, string status)
        {
            if (guidToIdMap == null || guidToIdMap.Count == 0) return;

            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var trans = conn.BeginTransaction())
                {
                    foreach (var kvp in guidToIdMap)
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = trans;
                            cmd.CommandText = @"
                                UPDATE ClusterSleeves_v2 
                                SET Status = @status, 
                                    ClusterInstanceId = @id,
                                    PlacedAt = CURRENT_TIMESTAMP
                                WHERE ClusterGUID = @guid;

                                -- ✅ STURDY PERSISTENCE: UPSERT into legacy table
                                -- 1. Try to insert from V2 if not exists
                                INSERT OR IGNORE INTO ClusterSleeves (
                                    ClusterGUID, ClusterInstanceId, ClusterBatchId,
                                    PlacementX, PlacementY, PlacementZ,
                                    ClusterWidth, ClusterHeight, ClusterDepth,
                                    RotationAngleRad, HostElementId, HostType, HostOrientation,
                                    Category, FamilyName, ConstituentZoneGuids, ComboId, FilterId,
                                    Status, ValidationStatus, CreatedAt, UpdatedAt
                                )
                                SELECT 
                                    ClusterGUID, @id, ClusterBatchId,
                                    PlacementX, PlacementY, PlacementZ,
                                    ClusterWidth, ClusterHeight, ClusterDepth,
                                    RotationAngleRad, HostElementId, HostType, HostOrientation,
                                    Category, FamilyName, ConstituentZoneGuids, ComboId, FilterId,
                                    @status, ValidationStatus, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                                FROM ClusterSleeves_v2
                                WHERE ClusterGUID = @guid;

                                -- 2. Update status/ID (in case it already existed)
                                UPDATE ClusterSleeves
                                SET ClusterInstanceId = @id,
                                    Status = @status,
                                    UpdatedAt = CURRENT_TIMESTAMP
                                WHERE ClusterGUID = @guid";
                            cmd.Parameters.AddWithValue("@status", status);
                            cmd.Parameters.AddWithValue("@id", kvp.Value);
                            cmd.Parameters.AddWithValue("@guid", kvp.Key);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    trans.Commit();
                }
            }
        }

        private void UpdateStatus(string guid, string status, string msg = null, int instanceId = -1)
        {
             using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
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
            var databaseUpdates = new List<(System.Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)>();

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
            var oldClusterInstanceIds = new List<int>();

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
                        int recoveredId = GetParentClusterIdFromV2(parsedGuid);
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
                DeleteOldClusterSleeves(oldClusterInstanceIds);
            }

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 PerformSwapDeletion: Retrieved {zones.Count} zones for cluster {cluster.ClusterGUID}\n");


            foreach (var z in zones)
            {
                // ✅ ROBUST PERSISTENCE: ALWAYS try to recover from snapshot first
                // This decouples "History" (AfterClusterSleeveId) from "Current State" (SleeveInstanceId)
                // If snapshot exists, it is the undeniable truth of what was there.
                int effectiveSleeveId = _repository.TryGetSleeveInstanceIdFromSnapshot(z.Id);
                
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
                    toDeleteIds.Add(new ElementId(effectiveSleeveId));
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
                    MarkedForClusterProcess: true, // ✅ Keep true - this zone is in a cluster (was set during calculation)
                    AfterClusterSleeveId: effectiveSleeveId // ✅ Correctly persists recovered ID
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
                var element = doc.GetElement(new ElementId(clusterElementId)) as FamilyInstance;
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
        private int GetParentClusterIdFromV2(Guid zoneGuid)
        {
            try
            {
                using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
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
                        if (result != null && int.TryParse(result.ToString(), out int id))
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
        /// ✅ FIX: Update ClashZones table with calculated cluster dimensions
        /// Populates: CalculatedSleeveWidth, Height, Depth, Rotation, FamilyName, PlacedAt
        /// </summary>
        private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance = null)
        {
            // 🔍 LOG BLOCK 1: Entry point
            SafeFileLogger.SafeAppendText("debug_db.log", 
                $"\n[{DateTime.Now:HH:mm:ss}] === UpdateClashZonesCalculatedColumns START ===\n" +
                $"    ZoneGuids Count: {zoneGuids?.Count ?? 0}\n" +
                $"    ClusterInstanceId: {clusterInstanceId}\n" +
                $"    Width: {cluster.ClusterWidth:F4}, Height: {cluster.ClusterHeight:F4}\n" +
                $"    HasActualInstance: {actualInstance != null}\n");
            
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

                    SafeFileLogger.SafeAppendText("debug_db.log", 
                        $"    🎯 ACTUAL GEOMETRY: W={finalWidth * 304.8:F1}mm, H={finalHeight * 304.8:F1}mm, R={finalRotation:F4}\n");
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("debug_db.log", $"    ⚠️ Error reading actual values: {ex.Message}\n");
                }
            }
            
            if (zoneGuids == null || !zoneGuids.Any()) 
            {
                SafeFileLogger.SafeAppendText("debug_db.log", "    ❌ EARLY RETURN: No GUIDs\n");
                return;
            }

            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            
                            // Build WHERE clause with all GUIDs
                            var guidParams = new List<string>();
                            for (int i = 0; i < zoneGuids.Count; i++)
                            {
                                guidParams.Add($"@Guid{i}");
                                // ✅ CRITICAL FIX: Don't use ToUpperInvariant() - database stores GUIDs in original case
                                // Using UPPER() in WHERE clause handles case-insensitive matching
                                cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString());
                            }
                            
                            // ✅ UPDATE: Populate all calculated columns
                            // ✅ CRITICAL FIX: DO NOT update SleevePlacementX/Y/Z with cluster placement point!
                            // Individual sleeves must keep their own placement points. Cluster placement point is stored in ClusterSleeves/ClusterSleeves_v2 tables.
                            cmd.CommandText = $@"
                                UPDATE ClashZones 
                                SET 
                                    CalculatedSleeveWidth = @Width,
                                    CalculatedSleeveHeight = @Height,
                                    CalculatedSleeveDepth = @Depth,
                                    CalculatedRotation = @Rotation,
                                    CalculatedFamilyName = @FamilyName,
                                    -- ✅ REMOVED: SleevePlacementX/Y/Z - DO NOT overwrite individual sleeve placement points with cluster placement point!
                                    -- Cluster placement point is stored in ClusterSleeves/ClusterSleeves_v2 tables, not in ClashZones
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
                            
                            // ✅ CRITICAL: Log ClusterInstanceId being saved to ClashZones
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 💾 SAVING ClusterInstanceId={clusterInstanceId} to ClashZones for {zoneGuids.Count} zones (GUIDs: {string.Join(", ", zoneGuids.Take(3))}...)\n");
                            
                            // ✅ CRITICAL FIX: DO NOT save cluster placement point to SleevePlacementX/Y/Z
                            // Individual sleeves must keep their own placement points
                            // Cluster placement point is stored in ClusterSleeves/ClusterSleeves_v2 tables via SaveToClusterSleevesLegacy
                            
                            // ✅ LOGGING: Log cluster placement point (for debugging, but NOT saving to ClashZones.SleevePlacementX/Y/Z)
                            var loc = actualInstance?.Location as LocationPoint;
                            double actualPX = loc?.Point?.X ?? cluster.PlacementX;
                            double actualPY = loc?.Point?.Y ?? cluster.PlacementY;
                            double actualPZ = loc?.Point?.Z ?? cluster.PlacementZ;
                            
                            SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                $"\n[{DateTime.Now:HH:mm:ss}] 📍 CLUSTER PLACEMENT POINT (NOT saved to ClashZones.SleevePlacementX/Y/Z) - ClusterInstanceId={clusterInstanceId}:\n" +
                                $"    THEORETICAL (from calculation): X={cluster.PlacementX:F6}, Y={cluster.PlacementY:F6}, Z={cluster.PlacementZ:F6}\n" +
                                $"    ACTUAL (from Revit instance):   X={(loc?.Point?.X ?? double.NaN):F6}, Y={(loc?.Point?.Y ?? double.NaN):F6}, Z={(loc?.Point?.Z ?? double.NaN):F6}\n" +
                                $"    HostType: {cluster.HostType}, HostOrientation: {cluster.HostOrientation}\n" +
                                $"    Rotation: {cluster.RotationAngleRad:F6} rad ({cluster.RotationAngleRad * 180.0 / Math.PI:F2}°)\n" +
                                $"    ⚠️ NOTE: This cluster placement point is stored in ClusterSleeves/ClusterSleeves_v2, NOT in ClashZones.SleevePlacementX/Y/Z\n");
                            
                            // Calculate displacement if both are available
                            if (loc?.Point != null)
                            {
                                double deltaX = loc.Point.X - cluster.PlacementX;
                                double deltaY = loc.Point.Y - cluster.PlacementY;
                                double deltaZ = loc.Point.Z - cluster.PlacementZ;
                                double totalDisplacement = Math.Sqrt(deltaX * deltaX + deltaY * deltaY + deltaZ * deltaZ);
                                
                                SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                    $"    DISPLACEMENT: ΔX={deltaX * 304.8:F2}mm, ΔY={deltaY * 304.8:F2}mm, ΔZ={deltaZ * 304.8:F2}mm, Total={totalDisplacement * 304.8:F2}mm\n");
                                
                                if (totalDisplacement > 0.01) // More than 1cm displacement
                                {
                                    SafeFileLogger.SafeAppendText("cluster_placementpoint.log", 
                                        $"    ⚠️ WARNING: Significant displacement detected! (>1cm)\n");
                                }
                            }
                            
                            // ✅ Corners
                            var cornerService = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
                            // ✅ HIGH ACCURACY: Pass HostType and HostOrientation to ensure correct strategy selection
                            var corners = actualInstance != null ? cornerService.CalculateCornersFromInstance(actualInstance, cluster.HostOrientation, cluster.HostType) : null;
                            
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
                            
                            int rowsAffected = cmd.ExecuteNonQuery();
                            
                            // 🔍 LOG BLOCK 2: Check rows affected
                            SafeFileLogger.SafeAppendText("debug_db.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATE EXECUTED: {rowsAffected} rows affected\n");

                            if (rowsAffected == 0)
                            {
                                SafeFileLogger.SafeAppendText("debug_db.log", 
                                    $"    ⚠️ ZERO ROWS! GUID MISMATCH!\n" +
                                    $"    First GUID from code: {zoneGuids.First()}\n");
                                
                                // Check sample GUID from database
                                using (var checkCmd = conn.CreateCommand())
                                {
                                    checkCmd.Transaction = transaction;
                                    checkCmd.CommandText = "SELECT ClashZoneGuid FROM ClashZones WHERE ClashZoneGuid IS NOT NULL LIMIT 1";
                                    var dbGuid = checkCmd.ExecuteScalar()?.ToString();
                                    SafeFileLogger.SafeAppendText("debug_db.log", 
                                        $"    Sample GUID from DB: {dbGuid}\n");
                                }
                            }
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 📊 UPDATED {rowsAffected} ClashZones calculated columns (Width={cluster.ClusterWidth:F3}, Height={cluster.ClusterHeight:F3}, ClusterInstanceId={clusterInstanceId})\n");
                            
                            // ✅ VERIFY: Check if ClusterInstanceId was actually saved to ClashZones
                            if (rowsAffected > 0)
                            {
                                using (var verifyCmd = conn.CreateCommand())
                                {
                                    verifyCmd.Transaction = transaction;
                                    var verifyGuidParams = new List<string>();
                                    for (int i = 0; i < zoneGuids.Count; i++)
                                    {
                                        verifyGuidParams.Add($"@VerifyGuid{i}");
                                        verifyCmd.Parameters.AddWithValue($"@VerifyGuid{i}", zoneGuids[i].ToString());
                                    }
                                    verifyCmd.CommandText = $@"
                                        SELECT COUNT(*) FROM ClashZones 
                                        WHERE UPPER(ClashZoneGuid) IN ({string.Join(", ", verifyGuidParams.Select((_, i) => $"UPPER(@VerifyGuid{i})"))})
                                          AND ClusterInstanceId = @VerifyClusterInstanceId";
                                    verifyCmd.Parameters.AddWithValue("@VerifyClusterInstanceId", clusterInstanceId);
                                    var verifiedCount = Convert.ToInt32(verifyCmd.ExecuteScalar());
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ VERIFIED: {verifiedCount} ClashZones now have ClusterInstanceId={clusterInstanceId} saved in database\n");
                                }
                            }
                        }
                        
                        transaction.Commit();
                        SafeFileLogger.SafeAppendText("debug_db.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ TRANSACTION COMMITTED\n");
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        
                        // 🔍 LOG BLOCK 3: Exception details
                        SafeFileLogger.SafeAppendText("debug_db.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ EXCEPTION: {ex.Message}\n");
                        
                        if (ex.Message.Contains("no such column"))
                        {
                            SafeFileLogger.SafeAppendText("debug_db.log", 
                                $"    🔥 COLUMN NAME MISMATCH! Run: PRAGMA table_info(ClashZones)\n");
                        }

                        SafeFileLogger.SafeAppendText("placement_errors.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ Failed to update ClashZones calculated columns: {ex.Message}\n");
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// ✅ FIX: Save cluster to ClusterSleeves (legacy) table for PATH 1 compatibility
        /// This enables cluster replay from database
        /// </summary>
        private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId, FamilyInstance actualInstance = null)
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
                            comboId = firstZone.ComboId;
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
            List<FamilyInstance> placedInstances = null)
        {
            // 🔥 DIAGNOSTIC: Log entry and transaction state
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
                
                // ✅ LOGGING: Log theoretical placement point before placement
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
                    var existingId = existingClusters.First().Id.IntegerValue;
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
                    instance = doc.Create.NewFamilyInstance(location, symbol, structuralType);
                }

                if (instance != null)
                {
                    int clusterInstanceId = instance.Id.IntegerValue;

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
                    if (Math.Abs(cluster.RotationAngleRad) > 1e-6)
                    {
                        using (parentTracker != null ? parentTracker.TrackSubOperation("Revit Rotate Cluster") : _performanceMonitor?.TrackOperation("Revit Rotate Cluster"))
                        {
                            Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
                        }
                    }
                    
                    // ✅ CRITICAL FIX: Regenerate document to ensure all transformations are applied
                    doc.Regenerate();
                    
                    // ✅ CRITICAL FIX: Refresh instance reference after regeneration
                    instance = doc.GetElement(new ElementId(clusterInstanceId)) as FamilyInstance;
                    
                    // ✅ LOGGING: Log actual placement point after all transformations
                    if (instance != null && instance.IsValidObject)
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
                    else
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
        private void DeleteOldClusterSleeves(List<int> clusterInstanceIds)
        {
            if (clusterInstanceIds == null || !clusterInstanceIds.Any()) return;

            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try 
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

        /// <summary>
        /// ✅ NEW: High-performance Bulk Placement for Clusters
        /// Uses doc.Create.NewFamilyInstances2 to place all clusters in 1-2 API calls.
        /// </summary>
        private int PlaceBulkClusters(
            Document doc, 
            List<BatchClusterData> clusters, 
            Dictionary<string, FamilySymbol> symbolCache, 
            Dictionary<Guid, ClashZone> zoneCache,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker placementTracker,
            List<FamilyInstance> placedInstances)
        {
            var buildTimestamp = "2026-02-06 12:13:00"; // Updated after rebuild
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ BUILD TIMESTAMP: {buildTimestamp}\n");
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC 2: PlaceBulkClusters method ENTERED\n");
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🚀 TRUE BULK PLACEMENT START: {clusters.Count} clusters\n");
            
            int totalPlaced = 0;
            var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
            var clusterMap = new List<BatchClusterData>();
            
            // ✅ DEDUPLICATION: Track placement locations to prevent duplicates within this batch
            var placedLocations = new HashSet<string>(); // Key: "{X:F6}_{Y:F6}_{Z:F6}"
            double tolerance = 0.1; // 0.1 feet tolerance for duplicate detection
            int skippedDuplicateCount = 0;
            int skippedExistingCount = 0;

            // 1. Prepare Creation Data
            using (placementTracker?.TrackSubOperation("Prepare Creation Data & Deduplication"))
            {
                foreach (var cluster in clusters)
            {
                if (!symbolCache.TryGetValue(cluster.FamilyName, out var symbol)) continue;
                if (!symbol.IsActive) symbol.Activate();

                XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                
                // ✅ DEDUPLICATION: Check for duplicate placement points within this batch
                string locationKey = $"{cluster.PlacementX:F6}_{cluster.PlacementY:F6}_{cluster.PlacementZ:F6}";
                if (placedLocations.Contains(locationKey))
                {
                    skippedDuplicateCount++;
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING DUPLICATE CLUSTER in bulk placement: ClusterGUID={cluster.ClusterGUID}, Location=({cluster.PlacementX:F6}, {cluster.PlacementY:F6}, {cluster.PlacementZ:F6})\n");
                    continue;
                }
                
                // Note: NewFamilyInstances2 doesn't always handle hosts perfectly for all family types,
                // but we'll try to use the host and level if available.
                Autodesk.Revit.Creation.FamilyInstanceCreationData data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, StructuralType.NonStructural);

                    creationDataList.Add(data);
                    clusterMap.Add(cluster);
                    placedLocations.Add(locationKey); // Mark this location as used
                }
            }

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📊 BULK PLACEMENT SUMMARY: Input={clusters.Count}, Added to creation list={creationDataList.Count}, Skipped duplicates={skippedDuplicateCount}, Skipped existing={skippedExistingCount}\n");
            
            if (creationDataList.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ No clusters to place after deduplication\n");
                return 0;
            }

            // ✅ CRITICAL FIX: Delete individual sleeves BEFORE placing clusters to prevent duplicates
            // Collect all individual sleeves that need to be deleted
            var allIndividualSleevesToDelete = new List<ElementId>();
            var clusterToSleeveIdsMap = new Dictionary<int, List<ElementId>>(); // Track which sleeves belong to which cluster
            
            using (placementTracker?.TrackSubOperation("Collect Individual Sleeves to Delete"))
            {
                foreach (var cluster in clusterMap)
                {
                    if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                    
                    var guids = cluster.ConstituentZoneGuids
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(g => g.Trim())
                        .Where(g => !string.IsNullOrEmpty(g) && Guid.TryParse(g, out _))
                        .Select(g => Guid.Parse(g))
                        .ToList();
                    
                    if (!guids.Any()) continue;
                    
                    var sleeveIdsForThisCluster = new List<ElementId>();
                    
                    foreach (var guid in guids)
                    {
                        // ✅ PHASE 8: Use zoneCache instead of DB query
                        if (zoneCache.TryGetValue(guid, out var z))
                        {
                            // ✅ ROBUST PERSISTENCE: ALWAYS try to recover from snapshot first
                            int effectiveSleeveId = _repository.TryGetSleeveInstanceIdFromSnapshot(z.Id);
                            if (effectiveSleeveId <= 0)
                            {
                                effectiveSleeveId = z.SleeveInstanceId;
                            }
                            
                            if (effectiveSleeveId > 0)
                            {
                                var elementId = new ElementId(effectiveSleeveId);
                                // Verify element exists before adding to delete list
                                if (doc.GetElement(elementId) != null)
                                {
                                    allIndividualSleevesToDelete.Add(elementId);
                                    sleeveIdsForThisCluster.Add(elementId);
                                }
                            }
                        }
                    }
                    
                    // Store mapping for later use in PerformSwapDeletion
                    if (sleeveIdsForThisCluster.Any())
                    {
                        clusterToSleeveIdsMap[clusterMap.IndexOf(cluster)] = sleeveIdsForThisCluster;
                    }
                }
            }
            
            // ✅ DELETE INDIVIDUAL SLEEVES BEFORE PLACING CLUSTERS
            if (allIndividualSleevesToDelete.Any())
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🗑️ DELETING {allIndividualSleevesToDelete.Count} individual sleeves BEFORE cluster placement to prevent duplicates\n");
                
                try
                {
                    using (placementTracker?.TrackSubOperation("Delete Individual Sleeves (Before Cluster Placement)"))
                    {
                        doc.Delete(allIndividualSleevesToDelete.Distinct().ToList()); // Bulk delete, remove duplicates
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ DELETED {allIndividualSleevesToDelete.Count} individual sleeves before cluster placement\n");
                    }
                }
                catch (Exception delEx)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Warning: Failed to delete some individual sleeves before cluster placement: {delEx.Message}\n");
                }
            }

            // 2. Execute Bulk Placement
            ICollection<ElementId> placedIds;
            using (placementTracker?.TrackSubOperation("Step 2a: Revit NewFamilyInstances2"))
            {
                placedIds = doc.Create.NewFamilyInstances2(creationDataList);
            }

            var placedIdList = placedIds.ToList();
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ Bulk Placement Result: {placedIdList.Count} instances created\n");

            // 3. Post-Placement Logic (Rotation, Parameters, Cleanup)
            using (placementTracker?.TrackSubOperation("Step 3: APPLY ROTATION & PARAMS"))
            {
                for (int i = 0; i < placedIdList.Count; i++)
                {
                    var eid = placedIdList[i];
                    var cluster = clusterMap[i];
                    var instance = doc.GetElement(eid) as FamilyInstance;

                    if (instance == null) continue;

                    int clusterInstanceId = instance.Id.IntegerValue;
                    placedInstances.Add(instance);

                    // A. Rotation
                    if (Math.Abs(cluster.RotationAngleRad) > 1e-6)
                    {
                        XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                        Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
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

                    bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || 
                                     cluster.FamilyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;

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
                        
                        _parameterService.SetClusterSleeveInstanceId(instance, clusterInstanceId);
                        _parameterService.SetSleeveInstanceId(instance, -1);
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Bulk Param Setting failed: {ex.Message}\n");
                    }
                }
            }
            
            // C. Database & Cleanup Stage 1
            using (placementTracker?.TrackSubOperation("Step 6: SAVE PLACED DATA TO DB"))
            {
                // ✅ PERFORMANCE OPTIMIZATION: Batch database saves instead of sequential
                // Pre-fetch all ComboIds and FilterIds, then save all clusters in one batch operation
                
                var clusterSaveDataList = new List<JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSaveData>();
                var comboIdMap = new Dictionary<int, (int comboId, int filterId, string category, string hostType, string hostOrientation)>();
                
                // ✅ STAGE 2 CLEANUP FIX: Track GUIDs for ALL successfully placed instances
                var placedGuidMap = new Dictionary<string, int>(); 
                // Step 1: Pre-fetch ComboIds and FilterIds for all clusters
                using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(doc))
                {
                    var allComboIds = new HashSet<int>();
                    
                    // Collect all unique ComboIds from first zone of each cluster
                    for (int i = 0; i < placedIdList.Count; i++)
                    {
                        var cluster = clusterMap[i];
                        if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                        {
                            var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                            if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid firstGuid))
                            {
                                // ✅ PHASE 8: Use zoneCache instead of DB query
                                if (zoneCache.TryGetValue(firstGuid, out var firstZone))
                                {
                                    if (firstZone.ComboId > 0)
                                    {
                                        allComboIds.Add(firstZone.ComboId);
                                        comboIdMap[i] = (
                                            firstZone.ComboId,
                                            -1, // FilterId will be fetched next
                                            firstZone.MepElementCategory ?? "",
                                            firstZone.StructuralElementType ?? "",
                                            firstZone.HostOrientation ?? ""
                                        );
                                    }
                                }
                            }
                        }
                    }
                    
                    // Batch fetch all FilterIds in one query
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
                                    int cid = reader.GetInt32(0);
                                    int fid = reader.GetInt32(1);
                                    filterIdMap[cid] = fid;
                                }
                            }
                        }
                    }
                    
                    // Update comboIdMap with fetched FilterIds
                    var updatedMap = new Dictionary<int, (int, int, string, string, string)>();
                    foreach (var kvp in comboIdMap)
                    {
                        var (comboId, _, category, hostType, hostOrientation) = kvp.Value;
                        int filterId = filterIdMap.ContainsKey(comboId) ? filterIdMap[comboId] : -1;
                        updatedMap[kvp.Key] = (comboId, filterId, category, hostType, hostOrientation);
                    }
                    comboIdMap = updatedMap;
                }
                
                // Step 2: Prepare ClusterSaveData for all clusters
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC 3: Starting batch save preparation loop for {placedIdList.Count} clusters\n");
                for (int i = 0; i < placedIdList.Count; i++)
                {
                    var eid = placedIdList[i];
                    var cluster = clusterMap[i];
                    var instance = doc.GetElement(eid) as FamilyInstance;

                    if (instance == null) continue;

                    int clusterInstanceId = instance.Id.IntegerValue;

                    try
                    {
                        // ✅ CRITICAL FIX: Skip deletion (already done before cluster placement) but still update flags
                        PerformSwapDeletion(doc, cluster, clusterInstanceId, skipPlacementPointUpdate: false, skipDeletion: true);
                        
                        // ✅ STAGE 2 CLEANUP FIX: Record GUID mapping for status update later
                        if (!string.IsNullOrEmpty(cluster.ClusterGUID))
                        {
                            placedGuidMap[cluster.ClusterGUID] = clusterInstanceId;
                        }
                        
                        // Get pre-fetched ComboId and FilterId
                        if (!comboIdMap.ContainsKey(i))
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Warning: No ComboId/FilterId found for cluster {clusterInstanceId}, skipping save\n");
                            continue;
                        }
                        
                        var (comboId, filterId, category, hostType, hostOrientation) = comboIdMap[i];
                        
                        if (comboId <= 0 || filterId <= 0)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌ Invalid ComboId ({comboId}) or FilterId ({filterId}) for cluster {clusterInstanceId}\n");
                            continue;
                        }
                        
                        // Extract corners from instance
                        var cornerService = new JSE_RevitAddin_MEP_OPENINGS.Services.Geometry.SleeveCornerCalculationService();
                        var corners = cornerService.CalculateCornersFromInstance(instance);
                        
                        if (!corners.HasValue)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Warning: Could not extract corners for cluster {clusterInstanceId}\n");
                            continue;
                        }
                        
                        // Extract parameters from instance
                        double width = (instance.LookupParameter("Width") ?? instance.LookupParameter("Element Width"))?.AsDouble() ?? cluster.ClusterWidth;
                        double height = (instance.LookupParameter("Height") ?? instance.LookupParameter("Element Height"))?.AsDouble() ?? cluster.ClusterHeight;
                        double depth = (instance.LookupParameter("Depth") ?? instance.LookupParameter("Element Depth") ?? instance.LookupParameter("Wall Width"))?.AsDouble() ?? cluster.ClusterDepth;
                        
                        var loc = instance.Location as LocationPoint;
                        double px = loc?.Point.X ?? cluster.PlacementX;
                        double py = loc?.Point.Y ?? cluster.PlacementY;
                        double pz = loc?.Point.Z ?? cluster.PlacementZ;
                        
                        // Parse zone GUIDs
                        var zoneGuids = cluster.ConstituentZoneGuids
                            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(g => Guid.Parse(g.Trim()))
                            .ToList();
                        
                        // Create ClusterSaveData
                        var saveData = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSaveData
                        {
                            ClusterInstanceId = clusterInstanceId,
                            ComboId = comboId,
                            FilterId = filterId,
                            Category = category,
                            BoundingBoxMinX = 0, // Not used by calculations that use corners
                            BoundingBoxMinY = 0,
                            BoundingBoxMinZ = 0,
                            BoundingBoxMaxX = 0,
                            BoundingBoxMaxY = 0,
                            BoundingBoxMaxZ = 0,
                            ClusterWidth = width,
                            ClusterHeight = height,
                            ClusterDepth = depth,
                            RotationAngleDeg = cluster.RotationAngleRad * (180.0 / Math.PI),
                            IsRotated = Math.Abs(cluster.RotationAngleRad) > 1e-6,
                            PlacementX = px,
                            PlacementY = py,
                            PlacementZ = pz,
                            HostType = hostType,
                            HostOrientation = hostOrientation,
                            ClashZoneIds = zoneGuids,
                            SleeveFamilyName = cluster.FamilyName,
                            Corner1X = corners.Value.corner1.X,
                            Corner1Y = corners.Value.corner1.Y,
                            Corner1Z = corners.Value.corner1.Z,
                            Corner2X = corners.Value.corner2.X,
                            Corner2Y = corners.Value.corner2.Y,
                            Corner2Z = corners.Value.corner2.Z,
                            Corner3X = corners.Value.corner3.X,
                            Corner3Y = corners.Value.corner3.Y,
                            Corner3Z = corners.Value.corner3.Z,
                            Corner4X = corners.Value.corner4.X,
                            Corner4Y = corners.Value.corner4.Y,
                            Corner4Z = corners.Value.corner4.Z
                        };
                        
                        clusterSaveDataList.Add(saveData);
                        totalPlaced++;
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ Failed to prepare save data for cluster {cluster.ClusterGUID}: {ex.Message}\n");
                    }
                }
                
                // Step 3: Batch save all clusters in one operation
                if (clusterSaveDataList.Any())
                {
                    try
                    {
                        using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(doc))
                        {
                            var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveRepository(
                                dbContext, 
                                msg => SafeFileLogger.SafeAppendText("batch_v2.log", msg));
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC 4: About to execute BATCH SAVE\n");
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] 💾 BATCH SAVE: Saving {clusterSaveDataList.Count} clusters in one operation\n");
                            
                            repo.BatchSaveClusterSleeves(clusterSaveDataList);
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVE COMPLETE: {clusterSaveDataList.Count} clusters saved\n");

                            // ✅ PHASE 8: Optimized Batch Status Update
                            UpdateStatusBatch(placedGuidMap, "Placed");
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH STATUS UPDATE COMPLETE: Status set to 'Placed' for {placedGuidMap.Count} clusters (Single Transaction)\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Batch save failed: {ex.Message}\n{ex.StackTrace}\n");
                        throw;
                    }
                }
            }

            return totalPlaced;
        }

        /// <summary>
        /// ✅ CRITICAL: Update cluster bounding boxes in database AFTER parameter flush and regeneration
        /// Updates both ClusterSleeves table and ClashZones.ClusterSleeveBoundingBox columns
        /// This ensures Stage 2 cleanup can find overlapping sleeves using DB-only queries
        /// </summary>
        private void UpdateClusterBoundingBoxesAfterPlacement(Document doc, List<FamilyInstance> placedInstances)
        {
            try
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 💾 BATCH BBOX UPDATE: Starting update for {placedInstances.Count} instances\n");
                
                using (var context = new SleeveDbContext(doc))
                {
                    context.Connection.Open();
                    using (var trans = context.Connection.BeginTransaction())
                    {
                        int updatedCount = 0;
                        foreach (var instance in placedInstances)
                        {
                            if (instance == null || !instance.IsValidObject) continue;
                            
                            int clusterInstanceId = instance.Id.IntegerValue;
                            var bbox = instance.get_BoundingBox(null);
                            if (bbox == null || !bbox.Enabled) continue;
                            
                            using (var cmd = context.Connection.CreateCommand())
                            {
                                cmd.Transaction = trans;
                                cmd.CommandText = @"
                                    UPDATE ClusterSleeves_v2 
                                    SET BoundingBoxMinX = @MinX,
                                        BoundingBoxMinY = @MinY,
                                        BoundingBoxMinZ = @MinZ,
                                        BoundingBoxMaxX = @MaxX,
                                        BoundingBoxMaxY = @MaxY,
                                        BoundingBoxMaxZ = @MaxZ
                                    WHERE ClusterInstanceId = @ClusterId;

                                    -- ✅ STURDY PERSISTENCE: Sync BBox to legacy table
                                    UPDATE ClusterSleeves 
                                    SET BoundingBoxMinX = @MinX,
                                        BoundingBoxMinY = @MinY,
                                        BoundingBoxMinZ = @MinZ,
                                        BoundingBoxMaxX = @MaxX,
                                        BoundingBoxMaxY = @MaxY,
                                        BoundingBoxMaxZ = @MaxZ
                                    WHERE ClusterInstanceId = @ClusterId";
                                
                                cmd.Parameters.AddWithValue("@ClusterId", clusterInstanceId);
                                cmd.Parameters.AddWithValue("@MinX", bbox.Min.X);
                                cmd.Parameters.AddWithValue("@MinY", bbox.Min.Y);
                                cmd.Parameters.AddWithValue("@MinZ", bbox.Min.Z);
                                cmd.Parameters.AddWithValue("@MaxX", bbox.Max.X);
                                cmd.Parameters.AddWithValue("@MaxY", bbox.Max.Y);
                                cmd.Parameters.AddWithValue("@MaxZ", bbox.Max.Z);
                                
                                updatedCount += (cmd.ExecuteNonQuery() > 0 ? 1 : 0);
                            }
                        }
                        trans.Commit();
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH BBOX UPDATE COMPLETE: Updated {updatedCount} clusters in BOTH v2 and legacy tables (Single Transaction)\n");
                    }
                }
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
        public string FamilyName { get; set; } = string.Empty;
        public string ConstituentZoneGuids { get; set; } = string.Empty;
    }
}
