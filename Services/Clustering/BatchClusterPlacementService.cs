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
            // This prevents tracking time when no clusters need to be placed
            List<BatchClusterData> pendingClusters;
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: About to call GetPendingClusters with batchId={(batchId ?? "NULL")}\n");
            using (_performanceMonitor?.TrackOperation("Query Pending Clusters"))
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
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📊 TABLE: Will save placed clusters to ClusterSleeves table (legacy) after placement\n");
            
            using (var tracker = _performanceMonitor?.TrackOperation("Cluster Placement Total"))
            {
                int placedCount = 0;
                int failedCount = 0;
                var placedInstances = new List<FamilyInstance>(); // Track placed instances for 2nd cleanup

                // ✅ OPTIMIZATION: Pre-fetch all required Family Symbols
                var symbolCache = new Dictionary<string, FamilySymbol>();
                using (_performanceMonitor?.TrackOperation("Cache Family Symbols"))
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
                using (_performanceMonitor?.TrackOperation("Pre-fetch Constituent Zones"))
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

                // Check for existing transaction (Revit specific)
                bool isNestedTransaction = doc.IsModifiable;

                // ✅ REF: Aggregated Logging Context
                using (var placementTracker = _performanceMonitor?.TrackOperation("Place Cluster Instances"))
                {
                    if (useSingleTransaction || isNestedTransaction)
                    {
                        if (OptimizationFlags.UseBulkClusterSleevePlacement && !isNestedTransaction)
                        {
                            // MODE A-3: TRUE BULK PLACEMENT (New Transaction)
                            // This uses doc.Create.NewFamilyInstances2 for maximum performance
                            using (Transaction t = new Transaction(doc, "Place Batch Clusters V2 (TRUE BULK)"))
                            {
                                t.Start();
                                placedCount = PlaceBulkClusters(doc, pendingClusters, symbolCache, zoneCache, placementTracker, placedInstances);
                                t.Commit();
                            }
                        }
                        else if (isNestedTransaction)
                        {
                            // MODE A-1: BULK (Existing Transaction - Sequential Loops)
                            foreach (var cluster in pendingClusters)
                            {
                                if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                else failedCount++;
                            }
                        }
                        else
                        {
                            // MODE A-2: BULK (New Transaction - Sequential Loops)
                            using (Transaction t = new Transaction(doc, "Place Batch Clusters V2 (Bulk)"))
                            {
                                t.Start();
                                foreach (var cluster in pendingClusters)
                                {
                                    if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                    else failedCount++;
                                }
                                t.Commit();
                            }
                        }
                    }
                    else
                    {
                        // MODE B: SEQUENTIAL (Transaction per Cluster)
                        foreach (var cluster in pendingClusters)
                        {
                            using (Transaction t = new Transaction(doc, $"Place Cluster {cluster.ClusterGUID}"))
                            {
                                try
                                {
                                    t.Start();
                                    if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances))
                                    {
                                        placedCount++;
                                        t.Commit();
                                    }
                                    else
                                    {
                                        failedCount++;
                                        t.RollBack();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ TRANS ERROR: {ex.Message}\n");
                                    failedCount++;
                                }
                            }
                        }
                    }
                    placementTracker?.SetItemCount(placedCount);
                }

                // ✅ CRITICAL FIX: Flush batched parameters - MUST be inside a transaction
                // BUT: Don't create a new transaction if we're already in one (nested transaction error)
                if (OptimizationFlags.UseBatchedParameterWrites && placedCount > 0)
                {
                    using (var flushTracker = _performanceMonitor?.TrackOperation("Set Cluster Parameters (Flush)"))
                    {
                        var clusterFlushTimer = System.Diagnostics.Stopwatch.StartNew();
                        int flushedCount = 0;
                        
                        try
                        {
                            // ✅ FIX: Check if already in transaction to avoid nested transaction error
                            bool alreadyInTransaction = doc.IsModifiable;
                            
                            if (alreadyInTransaction)
                            {
                                // We're inside an existing transaction - flush directly
                                flushedCount = _parameterService.FlushDeferredParameters(clearList: true);
                                doc.Regenerate();
                            }
                            else
                            {
                                // No active transaction - create one for flush
                                using (Transaction flushTx = new Transaction(doc, "Flush Cluster Parameters"))
                                {
                                    flushTx.Start();
                                    flushedCount = _parameterService.FlushDeferredParameters(clearList: true);
                                    doc.Regenerate();
                                    flushTx.Commit();
                                }
                            }
                            clusterFlushTimer.Stop();

                            SafeFileLogger.SafeAppendText("performance.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER-FLUSH] Clusters={placedCount}, Parameters={flushedCount}, Time={clusterFlushTimer.ElapsedMilliseconds}ms, InExistingTx={alreadyInTransaction}\n");
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FLUSH FAILED: {ex.Message}\n");
                        }
                        
                        flushTracker?.SetItemCount(flushedCount);
                    }
                }
                
                // ✅ CRITICAL: Update cluster bounding boxes in database AFTER parameter flush and regeneration
                // This ensures bounding boxes are accurate and saved to both ClusterSleeves and ClashZones tables
                // Stage 2 cleanup needs ClusterSleeveBoundingBoxMin/MaxX/Y/Z columns to find overlapping sleeves
                if (placedInstances.Count > 0)
                {
                    using (var bboxTracker = _performanceMonitor?.TrackOperation("Update Cluster Bounding Boxes"))
                    {
                        try
                        {
                            UpdateClusterBoundingBoxesAfterPlacement(doc, placedInstances);
                            bboxTracker?.SetItemCount(placedInstances.Count);
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to update cluster bounding boxes: {ex.Message}\n{ex.StackTrace}\n");
                        }
                    }
                }
                
                // ✅ CLEANUP REMOVED: Cleanup is now handled at orchestrator level after ALL categories are placed
                // This allows cleanup to run once for all categories instead of per-category
                // The placedInstances list is still populated for potential future use
                else
                {
                    if (_cleanupService == null)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP SKIPPED: CleanupService is null\n");
                    }
                    else if (placedInstances.Count == 0)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP SKIPPED: No placed instances (placedInstances.Count=0)\n");
                    }
                }
                tracker?.SetItemCount(placedCount);
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH COMPLETED: Placed={placedCount}, Failed={failedCount}\n");
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
                    var totalCount = Convert.ToInt32(countCmd.ExecuteScalar());
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Total clusters in ClusterSleeves (legacy): {totalCount}\n");
                }
                
                // ✅ DIAGNOSTIC: Check clusters needing placement count from ClusterSleeves
                using (var pendingCmd = conn.CreateCommand())
                {
                    pendingCmd.CommandText = @"SELECT COUNT(*) FROM ClusterSleeves 
                        WHERE ClusterInstanceId IS NOT NULL";
                    var totalWithInstanceId = Convert.ToInt32(pendingCmd.ExecuteScalar());
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Clusters in ClusterSleeves with ClusterInstanceId: {totalWithInstanceId}\n");
                }
                
                using (var cmd = conn.CreateCommand())
                {
                    // ✅ FIX: Query from ClusterSleeves table (legacy) where ClusterInstanceId is saved
                    // Join with ClusterSleeves_v2 to get additional metadata if needed
                    cmd.CommandText = @"
                        SELECT 
                            cs.ClusterGUID, 
                            cs.PlacementX, cs.PlacementY, cs.PlacementZ, 
                            cs.ClusterWidth, cs.ClusterHeight, cs.ClusterDepth, 
                            cs.RotationAngleRad,
                            cs.HostElementId, cs.HostType, cs.HostOrientation, 
                            cs.FamilyName, cs.ConstituentZoneGuids,
                            cs.ClusterInstanceId
                        FROM ClusterSleeves cs
                        WHERE cs.ClusterInstanceId IS NOT NULL";
                    // Note: ClusterSleeves table doesn't have batchId, so we fetch all
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Querying ClusterSleeves table (legacy) for clusters with ClusterInstanceId\n");
                    
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
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING: Cluster {clusterGuid} has invalid ClusterInstanceId={clusterInstanceId}\n");
                                continue;
                            }
                            
                            // ✅ CRITICAL: Check if cluster actually exists in Revit
                            bool existsInRevit = existenceChecker.ClusterSleeveExists(clusterInstanceId);
                            if (!existsInRevit)
                            {
                                // Cluster was placed before but doesn't exist now (was deleted) - needs re-placement
                                needsPlacementCount++;
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 PATH 1: Cluster {clusterGuid} has ClusterInstanceId={clusterInstanceId} but doesn't exist in Revit - will re-place\n");
                                
                                list.Add(new BatchClusterData
                                {
                                    ClusterGUID = clusterGuid,
                                    PlacementX = Convert.ToDouble(reader["PlacementX"]),
                                    PlacementY = Convert.ToDouble(reader["PlacementY"]),
                                    PlacementZ = Convert.ToDouble(reader["PlacementZ"]),
                                    ClusterWidth = Convert.ToDouble(reader["ClusterWidth"]),
                                    ClusterHeight = Convert.ToDouble(reader["ClusterHeight"]),
                                    ClusterDepth = Convert.ToDouble(reader["ClusterDepth"]),
                                    RotationAngleRad = Convert.ToDouble(reader["RotationAngleRad"]),
                                    HostElementId = Convert.ToInt64(reader["HostElementId"]),
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
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ PATH 1: Cluster {clusterGuid} with ClusterInstanceId={clusterInstanceId} already exists in Revit - skipping\n");
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
                                    sleeveFamilyName: cluster.FamilyName,
                                    actualInstance: actualInstance
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

                // B. Determine Level
                Element host = null;
                Level level = null;
                if (cluster.HostElementId > 0)
                {
                    try { host = doc.GetElement(new ElementId((int)cluster.HostElementId)); } catch { }
                    if (host != null) level = doc.GetElement(host.LevelId) as Level;
                }
                if (level == null) level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

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
                    if (host != null) instance = doc.Create.NewFamilyInstance(location, symbol, host, level, structuralType);
                    else instance = doc.Create.NewFamilyInstance(location, symbol, structuralType);
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
                
                // ✅ DEDUPLICATION: Check if cluster already exists at this location (when DB is not cleared)
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
                    skippedExistingCount++;
                    var existingId = existingClusters.First().Id.IntegerValue;
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING CLUSTER - already exists at location: ClusterGUID={cluster.ClusterGUID}, ExistingId={existingId}, Location=({cluster.PlacementX:F6}, {cluster.PlacementY:F6}, {cluster.PlacementZ:F6})\n");
                    continue;
                }
                
                // Note: NewFamilyInstances2 doesn't always handle hosts perfectly for all family types,
                // but we'll try to use the host and level if available.
                Element host = null;
                Level level = null;
                if (cluster.HostElementId > 0)
                {
                    try { host = doc.GetElement(new ElementId((int)cluster.HostElementId)); } catch { }
                    if (host != null) level = doc.GetElement(host.LevelId) as Level;
                }
                if (level == null) level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

                Autodesk.Revit.Creation.FamilyInstanceCreationData data = null;
                if (host != null && level != null)
                    data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, host, level, StructuralType.NonStructural);
                else
                    data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, StructuralType.NonStructural);

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
                
                var zones = _repository.GetClashZonesByGuids(guids);
                var sleeveIdsForThisCluster = new List<ElementId>();
                
                foreach (var z in zones)
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
                
                // Store mapping for later use in PerformSwapDeletion
                if (sleeveIdsForThisCluster.Any())
                {
                    // We'll set the clusterInstanceId after placement
                    // For now, use a temporary key based on cluster index
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
            using (placementTracker?.TrackSubOperation("Revit NewFamilyInstances2 (Clusters)"))
            {
                placedIds = doc.Create.NewFamilyInstances2(creationDataList);
            }

            var placedIdList = placedIds.ToList();
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ Bulk Placement Result: {placedIdList.Count} instances created\n");

            // 3. Post-Placement Logic (Rotation, Parameters, Cleanup)
            using (placementTracker?.TrackSubOperation("Post-Placement: Rotation & Parameters"))
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
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Bulk Param Setting failed: {ex.Message}\n");
                    }
                }
            }
            
            // C. Database & Cleanup Stage 1
            using (placementTracker?.TrackSubOperation("Post-Placement: Database Updates"))
            {
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
                        UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
                        SafeFileLogger.SafeAppendText("batch_v2.log",
                            $"[{DateTime.Now:HH:mm:ss}] 📊 TABLE: About to save cluster {clusterInstanceId} to ClusterSleeves table (legacy)\n");
                        SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId, instance);
                        totalPlaced++;
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ Bulk DB Update failed for {cluster.ClusterGUID}: {ex.Message}\n");
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
                    $"[{DateTime.Now:HH:mm:ss}] 💾 Starting cluster bounding box update for {placedInstances.Count} instances\n");
                
                using (var context = new SleeveDbContext(doc))
                {
                    var clusterRepo = new ClusterSleeveRepository(context, msg => SafeFileLogger.SafeAppendText("batch_v2.log", $"[ClusterRepo] {msg}\n"));
                    var clashZoneRepo = new ClashZoneRepository(context, msg => SafeFileLogger.SafeAppendText("batch_v2.log", $"[ClashZoneRepo] {msg}\n"));
                    
                    int clusterSleevesUpdated = 0;
                    int clashZonesUpdated = 0;
                    
                    foreach (var instance in placedInstances)
                    {
                        if (instance == null || !instance.IsValidObject) continue;
                        
                        int clusterInstanceId = instance.Id.IntegerValue;
                        
                        // ✅ Get bounding box from Revit element (after regeneration, this should be accurate)
                        var bbox = instance.get_BoundingBox(null);
                        if (bbox == null || !bbox.Enabled)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cluster {clusterInstanceId} has no valid bounding box\n");
                            continue;
                        }
                        
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 📦 Cluster {clusterInstanceId} bbox: Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}), Max=({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})\n");
                        
                        // Step 1: Update ClusterSleeves table
                        try
                        {
                            using (var cmd = context.Connection.CreateCommand())
                            {
                                cmd.CommandText = @"
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
                                
                                int rowsAffected = cmd.ExecuteNonQuery();
                                if (rowsAffected > 0)
                                {
                                    clusterSleevesUpdated++;
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ Updated ClusterSleeves table for cluster {clusterInstanceId}\n");
                                }
                                else
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ No ClusterSleeves record found for cluster {clusterInstanceId}\n");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to update ClusterSleeves for {clusterInstanceId}: {ex.Message}\n");
                        }
                        
                        // Step 2: SKIP - Cluster bounding boxes are already in ClusterSleeves table
                        // No need to duplicate them in ClashZones - cleanup service queries ClusterSleeves directly
                        // This was removed to avoid unnecessary data duplication
                        /*
                        try
                        {
                            // First, get constituent zone GUIDs for this cluster
                            var constituentGuids = new List<string>();
                            using (var queryCmd = context.Connection.CreateCommand())
                            {
                                // Try ClusterSleeves_v2 first (new table)
                                queryCmd.CommandText = @"
                                    SELECT ConstituentZoneGuids 
                                    FROM ClusterSleeves_v2 
                                    WHERE ClusterInstanceId = @ClusterId
                                    LIMIT 1";
                                queryCmd.Parameters.AddWithValue("@ClusterId", clusterInstanceId);
                                
                                var result = queryCmd.ExecuteScalar();
                                if (result != null && result != DBNull.Value)
                                {
                                    var guidString = result.ToString();
                                    if (!string.IsNullOrEmpty(guidString))
                                    {
                                        constituentGuids = guidString.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                            .Select(g => g.Trim())
                                            .ToList();
                                    }
                                }
                                
                                // Fallback to ClusterSleeves table if not found
                                if (constituentGuids.Count == 0)
                                {
                                    queryCmd.Parameters.Clear();
                                    queryCmd.CommandText = @"
                                        SELECT ClashZoneGuids 
                                        FROM ClusterSleeves 
                                        WHERE ClusterInstanceId = @ClusterId
                                        LIMIT 1";
                                    queryCmd.Parameters.AddWithValue("@ClusterId", clusterInstanceId);
                                    
                                    result = queryCmd.ExecuteScalar();
                                    if (result != null && result != DBNull.Value)
                                    {
                                        var guidString = result.ToString();
                                        if (!string.IsNullOrEmpty(guidString))
                                        {
                                            // ClashZoneGuids might be JSON array or comma-separated
                                            if (guidString.TrimStart().StartsWith("["))
                                            {
                                                // JSON array - parse it
                                                try
                                                {
                                                    var jsonArray = System.Text.Json.JsonSerializer.Deserialize<List<string>>(guidString);
                                                    if (jsonArray != null) constituentGuids = jsonArray;
                                                }
                                                catch { }
                                            }
                                            else
                                            {
                                                // Comma-separated
                                                constituentGuids = guidString.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(g => g.Trim())
                                                    .ToList();
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (constituentGuids.Count > 0)
                            {
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Found {constituentGuids.Count} constituent GUIDs for cluster {clusterInstanceId}\n");
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: GUIDs = {string.Join(", ", constituentGuids)}\n");
                                
                                // ✅ DIAGNOSTIC: First check if GUIDs exist in database BEFORE attempting UPDATE
                                using (var checkCmd = context.Connection.CreateCommand())
                                {
                                    var checkPlaceholders = string.Join(",", constituentGuids.Select((_, i) => $"@CheckGuid{i}"));
                                    checkCmd.CommandText = $@"
                                        SELECT COUNT(*) 
                                        FROM ClashZones 
                                        WHERE UPPER(ClashZoneGuid) IN ({checkPlaceholders})";
                                    for (int i = 0; i < constituentGuids.Count; i++)
                                    {
                                        checkCmd.Parameters.AddWithValue($"@CheckGuid{i}", constituentGuids[i].ToUpperInvariant());
                                    }
                                    var guidCount = Convert.ToInt32(checkCmd.ExecuteScalar());
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: {guidCount} of {constituentGuids.Count} GUIDs found in ClashZones table BEFORE UPDATE\n");
                                    
                                    if (guidCount == 0)
                                    {
                                        // Try without UPPER() to see if case is the issue
                                        checkCmd.Parameters.Clear();
                                        checkCmd.CommandText = $@"
                                            SELECT COUNT(*) 
                                            FROM ClashZones 
                                            WHERE ClashZoneGuid IN ({checkPlaceholders})";
                                        for (int i = 0; i < constituentGuids.Count; i++)
                                        {
                                            checkCmd.Parameters.AddWithValue($"@CheckGuid{i}", constituentGuids[i]);
                                        }
                                        var guidCountExact = Convert.ToInt32(checkCmd.ExecuteScalar());
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: {guidCountExact} of {constituentGuids.Count} GUIDs found with EXACT case match\n");
                                        
                                        // Show sample GUID from database for comparison
                                        checkCmd.Parameters.Clear();
                                        checkCmd.CommandText = "SELECT ClashZoneGuid FROM ClashZones WHERE ClashZoneGuid IS NOT NULL LIMIT 1";
                                        var sampleGuid = checkCmd.ExecuteScalar()?.ToString();
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Sample GUID from DB: {sampleGuid}\n");
                                    }
                                }
                                
                                // Update zones by GUID (more reliable)
                                try
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: About to execute UPDATE query for cluster {clusterInstanceId}\n");
                                    
                                    using (var cmd = context.Connection.CreateCommand())
                                    {
                                        var guidPlaceholders = string.Join(",", constituentGuids.Select((_, i) => $"@Guid{i}"));
                                        cmd.CommandText = $@"
                                            UPDATE ClashZones 
                                            SET ClusterSleeveBoundingBoxMinX = @MinX,
                                                ClusterSleeveBoundingBoxMinY = @MinY,
                                                ClusterSleeveBoundingBoxMinZ = @MinZ,
                                                ClusterSleeveBoundingBoxMaxX = @MaxX,
                                                ClusterSleeveBoundingBoxMaxY = @MaxY,
                                                ClusterSleeveBoundingBoxMaxZ = @MaxZ
                                            WHERE UPPER(ClashZoneGuid) IN ({guidPlaceholders})";
                                        
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: UPDATE SQL prepared with {constituentGuids.Count} GUID placeholders\n");
                                        
                                        for (int i = 0; i < constituentGuids.Count; i++)
                                        {
                                            cmd.Parameters.AddWithValue($"@Guid{i}", constituentGuids[i].ToUpperInvariant());
                                        }
                                        
                                        cmd.Parameters.AddWithValue("@MinX", bbox.Min.X);
                                        cmd.Parameters.AddWithValue("@MinY", bbox.Min.Y);
                                        cmd.Parameters.AddWithValue("@MinZ", bbox.Min.Z);
                                        cmd.Parameters.AddWithValue("@MaxX", bbox.Max.X);
                                        cmd.Parameters.AddWithValue("@MaxY", bbox.Max.Y);
                                        cmd.Parameters.AddWithValue("@MaxZ", bbox.Max.Z);
                                        
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: About to call ExecuteNonQuery()\n");
                                        
                                        int rowsAffected = cmd.ExecuteNonQuery();
                                        
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: UPDATE query returned {rowsAffected} rows affected\n");
                                        
                                        if (rowsAffected > 0)
                                        {
                                            clashZonesUpdated += rowsAffected;
                                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                                $"[{DateTime.Now:HH:mm:ss}] ✅ Updated {rowsAffected} ClashZones with cluster bounding box for cluster {clusterInstanceId} (via {constituentGuids.Count} GUIDs)\n");
                                        }
                                        else
                                        {
                                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ No ClashZones updated for cluster {clusterInstanceId} (found {constituentGuids.Count} GUIDs but 0 rows matched)\n");
                                        }
                                    }
                                }
                                catch (Exception updateEx)
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ EXCEPTION in UPDATE query for cluster {clusterInstanceId}: {updateEx.Message}\n{updateEx.StackTrace}\n");
                                    throw; // Re-throw to be caught by outer try-catch
                                }
                            }
                            else
                            {
                                // Fallback: Try by ClusterInstanceId (in case GUIDs not available)
                                // ✅ DIAGNOSTIC: First check if any rows exist with this ClusterInstanceId
                                using (var checkCmd = context.Connection.CreateCommand())
                                {
                                    checkCmd.CommandText = @"
                                        SELECT COUNT(*) 
                                        FROM ClashZones 
                                        WHERE ClusterInstanceId = @ClusterId";
                                    checkCmd.Parameters.AddWithValue("@ClusterId", clusterInstanceId);
                                    var count = Convert.ToInt32(checkCmd.ExecuteScalar());
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Found {count} ClashZones with ClusterInstanceId={clusterInstanceId}\n");
                                }
                                
                                using (var cmd = context.Connection.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                        UPDATE ClashZones 
                                        SET ClusterSleeveBoundingBoxMinX = @MinX,
                                            ClusterSleeveBoundingBoxMinY = @MinY,
                                            ClusterSleeveBoundingBoxMinZ = @MinZ,
                                            ClusterSleeveBoundingBoxMaxX = @MaxX,
                                            ClusterSleeveBoundingBoxMaxY = @MaxY,
                                            ClusterSleeveBoundingBoxMaxZ = @MaxZ
                                        WHERE ClusterInstanceId = @ClusterId";
                                    
                                    cmd.Parameters.AddWithValue("@ClusterId", clusterInstanceId);
                                    cmd.Parameters.AddWithValue("@MinX", bbox.Min.X);
                                    cmd.Parameters.AddWithValue("@MinY", bbox.Min.Y);
                                    cmd.Parameters.AddWithValue("@MinZ", bbox.Min.Z);
                                    cmd.Parameters.AddWithValue("@MaxX", bbox.Max.X);
                                    cmd.Parameters.AddWithValue("@MaxY", bbox.Max.Y);
                                    cmd.Parameters.AddWithValue("@MaxZ", bbox.Max.Z);
                                    
                                    int rowsAffected = cmd.ExecuteNonQuery();
                                    if (rowsAffected > 0)
                                    {
                                        clashZonesUpdated += rowsAffected;
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ Updated {rowsAffected} ClashZones with cluster bounding box for cluster {clusterInstanceId} (via ClusterInstanceId fallback)\n");
                                    }
                                    else
                                    {
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ UPDATE by ClusterInstanceId returned 0 rows for cluster {clusterInstanceId}\n");
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to update ClashZones for cluster {clusterInstanceId}: {ex.Message}\n{ex.StackTrace}\n");
                        }
                        */
                    }
                    
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 💾 Bounding box update complete: {clusterSleevesUpdated} ClusterSleeves, {clashZonesUpdated} ClashZones updated\n");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ UpdateClusterBoundingBoxesAfterPlacement failed: {ex.Message}\n{ex.StackTrace}\n");
            }
        }
    }

    public class BatchClusterData
    {
        public string ClusterGUID { get; set; }
        public double PlacementX { get; set; }
        public double PlacementY { get; set; }
        public double PlacementZ { get; set; }
        public double ClusterWidth { get; set; }
        public double ClusterHeight { get; set; }
        public double ClusterDepth { get; set; }
        public double RotationAngleRad { get; set; }
        public long HostElementId { get; set; }
        public string HostType { get; set; }
        public string HostOrientation { get; set; }
        public string FamilyName { get; set; }
        public string ConstituentZoneGuids { get; set; }
    }
}
