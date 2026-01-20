using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
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

        public (int placed, int failed) PlaceFromDatabase(Document doc, string batchId, bool useSingleTransaction = true)
        {
            using (var tracker = _performanceMonitor?.TrackOperation("Cluster Placement Total"))
            {
                int placedCount = 0;
                int failedCount = 0;
                var placedInstances = new List<FamilyInstance>(); // Track placed instances for 2nd cleanup

                // 1. Query Pending Clusters
                List<BatchClusterData> pendingClusters;
                using (_performanceMonitor?.TrackOperation("Query Pending Clusters"))
                {
                    pendingClusters = GetPendingClusters(batchId);
                }
                
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ STARTING PLACEMENT V2: Batch {batchId}, Pending={pendingClusters.Count}, Mode={(useSingleTransaction ? "BULK" : "SEQUENTIAL")}\n");

                if (pendingClusters.Count == 0) return (0, 0);

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
                        if (isNestedTransaction)
                        {
                            // MODE A-1: BULK (Existing Transaction)
                            foreach (var cluster in pendingClusters)
                            {
                                if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                else failedCount++;
                            }
                        }
                        else
                        {
                            // MODE A-2: BULK (New Transaction)
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

                // ✅ CRITICAL FIX: Flush batched parameters
                if (OptimizationFlags.UseBatchedParameterWrites && placedCount > 0)
                {
                    using (var flushTracker = _performanceMonitor?.TrackOperation("Set Cluster Parameters (Flush)"))
                    {
                        var clusterFlushTimer = System.Diagnostics.Stopwatch.StartNew();
                        int flushedCount = 0;
                        
                        try
                        {
                            flushedCount = _parameterService.FlushDeferredParameters(clearList: true);
                            doc.Regenerate();
                            clusterFlushTimer.Stop();

                            SafeFileLogger.SafeAppendText("performance.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER-FLUSH] Clusters={placedCount}, Parameters={flushedCount}, Time={clusterFlushTimer.ElapsedMilliseconds}ms\n");
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FLUSH FAILED: {ex.Message}\n");
                        }
                        
                        flushTracker?.SetItemCount(flushedCount);
                    }
                }
                
                // ✅ 2ND CLEANUP PASS: Delete individual sleeves within placed cluster bounding boxes
                // This matches sequential clustering behavior
                // IMPORTANT: This must run AFTER parameter flush (if enabled) and regeneration
                if (_cleanupService != null && placedInstances.Count > 0)
                {
                    using (var cleanupTracker = _performanceMonitor?.TrackOperation("Cleanup Individual Sleeves (2nd Pass)"))
                    {
                        try
                        {
                            // Extract target category from batchId (format: "20260117_180808_Ducts_3")
                            string targetCategory = null;
                            var batchParts = batchId.Split('_');
                            if (batchParts.Length >= 3)
                            {
                                targetCategory = batchParts[2]; // "Ducts", "Pipes", etc.
                            }
                            
                            int deletedCount = _cleanupService.CleanupSleevesWithinClusters(
                                doc, 
                                placedInstances, 
                                deferredParameters: null, // Batch mode doesn't use deferred parameters for cleanup
                                targetCategory: targetCategory);
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 2ND CLEANUP: Deleted {deletedCount} individual sleeves overlapping with {placedInstances.Count} cluster sleeves (Category={targetCategory ?? "ALL"})\n");
                            
                            cleanupTracker?.SetItemCount(deletedCount);
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ 2ND CLEANUP FAILED: {ex.Message}\n");
                        }
                    }
                }
                tracker?.SetItemCount(placedCount);
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH COMPLETED: Placed={placedCount}, Failed={failedCount}\n");
                return (placedCount, failedCount);
            }
        }





        private List<BatchClusterData> GetPendingClusters(string batchId)
        {
            var list = new List<BatchClusterData>();
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ClusterGUID, PlacementX, PlacementY, PlacementZ, 
                               ClusterWidth, ClusterHeight, ClusterDepth, RotationAngleRad,
                               HostElementId, FamilyName, ConstituentZoneGuids
                        FROM ClusterSleeves_v2
                        WHERE (ClusterBatchId = @batch OR @batch IS NULL) AND Status = 'Pending'";
                    // Allow batchId to be null to fetch ALL pending
                    if (string.IsNullOrEmpty(batchId))
                        cmd.Parameters.AddWithValue("@batch", DBNull.Value);
                    else
                        cmd.Parameters.AddWithValue("@batch", batchId);
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new BatchClusterData
                            {
                                ClusterGUID = reader["ClusterGUID"].ToString(),
                                PlacementX = Convert.ToDouble(reader["PlacementX"]),
                                PlacementY = Convert.ToDouble(reader["PlacementY"]),
                                PlacementZ = Convert.ToDouble(reader["PlacementZ"]),
                                ClusterWidth = Convert.ToDouble(reader["ClusterWidth"]),
                                ClusterHeight = Convert.ToDouble(reader["ClusterHeight"]),
                                ClusterDepth = Convert.ToDouble(reader["ClusterDepth"]),
                                RotationAngleRad = Convert.ToDouble(reader["RotationAngleRad"]),
                                HostElementId = Convert.ToInt64(reader["HostElementId"]),
                                FamilyName = reader["FamilyName"].ToString(),
                                ConstituentZoneGuids = reader["ConstituentZoneGuids"].ToString()
                            });
                        }
                    }
                }
            }
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

        private void PerformSwapDeletion(Document doc, BatchClusterData cluster, int clusterElementId)
        {
            if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) return;

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
            var oldClusterInstanceIds = zones
                .Where(z => z.ClusterInstanceId > 0 && z.ClusterInstanceId != clusterElementId)
                .Select(z => z.ClusterInstanceId)
                .Distinct()
                .ToList();

            if (oldClusterInstanceIds.Any())
            {
                DeleteOldClusterSleeves(oldClusterInstanceIds);
            }

            var toDeleteIds = new List<ElementId>();
            var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)>();


            foreach (var z in zones)
            {
                // Delete from Revit
                if (z.SleeveInstanceId > 0)
                {
                    toDeleteIds.Add(new ElementId(z.SleeveInstanceId));
                }

                // ✅ FIX: Update DB with FULL cluster data (flags + calculated columns)
                // USER REQUEST: SleeveInstanceId must be -1 for clustered sleeves
                // USER REQUEST: Populate IsClusteredFlag, MarkedForClusterProcess, and AfterClusterSleeveId
                updates.Add((
                    z.Id,                      // ClashZoneId
                    true,                       // IsResolved
                    true,                       // IsClusterResolved
                    false,                      // IsCombinedResolved
                    -1,                         // SleeveInstanceId (Explicitly -1)
                    clusterElementId,           // ClusterInstanceId
                    true,                       // IsClusteredFlag
                    true,                       // MarkedForClusterProcess
                    clusterElementId            // AfterClusterSleeveId
                ));
            }

            if (toDeleteIds.Any())
            {
                try
                {
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

            // ✅ FIX 1: Update Flags in DB (existing functionality)
            _repository.BatchUpdateFlags(updates);
            
            // ✅ FIX 2: Update ClashZones with calculated cluster dimensions
            UpdateClashZonesCalculatedColumns(guids, cluster, clusterElementId);
            
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] ✅ UPDATED {guids.Count} zones with cluster data (ClusterInstanceId={clusterElementId})\n");
        }

        /// <summary>
        /// ✅ FIX: Update ClashZones table with calculated cluster dimensions
        /// Populates: CalculatedSleeveWidth, Height, Depth, Rotation, FamilyName, PlacedAt
        /// </summary>
        private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId)
        {
            // 🔍 LOG BLOCK 1: Entry point
            SafeFileLogger.SafeAppendText("debug_db.log", 
                $"\n[{DateTime.Now:HH:mm:ss}] === UpdateClashZonesCalculatedColumns START ===\n" +
                $"    ZoneGuids Count: {zoneGuids?.Count ?? 0}\n" +
                $"    ClusterInstanceId: {clusterInstanceId}\n" +
                $"    Width: {cluster.ClusterWidth:F4}, Height: {cluster.ClusterHeight:F4}\n");
            
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
                            cmd.CommandText = $@"
                                UPDATE ClashZones 
                                SET 
                                    CalculatedSleeveWidth = @Width,
                                    CalculatedSleeveHeight = @Height,
                                    CalculatedSleeveDepth = @Depth,
                                    CalculatedRotation = @Rotation,
                                    CalculatedFamilyName = @FamilyName,
                                    PlacedAt = CURRENT_TIMESTAMP,
                                    PlacementStatus = 'Placed',
                                    ClusterInstanceId = @ClusterInstanceId,
                                    IsClusterResolvedFlag = 1,
                                    SleeveState = 2,
                                    UpdatedAt = CURRENT_TIMESTAMP
                                WHERE ClashZoneGuid IN ({string.Join(", ", guidParams)})";
                            
                            cmd.Parameters.AddWithValue("@Width", cluster.ClusterWidth);
                            cmd.Parameters.AddWithValue("@Height", cluster.ClusterHeight);
                            cmd.Parameters.AddWithValue("@Depth", cluster.ClusterDepth);
                            cmd.Parameters.AddWithValue("@Rotation", cluster.RotationAngleRad);
                            cmd.Parameters.AddWithValue("@FamilyName", cluster.FamilyName ?? "");
                            cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                            
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
                                $"[{DateTime.Now:HH:mm:ss}] 📊 UPDATED {rowsAffected} ClashZones calculated columns (Width={cluster.ClusterWidth:F3}, Height={cluster.ClusterHeight:F3})\n");
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
        private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
        {
            // ✅ DIAGNOSTIC: Log method entry
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"\n[{DateTime.Now:HH:mm:ss}] 🔴 SaveToClusterSleevesLegacy CALLED!\n" +
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

                                // ✅ DIAGNOSTIC: Log before SaveClusterSleeve call
                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     ✅ IDs valid, about to call SaveClusterSleeve\n");

                                repo.SaveClusterSleeve(
                                    clusterInstanceId: clusterInstanceId,
                                    comboId: comboId,
                                    filterId: filterId,
                                    category: category,
                                    boundingBoxMinX: cluster.PlacementX - halfWidth,
                                    boundingBoxMinY: cluster.PlacementY - halfHeight,
                                    boundingBoxMinZ: cluster.PlacementZ - halfDepth,
                                    boundingBoxMaxX: cluster.PlacementX + halfWidth,
                                    boundingBoxMaxY: cluster.PlacementY + halfHeight,
                                    boundingBoxMaxZ: cluster.PlacementZ + halfDepth,
                                    clusterWidth: cluster.ClusterWidth,
                                    clusterHeight: cluster.ClusterHeight,
                                    clusterDepth: cluster.ClusterDepth,
                                    rotationAngleDeg: cluster.RotationAngleRad * (180.0 / Math.PI), // Convert to degrees
                                    isRotated: Math.Abs(cluster.RotationAngleRad) > 1e-6,
                                    placementX: cluster.PlacementX,
                                    placementY: cluster.PlacementY,
                                    placementZ: cluster.PlacementZ,
                                    hostType: hostType,
                                    hostOrientation: hostOrientation,
                                    clashZoneIds: zoneGuids,
                                    sleeveFamilyName: cluster.FamilyName,
                                    // ✅ CRITICAL: Calculate actual corners for combined sleeve width calculation
                                    // Corners are in world coordinates, rotated around placement center
                                    corner1X: cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner1Y: cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner1Z: cluster.PlacementZ,
                                    corner2X: cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner2Y: cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner2Z: cluster.PlacementZ,
                                    corner3X: cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner3Y: cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner3Z: cluster.PlacementZ,
                            corner4X: cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner4Y: cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner4Z: cluster.PlacementZ
                                );

                                // ✅ DIAGNOSTIC: Log after SaveClusterSleeve returns
                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     ✅ SaveClusterSleeve RETURNED\n");

                                // 🔍 LOG BLOCK 4: SaveToClusterSleevesLegacy Success
                                SafeFileLogger.SafeAppendText("debug_db.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 💾 SAVED to ClusterSleeves (legacy): ClusterInstanceId={clusterInstanceId}, ComboId={comboId}\n");

                                SafeFileLogger.SafeAppendText("batch_v2.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 💾 SAVED to ClusterSleeves (legacy): ClusterInstanceId={clusterInstanceId}, ComboId={comboId}\n");
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
                        using (parentTracker != null ? parentTracker.TrackSubOperation("Swap and DB Update Cluster") : _performanceMonitor?.TrackOperation("Swap and DB Update Cluster"))
                        {
                            // F. SWAP LOGIC
                            PerformSwapDeletion(doc, cluster, clusterInstanceId);
                            // G. Update ClusterSleeves_v2 Status
                            UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
                            // H. Save to ClusterSleeves (legacy) table
                            SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ DB OPERATION FAILED for Cluster {cluster.ClusterGUID}: {ex.Message}\n");
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
        /// Legacy Cleanup: Deletes old ClusterSleeves by Instance IDs.
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
                            cmd.CommandText = $"DELETE FROM ClusterSleeves WHERE ClusterInstanceId IN ({ids})";
                            int rows = cmd.ExecuteNonQuery();
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🗑️ DELETED {rows} old cluster rows from ClusterSleeves (Ids: {ids}) to prevent duplication.\n");
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
        public string FamilyName { get; set; }
        public string ConstituentZoneGuids { get; set; }
    }
}
