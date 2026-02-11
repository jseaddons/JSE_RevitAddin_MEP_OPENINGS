using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// Service for Phase 1: Pure Calculation of Clusters (No Revit Creation) - V2 Architecture
    /// Responsibilities:
    /// 1. Group ClashZones (Bucketing/Grouping)
    /// 2. Run Clustering Algorithm (Proximity)
    /// 3. Calculate Position (Rotation/Centroid)
    /// 4. Deduplicate (Location-Based)
    /// 5. Save to DB (Status='Pending')
    /// </summary>
    public class BatchClusterCalculationService
    {
        private readonly IClusterAlgorithmService _algorithmService;
        private readonly IClusterRotationService _rotationService;
        private readonly string _databasePath;

        // Shared Deduplication Cache (Batch-Scope)
        // Key: "{BatchId}_{X:F1}_{Y:F1}_{Z:F1}" -> Prevents stacking at same location in same batch
        private readonly ConcurrentDictionary<string, byte> _placedLocations = new ConcurrentDictionary<string, byte>();

        public BatchClusterCalculationService(
            IClusterAlgorithmService algorithmService,
            IClusterRotationService rotationService,
            string databasePath)
        {
            _algorithmService = algorithmService ?? throw new ArgumentNullException(nameof(algorithmService));
            _rotationService = rotationService ?? throw new ArgumentNullException(nameof(rotationService));
            _databasePath = databasePath;
        }

        /// <summary>
        /// ✅ PHASE 1: PURE CALCULATION (CPU Bound, Thread Safe)
        /// Calculates clusters without touching the database (except for reading).
        /// Returns calculated results in memory.
        /// </summary>
        public List<BatchClusterCalculationResult> CalculateOnly(
            List<ClashZone> clashZones, 
            string targetCategory, 
            int comboId, 
            int filterId,
            Document doc)
        {
            // ✅ PHASE 10: EARLY EXIT
            if (clashZones == null || clashZones.Count == 0)
                return new List<BatchClusterCalculationResult>();

            // 1. Generate Batch ID
            string batchId = $"{DateTime.Now:yyyyMMdd_HHmmss}_{targetCategory}_{filterId}";
            
            // 🔍 DIAGNOSTIC: Log batch ID generation
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 CALCULATE ONLY: Generated batchId={batchId} for {clashZones.Count} zones, Category={targetCategory}\n");

            var settings = ApplicationProfileService.Instance.GetCurrentSettings();
            double toleranceMM = settings.JoinOpeningsDistance;
            var toleranceDist = RevitUnitConversionService.Instance.ToInternalMillimeters(toleranceMM);
            // Floor/pipes often form clusters of 2 ("in and out") because one pipe through a slab = two penetrations
            // within tolerance; other pipes are further away. To get larger floor/pipe clusters, increase JoinOpeningsDistance.
            
            // 2. Group Zones (Optimized Phase 10)
            // Pre-compute keys once to avoid string operations in GroupBy loop
            var wrappedZones = new List<ClashZoneWorkItem>(clashZones.Count);
            foreach (var z in clashZones)
            {
                var hostType = GetBatchHostType(z);
                var groupKey = new SleeveGroupKey(
                    $"{hostType}_{z.StructuralElementIdValue}",
                    z.MepElementCategory ?? "Unknown", 
                    z.HostOrientation ?? "Unknown",
                    0, 0, 0);

                wrappedZones.Add(new ClashZoneWorkItem
                {
                    SleeveInstanceId = z.SleeveInstanceId,
                    ClashZone = z,
                    GroupKey = groupKey
                });
            }
            
            var groupedZones = wrappedZones.GroupBy(w => w.GroupKey);
            
            var clustersByGroup = _algorithmService.FormClusters(groupedZones, toleranceDist, doc, enableParallel: true);

            // ✅ CRITICAL PERFORMANCE FIX: Preload ClashZone cache BEFORE parallel execution
            // Without this, each parallel thread will hit the database via _getClashZoneFunc inside CalculateRotatedBoundingBox,
            // causing SQLite lock contention and serializing what should be parallel work.
            var clashZoneDict = clashZones.ToDictionary(z => z.SleeveInstanceId, z => z);
            int preloadedCount = _rotationService.PreloadClashZonesFromDictionary(clashZoneDict);
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] ✅ CACHE PRELOAD: Loaded {preloadedCount} ClashZones into rotation service cache before parallel calculation\n");
            }

            // 4. Process Results (Parallel Calculation)
            var validClusters = new ConcurrentBag<BatchClusterCalculationResult>();

            System.Threading.Tasks.Parallel.ForEach(clustersByGroup, (kvp) =>
            {
                foreach (var clusterList in kvp.Value)
                {
                    if (clusterList == null || clusterList.Count < 2) continue;
                    
                    try
                    {
                        // Branch logic by Host Type EARLY
                        var first = (clusterList[0] is ClashZone z) ? z : (ClashZone)((ClashZoneWorkItem)clusterList[0]).ClashZone;
                        bool isFloor = (first.StructuralElementType ?? "").IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0;

                        BatchClusterCalculationResult result;
                        if (isFloor)
                        {
                            result = CalculateFloorCluster(clusterList, batchId, comboId, filterId);
                        }
                        else
                        {
                            result = CalculateWallFramingCluster(clusterList, batchId, comboId, filterId);
                        }

                        if (result != null)
                        {
                            validClusters.Add(result);
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ CALC ERROR: {ex.Message}\n");
                    }
                }
            });

            return validClusters.ToList();
        }

        /// <summary>
        /// ✅ PHASE 2: BATCH SAVE (IO Bound, Single Transaction)
        /// Saves a list of calculated results to the database.
        /// Publicly exposed for Orchestrator to call after parallel calculation.
        /// </summary>
        public void BatchSave(IEnumerable<BatchClusterCalculationResult> results)
        {
             // Delegate to existing private method, but make sure it handles generic IEnumerable
             var bag = new ConcurrentBag<BatchClusterCalculationResult>(results);
             
             // ✅ USE EXISTING BATCH ID: Don't create a new one, use what was set during calculation
             // The batch ID was already set in CalculateOnly (e.g., "20260206_180119_Cable Trays_0")
             string batchId = bag.FirstOrDefault()?.ClusterBatchId ?? "UNKNOWN_BATCH";
             
             SafeFileLogger.SafeAppendText("batch_v2.log", 
                 $"[{DateTime.Now:HH:mm:ss}] 📦 BATCH SAVE: Saving {bag.Count} clusters with batchId={batchId}\n");
             
             SaveToDatabase(bag, batchId); 
        }

        /// <summary>
        /// ✅ PHASE 3: FLAG UPDATE (IO Bound, Single Transaction)
        /// Updates ClashZone flags for clustered items.
        /// </summary>
        public void BatchUpdateFlags(IEnumerable<BatchClusterCalculationResult> results)
        {
             var bag = new ConcurrentBag<BatchClusterCalculationResult>(results);
             UpdateMarkedForClusterProcessFlags(bag);
        }

        // DEPRECATED: Old combined method (kept for compatibility if needed, or can be removed)
        public string CalculateAndSave(
            List<ClashZone> clashZones, 
            string targetCategory, 
            int comboId, 
            int filterId,
            Document doc)
        {
            var results = CalculateOnly(clashZones, targetCategory, comboId, filterId, doc);
            BatchSave(results);
            BatchUpdateFlags(results);
            
            return results.FirstOrDefault()?.ClusterBatchId ?? $"{DateTime.Now:yyyyMMdd_HHmmss}_{targetCategory}_{filterId}";
        }
        
        /// <summary>
        /// ✅ PATH 1 FAST PATH: Check if cluster data already exists in database
        /// Returns existing batchId if clusters exist, null otherwise
        /// </summary>
        private string CheckExistingClusterData(int comboId, int filterId, string category)
        {
            try
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        // ✅ PATH 1: Check if clusters exist for this combo/filter/category
                        // Use Path 1 if clusters exist (either Pending or Placed)
                        // Placement will check if they need to be placed (Status='Pending' OR Status='Placed' with ClusterInstanceId=0)
                        cmd.CommandText = @"
                            SELECT DISTINCT ClusterBatchId 
                            FROM ClusterSleeves_v2 
                            WHERE ComboId = @ComboId 
                            AND FilterId = @FilterId 
                            AND Category = @Category
                            AND (Status = 'Pending' OR Status = 'Placed')
                            LIMIT 1";
                        
                        cmd.Parameters.AddWithValue("@ComboId", comboId);
                        cmd.Parameters.AddWithValue("@FilterId", filterId);
                        cmd.Parameters.AddWithValue("@Category", category);
                        
                        var result = cmd.ExecuteScalar();
                        if (result != null && !string.IsNullOrEmpty(result.ToString()))
                        {
                            return result.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error checking existing cluster data: {ex.Message} - will proceed with calculation\n");
            }
            
            return null;
        }




        private void SaveToDatabase(ConcurrentBag<BatchClusterCalculationResult> results, string batchId)
        {
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: SaveToDatabase called with {results.Count} clusters, batchId={batchId}\n");
            
            if (results.Count == 0) return;

            try
            {
                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        // ✅ OPTIMIZATION Step 1: Pre-fetch existing locations to avoid N database reads
                        // Instead of checking each cluster individually, load all relevant existing locations into memory
                        var existingLocations = new HashSet<string>();
                        using (var checkCmd = conn.CreateCommand())
                        {
                            checkCmd.Transaction = trans;
                            checkCmd.CommandText = @"
                                SELECT PlacementX, PlacementY, PlacementZ, FamilyName 
                                FROM ClusterSleeves_v2 
                                WHERE Status IN ('Pending', 'Placed')";
                            
                            using (var reader = checkCmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    double x = reader.GetDouble(0);
                                    double y = reader.GetDouble(1);
                                    double z = reader.GetDouble(2);
                                    string fam = reader.IsDBNull(3) ? "" : reader.GetString(3);
                                    // Use same tolerance formatting as check logic (F1 or roughly 0.1)
                                    // Actually, let's use a spatial key with 0.1 tolerance
                                    existingLocations.Add(GetLocationKey(x, y, z, fam));
                                }
                            }
                        }

                        // Filter results in memory
                        var infoToSave = new List<BatchClusterCalculationResult>();
                        int skippedCount = 0;
                        
                        foreach (var r in results)
                        {
                            string key = GetLocationKey(r.PlacementX, r.PlacementY, r.PlacementZ, r.FamilyName ?? "");
                            if (existingLocations.Contains(key))
                            {
                                skippedCount++;
                            }
                            else
                            {
                                infoToSave.Add(r);
                                // Add to set to prevents duplicates WITHIN this batch too
                                existingLocations.Add(key); 
                            }
                        }

                        if (skippedCount > 0)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: Skipped {skippedCount} duplicate clusters (already exist in DB or batch)\n");
                        }

                        if (infoToSave.Count == 0)
                        {
                             SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: No unique clusters to save after filtering.\n");
                             trans.Commit();
                             return;
                        }

                        // ✅ OPTIMIZATION Step 2: Batch Insert
                        // SQLite max variables = 999. We have ~19 params per row. 
                        // Safe chunk size = 50 rows (~950 params).
                        const int CHUNK_SIZE = 50;
                        int savedCount = 0;
                        
                        for (int i = 0; i < infoToSave.Count; i += CHUNK_SIZE)
                        {
                            var chunk = infoToSave.Skip(i).Take(CHUNK_SIZE).ToList();
                            
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = trans;
                                var valueClauses = new List<string>();
                                
                                for (int k = 0; k < chunk.Count; k++)
                                {
                                    var r = chunk[k];
                                    string p = $"@p{k}_"; // Prefix for this row's params
                                    
                                    valueClauses.Add($"({p}guid, {p}batch, {p}x, {p}y, {p}z, {p}w, {p}h, {p}d, {p}rot, {p}host, {p}htype, {p}horient, {p}cat, {p}fam, {p}zones, {p}combo, {p}filter, {p}status, {p}valid)");
                                    
                                    // 🔍 DIAGNOSTIC: Log what we're about to INSERT (ClusterBatchId = timestamp-style, same for whole batch)
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔍 SQL INSERT #{k}: ClusterBatchId={r.ClusterBatchId}, GUID={r.ClusterGUID.Substring(0,8)}, " +
                                        $"Placement=({r.PlacementX:F3},{r.PlacementY:F3},{r.PlacementZ:F3}), " +
                                        $"Size=({r.ClusterWidth:F3}x{r.ClusterHeight:F3}x{r.ClusterDepth:F3})\n");
                                    
                                    cmd.Parameters.AddWithValue($"{p}guid", r.ClusterGUID);
                                    cmd.Parameters.AddWithValue($"{p}batch", r.ClusterBatchId);
                                    cmd.Parameters.AddWithValue($"{p}x", r.PlacementX);
                                    cmd.Parameters.AddWithValue($"{p}y", r.PlacementY);
                                    cmd.Parameters.AddWithValue($"{p}z", r.PlacementZ);
                                    cmd.Parameters.AddWithValue($"{p}w", r.ClusterWidth);
                                    cmd.Parameters.AddWithValue($"{p}h", r.ClusterHeight);
                                    cmd.Parameters.AddWithValue($"{p}d", r.ClusterDepth);
                                    cmd.Parameters.AddWithValue($"{p}rot", r.RotationAngleRad);
                                    cmd.Parameters.AddWithValue($"{p}host", r.HostElementId);
                                    cmd.Parameters.AddWithValue($"{p}htype", r.HostType ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue($"{p}horient", r.HostOrientation ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue($"{p}cat", r.Category ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue($"{p}fam", r.FamilyName ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue($"{p}zones", r.ConstituentZoneGuids ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue($"{p}combo", r.ComboId);
                                    cmd.Parameters.AddWithValue($"{p}filter", r.FilterId);
                                    cmd.Parameters.AddWithValue($"{p}status", r.Status ?? "Pending");
                                    cmd.Parameters.AddWithValue($"{p}valid", r.ValidationStatus ?? "Valid");
                                }
                                
                                try 
                                {
                                    // BATCH MODE: Execute the big INSERT string
                                    // This is the "Fast Path" that mimics "Slow Mode" but in one go
                                    cmd.CommandText = @"
                                    INSERT INTO ClusterSleeves_v2 (
                                        ClusterGUID, ClusterBatchId, PlacementX, PlacementY, PlacementZ, 
                                        ClusterWidth, ClusterHeight, ClusterDepth, RotationAngleRad,
                                        HostElementId, HostType, HostOrientation, Category, FamilyName, 
                                        ConstituentZoneGuids, ComboId, FilterId, Status, ValidationStatus
                                    ) VALUES " + string.Join(",", valueClauses) + ";";

                                    int rowsAffected = cmd.ExecuteNonQuery();
                                    savedCount += rowsAffected;
                                    
                                    SafeFileLogger.SafeAppendText("batch_sql_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SQL SUCCESS: Inserted {rowsAffected} rows into ClusterSleeves_v2 (Batch of {results.Count}).\n");
                                }
                                catch (Exception sqlEx)
                                {
                                    SafeFileLogger.SafeAppendText("batch_sql_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ❌ BATCH SQL ERROR: {sqlEx.Message}\n" +
                                        $"SQL: {cmd.CommandText.Substring(0, Math.Min(500, cmd.CommandText.Length))}...\n");
                                    throw; // Re-throw to trigger rollback
                                }
                            }

                            // ✅ RESTORE LEGACY PERSISTENCE: Sync to ClusterSleeves (Legacy)
                            // This ensures the legacy table (used by previous logic steps) also contains the pending clusters.
                            try 
                            {
                                using (var legCmd = conn.CreateCommand())
                                {
                                    legCmd.Transaction = trans;
                                    var legClauses = new List<string>();
                                    
                                    for (int k = 0; k < chunk.Count; k++)
                                    {
                                        var r = chunk[k];
                                        string p = $"@l{k}_";
                                        
                                        double halfW = r.ClusterWidth / 2.0;
                                        double halfH = r.ClusterHeight / 2.0;
                                        double halfD = r.ClusterDepth / 2.0;
                                        
                                        legClauses.Add($"({p}id, {p}guid, {p}combo, {p}filt, {p}cat, {p}minx, {p}miny, {p}minz, {p}maxx, {p}maxy, {p}maxz, {p}w, {p}h, {p}d, {p}rot, {p}isrot, {p}px, {p}py, {p}pz, {p}ht, {p}ho, {p}guids, {p}json)");

                                        legCmd.Parameters.AddWithValue($"{p}id", -1); // Pending = -1
                                        legCmd.Parameters.AddWithValue($"{p}guid", r.ClusterGUID ?? Guid.NewGuid().ToString()); // ✅ FIX: Persist ClusterGUID
                                        legCmd.Parameters.AddWithValue($"{p}combo", r.ComboId);
                                        legCmd.Parameters.AddWithValue($"{p}filt", r.FilterId);
                                        legCmd.Parameters.AddWithValue($"{p}cat", r.Category ?? "");
                                        legCmd.Parameters.AddWithValue($"{p}minx", r.PlacementX - halfW);
                                        legCmd.Parameters.AddWithValue($"{p}miny", r.PlacementY - halfH);
                                        legCmd.Parameters.AddWithValue($"{p}minz", r.PlacementZ - halfD);
                                        legCmd.Parameters.AddWithValue($"{p}maxx", r.PlacementX + halfW);
                                        legCmd.Parameters.AddWithValue($"{p}maxy", r.PlacementY + halfH);
                                        legCmd.Parameters.AddWithValue($"{p}maxz", r.PlacementZ + halfD);
                                        legCmd.Parameters.AddWithValue($"{p}w", r.ClusterWidth);
                                        legCmd.Parameters.AddWithValue($"{p}h", r.ClusterHeight);
                                        legCmd.Parameters.AddWithValue($"{p}d", r.ClusterDepth);
                                        legCmd.Parameters.AddWithValue($"{p}rot", r.RotationAngleRad * 57.2958); // Rad to Deg
                                        legCmd.Parameters.AddWithValue($"{p}isrot", Math.Abs(r.RotationAngleRad) > 0.01 ? 1 : 0);
                                        legCmd.Parameters.AddWithValue($"{p}px", r.PlacementX);
                                        legCmd.Parameters.AddWithValue($"{p}py", r.PlacementY);
                                        legCmd.Parameters.AddWithValue($"{p}pz", r.PlacementZ);
                                        legCmd.Parameters.AddWithValue($"{p}ht", r.HostType ?? "");
                                        legCmd.Parameters.AddWithValue($"{p}ho", r.HostOrientation ?? "");
                                        legCmd.Parameters.AddWithValue($"{p}guids", r.ConstituentZoneGuids ?? "");
                                        
                                        // ✅ FIX: Populate JSON correctly for legacy readers
                                        var jsonStr = "[]";
                                        if (!string.IsNullOrEmpty(r.ConstituentZoneGuids))
                                        {
                                            var gList = r.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                .Select(g => $"\"{g.Trim()}\"");
                                            jsonStr = $"[{string.Join(",", gList)}]";
                                        }
                                        legCmd.Parameters.AddWithValue($"{p}json", jsonStr);
                                    }
                                    
                                    legCmd.CommandText = @"
                                        INSERT INTO ClusterSleeves (
                                            ClusterInstanceId, ClusterGuid, ComboId, FilterId, Category,
                                            BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                            BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                            ClusterWidth, ClusterHeight, ClusterDepth,
                                            RotationAngleDeg, IsRotated,
                                            PlacementX, PlacementY, PlacementZ,
                                            HostType, HostOrientation,
                                            ClashZoneGuids, ClashZoneIdsJson
                                        ) VALUES " + string.Join(",", legClauses) + ";";
                                    
                                    legCmd.ExecuteNonQuery();
                                }
                            }
                            catch (Exception legEx)
                            {
                                SafeFileLogger.SafeAppendText("batch_sql_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ LEGACY SQL IGNORED: {legEx.Message}\n");
                            }
                        }
                        
                        trans.Commit();
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ TABLE: Bulk Inserted {savedCount} clusters into ClusterSleeves_v2\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ TABLE CRITICAL ERROR: SaveToDatabase failed: {ex.Message}\n{ex.StackTrace}\n");
                throw;
            }
        }
        
        private string GetLocationKey(double x, double y, double z, string family)
        {
            // Simple quantization to 1 decimal place (approx 0.1 ft) for key matching
            return $"{x:F1}_{y:F1}_{z:F1}_{family}";
        }

        private SleeveGroupKey[] GroupZones(List<ClashZone> zones)
        {
            // Reusing existing grouping logic (Host, Orientation, 10m Buckets)
            var groups = zones
                .GroupBy(z =>
                {
                    // ✅ FIX: Disable spatial bucketing as requested
                    return new SleeveGroupKey(
                        z.StructuralElementIdValue.ToString(), 
                        z.MepElementCategory ?? "Unknown", 
                        z.HostOrientation ?? "Unknown",
                        0, 0, 0);
                })
                .Select(g => 
                {
                    // Convert ClashZones to 'dynamic' explicitly
                    var items = g.Select(z => (dynamic)z).ToList();
                    var key = g.Key;
                    
                    // ✅ DIAGNOSTIC: Log large groups (>50 items) which trigger O(N^2) slowness
                    if (items.Count > 50)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ LARGE GROUP DETECTED: Host={key.hostType}, Cat={key.systemType}, Orient={key.orientation}, Count={items.Count} items. This may cause slow clustering.\n");
                    }
                    
                    // SleeveGroupKey doesn't have Items property, return as-is
                    return key;
                })
                .ToArray();
                
            return groups;
        }

        /// <summary>
        /// ✅ SPECIALIZED FLOOR CALCULATION: Strictly separated from wall logic.
        /// Floor rotation uses same DetermineRotationAngle as RefactoredClusterService (MepElementRotationAngle for rectangular, 0 for circular).
        /// </summary>
        private BatchClusterCalculationResult CalculateFloorCluster(List<dynamic> clusterItems, string batchId, int comboId, int filterId)
        {
            var zones = clusterItems.Select(x => (x is ClashZone z) ? z : ((ClashZoneWorkItem)x).ClashZone).ToList();
            if (zones.Count == 0) return null;

            var first = zones[0];
            // ✅ FLOOR ROTATION: Use ClusterRotationService so rotated MEP on floors gets correct bbox (same path as individual sleeves)
            double rotationAngle = _rotationService.DetermineRotationAngle(clusterItems);

            var bboxResult = _rotationService.CalculateRotatedBoundingBox(clusterItems, null, rotationAngle);
            
            // 🔍 DIAGNOSTIC: Log what CalculateRotatedBoundingBox returned
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 BBOX CALC RESULT: Zones={zones.Count}, " +
                $"MinX={bboxResult.rotatedMinX}, MinY={bboxResult.rotatedMinY}, MinZ={bboxResult.rotatedMinZ}, " +
                $"MaxX={bboxResult.rotatedMaxX}, MaxY={bboxResult.rotatedMaxY}, MaxZ={bboxResult.rotatedMaxZ}, " +
                $"Width={bboxResult.width}, Height={bboxResult.height}, Depth={bboxResult.depth}\n");
            
            // ✅ CRITICAL VALIDATION: Reject clusters with invalid bounding boxes
            if (bboxResult.rotatedMinX == null || bboxResult.rotatedMinY == null || bboxResult.rotatedMinZ == null ||
                bboxResult.rotatedMaxX == null || bboxResult.rotatedMaxY == null || bboxResult.rotatedMaxZ == null)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ BBOX VALIDATION FAILED: Cluster has null bbox values. " +
                    $"Zones={zones.Count}, Category={first.MepElementCategory}, HostType={first.StructuralElementType}, " +
                    $"ZoneGuids=[{string.Join(",", zones.Take(3).Select(z => z.ClashZoneGuid))}...]\n");
                return null; // Skip this cluster - invalid data
            }
            
            double depth = first.StructuralElementThickness > 0.001 ? first.StructuralElementThickness : bboxResult.depth;

            var placementService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementCalculationService(
                (id, _) => zones.FirstOrDefault(z => z.SleeveInstanceId == id) ?? zones.FirstOrDefault()
            );

            XYZ bboxMin = new XYZ(bboxResult.rotatedMinX.Value, bboxResult.rotatedMinY.Value, bboxResult.rotatedMinZ.Value);
            XYZ bboxMax = new XYZ(bboxResult.rotatedMaxX.Value, bboxResult.rotatedMaxY.Value, bboxResult.rotatedMaxZ.Value);

            XYZ placementPoint = placementService.CalculatePlacementPoint(clusterItems, bboxResult.width, bboxResult.height, depth, bboxMin, bboxMax, bboxResult.mid);

            // 🔍 DIAGNOSTIC: Log placement point calculation result
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 PLACEMENT CALC: PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6}), " +
                $"Width={bboxResult.width:F3}, Height={bboxResult.height:F3}, Depth={depth:F3}, Rotation={rotationAngle:F3}\n");

            if (!IsUniqueLocation(batchId, placementPoint)) return null;

            return CreateResult(zones, batchId, comboId, filterId, placementPoint, bboxResult.width, bboxResult.height, depth, rotationAngle);
        }

        /// <summary>
        /// ✅ SPECIALIZED WALL/FRAMING CALCULATION: Strictly separated from floor logic.
        /// </summary>
        private BatchClusterCalculationResult CalculateWallFramingCluster(List<dynamic> clusterItems, string batchId, int comboId, int filterId)
        {
            var zones = clusterItems.Select(x => (x is ClashZone z) ? z : ((ClashZoneWorkItem)x).ClashZone).ToList();
            if (zones.Count == 0) return null;

            var first = zones[0];
            double rotationAngle = _rotationService.DetermineRotationAngle(clusterItems);

            var bboxResult = _rotationService.CalculateRotatedBoundingBox(clusterItems, null, rotationAngle);
            
            // 🔍 DIAGNOSTIC: Log what CalculateRotatedBoundingBox returned (WALL/FRAMING)
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 BBOX CALC RESULT (WALL): Zones={zones.Count}, " +
                $"MinX={bboxResult.rotatedMinX}, MinY={bboxResult.rotatedMinY}, MinZ={bboxResult.rotatedMinZ}, " +
                $"MaxX={bboxResult.rotatedMaxX}, MaxY={bboxResult.rotatedMaxY}, MaxZ={bboxResult.rotatedMaxZ}, " +
                $"Width={bboxResult.width}, Height={bboxResult.height}, Depth={bboxResult.depth}\n");
            
            // ✅ CRITICAL VALIDATION: Reject clusters with invalid bounding boxes
            if (bboxResult.rotatedMinX == null || bboxResult.rotatedMinY == null || bboxResult.rotatedMinZ == null ||
                bboxResult.rotatedMaxX == null || bboxResult.rotatedMaxY == null || bboxResult.rotatedMaxZ == null)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ BBOX VALIDATION FAILED (WALL): Cluster has null bbox values. " +
                    $"Zones={zones.Count}, Category={first.MepElementCategory}, HostType={first.StructuralElementType}, " +
                    $"ZoneGuids=[{string.Join(",", zones.Take(3).Select(z => z.ClashZoneGuid))}...]\n");
                return null; // Skip this cluster - invalid data
            }
            
            // Depth = structural thickness for all host types (Floor, Wall, Framing)
            double depth = first.StructuralElementThickness > 0.001 ? first.StructuralElementThickness : bboxResult.depth;

            var placementService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementCalculationService(
                (id, _) => zones.FirstOrDefault(z => z.SleeveInstanceId == id) ?? zones.FirstOrDefault()
            );

            XYZ bboxMin = new XYZ(bboxResult.rotatedMinX.Value, bboxResult.rotatedMinY.Value, bboxResult.rotatedMinZ.Value);
            XYZ bboxMax = new XYZ(bboxResult.rotatedMaxX.Value, bboxResult.rotatedMaxY.Value, bboxResult.rotatedMaxZ.Value);

            XYZ placementPoint = placementService.CalculatePlacementPoint(clusterItems, bboxResult.width, bboxResult.height, depth, bboxMin, bboxMax, bboxResult.mid);

            if (!IsUniqueLocation(batchId, placementPoint)) return null;

            return CreateResult(zones, batchId, comboId, filterId, placementPoint, bboxResult.width, bboxResult.height, depth, rotationAngle);
        }

        private bool IsUniqueLocation(string batchId, XYZ pt)
        {
            string locKey = $"{batchId}_{pt.X:F1}_{pt.Y:F1}_{pt.Z:F1}";
            return _placedLocations.TryAdd(locKey, 0);
        }

        /// <summary>Returns "Floor", "Wall", or "Structural Framing" so batch grouping keeps floor strictly separate from wall/framing.</summary>
        private static string GetBatchHostType(ClashZone cz)
        {
            if (cz?.StructuralElementType == null) return "Other";
            var t = cz.StructuralElementType.Trim();
            if (t.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0) return "Floor";
            if (t.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0) return "Wall";
            if (t.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)) return "Structural Framing";
            return "Other";
        }

        private BatchClusterCalculationResult CreateResult(List<ClashZone> zones, string batchId, int comboId, int filterId, XYZ pt, double w, double h, double d, double rot)
        {
            var first = zones[0];
            
            // 🔍 DIAGNOSTIC: Log what CreateResult receives
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔍 CREATE RESULT: batchId={batchId}, " +
                $"Placement=({pt.X:F3},{pt.Y:F3},{pt.Z:F3}), Size=({w:F3}x{h:F3}x{d:F3}), " +
                $"Zones={zones.Count}, Category={first.MepElementCategory}\n");
            
            return new BatchClusterCalculationResult
            {
                ClusterGUID = Guid.NewGuid().ToString(),
                ClusterBatchId = batchId,
                PlacementX = pt.X, PlacementY = pt.Y, PlacementZ = pt.Z,
                ClusterWidth = w, ClusterHeight = h, ClusterDepth = d,
                RotationAngleRad = rot,
                HostElementId = first.StructuralElementIdValue,
                HostType = first.StructuralElementType,
                HostOrientation = first.HostOrientation,
                Category = first.MepElementCategory,
                FamilyName = JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementService.GetFamilyName(
                    first.StructuralElementType ?? "Unknown", 
                    first.MepElementCategory ?? "Unknown", 
                    Math.Max(w, h), 
                    zones.Count > 1),
                ConstituentZoneGuids = string.Join(",", zones.Select(z => z.ClashZoneGuid)),
                ComboId = comboId, FilterId = filterId, Status = "Pending", ValidationStatus = "Valid"
            };
        }

        /// <summary>
        /// ✅ FLAG MANAGEMENT: Update MarkedForClusterProcess flag in ClashZones table (mimic slow mode)
        /// Only set to true for zones in clusters with >1 sleeve (multi-sleeve clusters)
        /// Zones not in clusters should remain false/null
        /// </summary>
        private void UpdateMarkedForClusterProcessFlags(ConcurrentBag<BatchClusterCalculationResult> validClusters)
        {
            if (validClusters == null || validClusters.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 FLAG UPDATE: No clusters to update flags for\n");
                return;
            }

            try
            {
                // ✅ PHASE 7 FIX: Collect ALL GUIDs from ALL multi-sleeve clusters first
                var allZoneGuids = new HashSet<string>();
                int multiSleeveClusterCount = 0;

                foreach (var cluster in validClusters)
                {
                    if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                    
                    var zoneGuids = cluster.ConstituentZoneGuids
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(g => g.Trim())
                        .Where(g => !string.IsNullOrEmpty(g))
                        .ToList();
                    
                    if (zoneGuids.Count >= 2)
                    {
                        multiSleeveClusterCount++;
                        foreach (var guid in zoneGuids)
                        {
                            allZoneGuids.Add(guid);
                        }
                    }
                }

                if (allZoneGuids.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 FLAG UPDATE: No multi-sleeve clusters found, no flags to update.\n");
                    return;
                }

                using (var conn = new SQLiteConnection(SqliteConnStr.Build(_databasePath)))
                {
                    conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        // ✅ SINGLE BATCH UPDATE: Use IN clause for all GUIDs at once
                        // SQLite has a limit on parameters (usually 999), so we batch if more than that
                        var guidList = allZoneGuids.ToList();
                        int batchSize = 900;
                        int totalUpdated = 0;

                        for (int i = 0; i < guidList.Count; i += batchSize)
                        {
                            var currentBatch = guidList.Skip(i).Take(batchSize).ToList();
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = trans;
                                // ✅ FIX: Use robust, brace-insensitive GUID matching with REPLACE and UPPER
                                var placeholders = string.Join(",", currentBatch.Select((_, idx) => $"@g{idx}"));
                                cmd.CommandText = $@"
                                    UPDATE ClashZones
                                    SET MarkedForClusterProcess = 1,
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE REPLACE(REPLACE(UPPER(ClashZoneGuid), '{{', ''), '}}', '') IN ({placeholders})";

                                for (int idx = 0; idx < currentBatch.Count; idx++)
                                {
                                    // Clean input GUID for matching
                                    var cleanGuid = currentBatch[idx].Replace("{", "").Replace("}", "").ToUpperInvariant();
                                    cmd.Parameters.AddWithValue($"@g{idx}", cleanGuid);
                                }

                                var affected = cmd.ExecuteNonQuery();
                                totalUpdated += affected;
                                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] [BatchClusterCalculationService] LOUD-DEBUG: Batch of {currentBatch.Count} GUIDs -> {affected} database rows marked for cluster.\n");
                            }
                        }
                        
                        trans.Commit();
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ FLAG UPDATE: Updated MarkedForClusterProcess=TRUE for {totalUpdated} zones across {multiSleeveClusterCount} clusters (Single Transaction)\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ FLAG UPDATE ERROR: Failed to update MarkedForClusterProcess flags: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

    }

    public class ClashZoneWorkItem
    {
        public int SleeveInstanceId { get; set; }
        public ClashZone ClashZone { get; set; }
        public SleeveGroupKey GroupKey { get; set; }
        // For compatibility with legacy dynamics
        public dynamic GetClashZone() => ClashZone;
    }

    public class BatchClusterCalculationResult
    {
        public string ClusterGUID { get; set; } = string.Empty;
        public string ClusterBatchId { get; set; } = string.Empty;
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
        public string? Category { get; set; }
        public string? FamilyName { get; set; }
        public string? ConstituentZoneGuids { get; set; } // JSON or csv
        public int ComboId { get; set; }
        public int FilterId { get; set; }
        public string? Status { get; set; }
        public string? ValidationStatus { get; set; }
    }
}
