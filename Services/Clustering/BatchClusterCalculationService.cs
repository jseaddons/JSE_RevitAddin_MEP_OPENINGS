using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
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

        public string CalculateAndSave(
            List<ClashZone> clashZones, 
            string targetCategory, 
            int comboId, 
            int filterId,
            Document doc)
        {
            // ✅ PATH 1 REMOVED: Always recalculate clusters (User Request: "Path 1 is cause of bug")
            // CheckExistingClusterData logic removed to force fresh calculation.
            
            // 1. Generate Batch ID
            string batchId = $"{DateTime.Now:yyyyMMdd_HHmmss}_{targetCategory}_{filterId}";
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🚀 STARTING BATCH V2: {batchId}, Zones={clashZones.Count}\n");

            // 2. Group Zones (Standard Grouping Logic: Host, Orientation, Spatial Bucket)
            // Note: Bucketing is used for "Primary Grouping" to limit N^2 complexity. 
            // Splitting across buckets is handled by Deduplication logic later.
            var sleeveGroups = GroupZones(clashZones);
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 📦 Grouped into {sleeveGroups.Length} processing buckets.\n");

            // 3. Form Clusters (Parallel)
            // ✅ READ FROM USER SETTINGS: "Join openings if distance < (mm)"
            var settings = ApplicationProfileService.Instance.GetCurrentSettings();
            double toleranceMM = settings.JoinOpeningsDistance; // User setting (default: 100mm)
            var toleranceDist = RevitUnitConversionService.Instance.ToInternalMillimeters(toleranceMM);
            
            // ✅ CRITICAL FIX: Wrap ClashZone objects in anonymous objects with ClashZone property
            // FormClusters expects items to have a ClashZone property, not be ClashZone objects directly
            var wrappedZones = clashZones.Select(z => new
            {
                SleeveInstanceId = z.SleeveInstanceId,
                Category = z.MepElementCategory,
                ClashZone = z,
                // ✅ FIX: Provide BoundingBox property for fallback ProximityCheckers (BoundingBoxProximityChecker)
                // This mimics Revit's BoundingBoxXYZ structure for dynamic access
                BoundingBox = new
                {
                    Min = new { X = z.SleeveBoundingBoxMinX, Y = z.SleeveBoundingBoxMinY, Z = z.SleeveBoundingBoxMinZ },
                    Max = new { X = z.SleeveBoundingBoxMaxX, Y = z.SleeveBoundingBoxMaxY, Z = z.SleeveBoundingBoxMaxZ }
                }
            });
            
            // Group the wrapped zones
            var groupedZones = wrappedZones.GroupBy(w => new SleeveGroupKey(
                w.ClashZone.StructuralElementIdValue.ToString(), // ✅ FIX: Group by Host ID (stringified) to prevent over-clustering
                w.ClashZone.MepElementCategory ?? "Unknown", 
                w.ClashZone.HostOrientation ?? "Unknown",
                0, 0, 0)); // ✅ FIX: Disable spatial bucketing as requested
            
            var clustersByGroup = _algorithmService.FormClusters(groupedZones, toleranceDist, doc, enableParallel: true);

            // 4. Process Results (Parallel Calculation + Serial DB Save? Or Parallel Save?)
            // We'll collect all Valid Clusters in a concurrent bag first.
            var validClusters = new ConcurrentBag<BatchClusterCalculationResult>();

            System.Threading.Tasks.Parallel.ForEach(clustersByGroup, (kvp) =>
            {
                foreach (var clusterList in kvp.Value)
                {
                    if (clusterList == null || clusterList.Count < 2) continue;
                    
                    try
                    {
                        // ✅ PHASE 2: Branch logic by Host Type EARLY
                        var first = (clusterList[0] is ClashZone z) ? z : (ClashZone)clusterList[0].ClashZone;
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

            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🧮 Calculated {validClusters.Count} candidate clusters. Saving to ClusterSleeves_v2 table...\n");

            // 5. Save to DB (Sequential for SQLite Safety, though SQLite handles concurrent reasonably well)
            SaveToDatabase(validClusters, batchId);
            
            // ✅ FLAG MANAGEMENT: Update MarkedForClusterProcess flag in ClashZones table (mimic slow mode)
            // Only set to true for zones in clusters with >1 sleeve (multi-sleeve clusters)
            // Zones not in clusters should remain false/null
            UpdateMarkedForClusterProcessFlags(validClusters);

            return batchId;
        }
        
        /// <summary>
        /// ✅ PATH 1 FAST PATH: Check if cluster data already exists in database
        /// Returns existing batchId if clusters exist, null otherwise
        /// </summary>
        private string CheckExistingClusterData(int comboId, int filterId, string category)
        {
            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={_databasePath};Version=3;"))
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
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📊 TABLE: Saving to ClusterSleeves_v2 table in database\n");
            
            if (results.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: No clusters to save to ClusterSleeves_v2, returning early\n");
                return;
            }

            // Re-using direct connection style for speed/custom table
            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={_databasePath};Version=3;"))
                {
                    conn.Open();
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Database connection opened, starting transaction for ClusterSleeves_v2\n");
                    
                    using (var trans = conn.BeginTransaction())
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = trans;
                        cmd.CommandText = @"
                            INSERT INTO ClusterSleeves_v2 (
                                ClusterGUID, ClusterBatchId, PlacementX, PlacementY, PlacementZ, 
                                ClusterWidth, ClusterHeight, ClusterDepth, RotationAngleRad,
                                HostElementId, HostType, HostOrientation, Category, FamilyName,
                                ConstituentZoneGuids, ComboId, FilterId, Status, ValidationStatus
                            ) VALUES (
                                @guid, @batch, @x, @y, @z, 
                                @w, @h, @d, @rot,
                                @host, @htype, @horient, @cat, @fam,
                                @zones, @combo, @filter, @status, @valid
                            )";

                        int savedCount = 0;
                        int skippedCount = 0;
                        double tolerance = 0.1; // 0.1 feet tolerance for duplicate detection
                        
                        foreach (var r in results)
                        {
                            try
                            {
                                // ✅ CRITICAL FIX: Check if cluster with same placement point already exists (when DB is not cleared)
                                // This prevents duplicate clusters from being saved when database is not cleared
                                using (var checkCmd = conn.CreateCommand())
                                {
                                    checkCmd.Transaction = trans;
                                    checkCmd.CommandText = @"
                                        SELECT COUNT(*) FROM ClusterSleeves_v2 
                                        WHERE ABS(PlacementX - @x) < @tol 
                                          AND ABS(PlacementY - @y) < @tol 
                                          AND ABS(PlacementZ - @z) < @tol
                                          AND FamilyName = @fam
                                          AND Status IN ('Pending', 'Placed')";
                                    checkCmd.Parameters.AddWithValue("@x", r.PlacementX);
                                    checkCmd.Parameters.AddWithValue("@y", r.PlacementY);
                                    checkCmd.Parameters.AddWithValue("@z", r.PlacementZ);
                                    checkCmd.Parameters.AddWithValue("@tol", tolerance);
                                    checkCmd.Parameters.AddWithValue("@fam", r.FamilyName ?? "");
                                    
                                    int existingCount = Convert.ToInt32(checkCmd.ExecuteScalar());
                                    if (existingCount > 0)
                                    {
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ SKIPPING DUPLICATE CLUSTER in DB: ClusterGUID={r.ClusterGUID}, Location=({r.PlacementX:F6}, {r.PlacementY:F6}, {r.PlacementZ:F6}), ExistingCount={existingCount}\n");
                                        skippedCount++;
                                        continue; // Skip this cluster - duplicate already exists
                                    }
                                }
                                
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@guid", r.ClusterGUID);
                                cmd.Parameters.AddWithValue("@batch", r.ClusterBatchId);
                                cmd.Parameters.AddWithValue("@x", r.PlacementX);
                                cmd.Parameters.AddWithValue("@y", r.PlacementY);
                                cmd.Parameters.AddWithValue("@z", r.PlacementZ);
                                cmd.Parameters.AddWithValue("@w", r.ClusterWidth);
                                cmd.Parameters.AddWithValue("@h", r.ClusterHeight);
                                cmd.Parameters.AddWithValue("@d", r.ClusterDepth);
                                cmd.Parameters.AddWithValue("@rot", r.RotationAngleRad);
                                cmd.Parameters.AddWithValue("@host", r.HostElementId);
                                cmd.Parameters.AddWithValue("@htype", r.HostType ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@horient", r.HostOrientation ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@cat", r.Category ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@fam", r.FamilyName ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@zones", r.ConstituentZoneGuids ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@combo", r.ComboId);
                                cmd.Parameters.AddWithValue("@filter", r.FilterId);
                                cmd.Parameters.AddWithValue("@status", r.Status ?? "Pending");
                                cmd.Parameters.AddWithValue("@valid", r.ValidationStatus ?? "Valid");
                                
                                int rowsAffected = cmd.ExecuteNonQuery();
                                if (rowsAffected > 0)
                                {
                                    savedCount++;
                                }
                                else
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ TABLE: INSERT into ClusterSleeves_v2 returned 0 rows for cluster {r.ClusterGUID}\n");
                                }
                            }
                            catch (Exception ex)
                            {
                                SafeFileLogger.SafeAppendText("batch_v2_errors.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ❌ TABLE ERROR: Failed to INSERT cluster {r.ClusterGUID} into ClusterSleeves_v2: {ex.Message}\n{ex.StackTrace}\n");
                                throw; // Re-throw to rollback transaction
                            }
                        }
                        
                        if (skippedCount > 0)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ DIAGNOSTIC: Skipped {skippedCount} duplicate clusters (already exist in DB)\n");
                        }
                        
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Inserted {savedCount} of {results.Count} clusters into ClusterSleeves_v2, committing transaction\n");
                        
                        trans.Commit();
                        
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ TABLE: Transaction committed successfully, {savedCount} clusters saved to ClusterSleeves_v2 table\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ TABLE CRITICAL ERROR: SaveToDatabase failed to save to ClusterSleeves_v2 table: {ex.Message}\n{ex.StackTrace}\n");
                throw;
            }
        }

        private SleeveGroupKey[] GroupZones(List<ClashZone> zones)
        {
            // Reusing existing grouping logic (Host, Orientation, 10m Buckets)
            return zones
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
                    // SleeveGroupKey doesn't have Items property, return as-is
                    return key;
                })
                .ToArray();
        }

        /// <summary>
        /// ✅ SPECIALIZED FLOOR CALCULATION: Strictly separated from wall logic.
        /// </summary>
        private BatchClusterCalculationResult CalculateFloorCluster(List<dynamic> clusterItems, string batchId, int comboId, int filterId)
        {
            var zones = clusterItems.Select(x => (x is ClashZone z) ? z : (ClashZone)x.ClashZone).ToList();
            if (zones.Count == 0) return null;

            var first = zones[0];
            double rotationAngle = 0; // Floors usually have 0 rotation or are handled by instance placement

            var bboxResult = _rotationService.CalculateRotatedBoundingBox(clusterItems, null, rotationAngle);
            double depth = first.StructuralElementThickness > 0.001 ? first.StructuralElementThickness : bboxResult.depth;

            var placementService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementCalculationService(
                (id, _) => zones.FirstOrDefault(z => z.SleeveInstanceId == id) ?? zones.FirstOrDefault()
            );

            XYZ bboxMin = new XYZ(bboxResult.rotatedMinX ?? 0, bboxResult.rotatedMinY ?? 0, bboxResult.rotatedMinZ ?? 0);
            XYZ bboxMax = new XYZ(bboxResult.rotatedMaxX ?? 0, bboxResult.rotatedMaxY ?? 0, bboxResult.rotatedMaxZ ?? 0);

            XYZ placementPoint = placementService.CalculatePlacementPoint(clusterItems, bboxResult.width, bboxResult.height, depth, bboxMin, bboxMax, bboxResult.mid);

            if (!IsUniqueLocation(batchId, placementPoint)) return null;

            return CreateResult(zones, batchId, comboId, filterId, placementPoint, bboxResult.width, bboxResult.height, depth, rotationAngle);
        }

        /// <summary>
        /// ✅ SPECIALIZED WALL/FRAMING CALCULATION: Strictly separated from floor logic.
        /// </summary>
        private BatchClusterCalculationResult CalculateWallFramingCluster(List<dynamic> clusterItems, string batchId, int comboId, int filterId)
        {
            var zones = clusterItems.Select(x => (x is ClashZone z) ? z : (ClashZone)x.ClashZone).ToList();
            if (zones.Count == 0) return null;

            var first = zones[0];
            double rotationAngle = _rotationService.DetermineRotationAngle(clusterItems);

            var bboxResult = _rotationService.CalculateRotatedBoundingBox(clusterItems, null, rotationAngle);
            
            // Wall/Framing Depth Logic (Authoritative)
            double depth = bboxResult.depth;
            bool isWall = first.StructuralElementType == "Wall" || first.StructuralElementType == "Walls";
            if (isWall && first.WallThickness > 0.001) depth = first.WallThickness;
            else if (!isWall && first.FramingThickness > 0.001) depth = first.FramingThickness;
            else if (first.StructuralElementThickness > 0.001) depth = first.StructuralElementThickness;

            var placementService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementCalculationService(
                (id, _) => zones.FirstOrDefault(z => z.SleeveInstanceId == id) ?? zones.FirstOrDefault()
            );

            XYZ bboxMin = new XYZ(bboxResult.rotatedMinX ?? 0, bboxResult.rotatedMinY ?? 0, bboxResult.rotatedMinZ ?? 0);
            XYZ bboxMax = new XYZ(bboxResult.rotatedMaxX ?? 0, bboxResult.rotatedMaxY ?? 0, bboxResult.rotatedMaxZ ?? 0);

            XYZ placementPoint = placementService.CalculatePlacementPoint(clusterItems, bboxResult.width, bboxResult.height, depth, bboxMin, bboxMax, bboxResult.mid);

            if (!IsUniqueLocation(batchId, placementPoint)) return null;

            return CreateResult(zones, batchId, comboId, filterId, placementPoint, bboxResult.width, bboxResult.height, depth, rotationAngle);
        }

        private bool IsUniqueLocation(string batchId, XYZ pt)
        {
            string locKey = $"{batchId}_{pt.X:F1}_{pt.Y:F1}_{pt.Z:F1}";
            return _placedLocations.TryAdd(locKey, 0);
        }

        private BatchClusterCalculationResult CreateResult(List<ClashZone> zones, string batchId, int comboId, int filterId, XYZ pt, double w, double h, double d, double rot)
        {
            var first = zones[0];
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
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={_databasePath};Version=3;"))
                {
                    conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        int updatedCount = 0;
                        
                        foreach (var cluster in validClusters)
                        {
                            if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) continue;
                            
                            // Parse zone GUIDs from cluster
                            var zoneGuids = cluster.ConstituentZoneGuids
                                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(g => g.Trim())
                                .Where(g => !string.IsNullOrEmpty(g))
                                .ToList();
                            
                            if (zoneGuids.Count == 0) continue;
                            
                            // ✅ CRITICAL: Only update flags for clusters with >1 zone (multi-sleeve clusters)
                            // Single-zone "clusters" should NOT have MarkedForClusterProcess = true
                            // This matches slow mode behavior where only multi-sleeve clusters get the flag
                            if (zoneGuids.Count < 2)
                            {
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 FLAG UPDATE: Skipping single-zone cluster {cluster.ClusterGUID} (zones={zoneGuids.Count})\n");
                                continue;
                            }
                            
                            // Update MarkedForClusterProcess = true for all zones in this multi-sleeve cluster
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = trans;
                                cmd.CommandText = @"
                                    UPDATE ClashZones 
                                    SET MarkedForClusterProcess = 1,
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE ClashZoneGuid IN (" + string.Join(",", zoneGuids.Select((_, i) => $"@guid{i}")) + ")";
                                
                                for (int i = 0; i < zoneGuids.Count; i++)
                                {
                                    cmd.Parameters.AddWithValue($"@guid{i}", zoneGuids[i]);
                                }
                                
                                int rowsAffected = cmd.ExecuteNonQuery();
                                updatedCount += rowsAffected;
                                
                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 FLAG UPDATE: Set MarkedForClusterProcess=TRUE for {rowsAffected} zones in cluster {cluster.ClusterGUID} (zones={zoneGuids.Count})\n");
                            }
                        }
                        
                        trans.Commit();
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ FLAG UPDATE: Updated MarkedForClusterProcess flag for {updatedCount} zones across {validClusters.Count} multi-sleeve clusters\n");
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
