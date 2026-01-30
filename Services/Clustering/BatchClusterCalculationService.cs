using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
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
                ClashZone = z
            });
            
            // Group the wrapped zones
            var groupedZones = wrappedZones.GroupBy(w => new SleeveGroupKey(
                w.ClashZone.StructuralElementType ?? "Unknown",
                w.ClashZone.MepElementCategory ?? "Unknown", 
                w.ClashZone.HostOrientation ?? "Unknown",
                (int)(w.ClashZone.IntersectionPointX / 100000), // ✅ FIX: Disable Bucketing (Large value effectively puts all in same bucket)
                (int)(w.ClashZone.IntersectionPointY / 100000), 
                (int)(w.ClashZone.IntersectionPointZ / 100000)));
            
            var clustersByGroup = _algorithmService.FormClusters(groupedZones, toleranceDist, doc, enableParallel: true);

            // 4. Process Results (Parallel Calculation + Serial DB Save? Or Parallel Save?)
            // We'll collect all Valid Clusters in a concurrent bag first.
            var validClusters = new ConcurrentBag<BatchClusterCalculationResult>();

            System.Threading.Tasks.Parallel.ForEach(clustersByGroup, (kvp) =>
            {
                foreach (var clusterList in kvp.Value)
                {
                    // ✅ FIX: Filter out single-item "clusters"
                    // If Count < 2, no clustering happened (proximity failed).
                    // These should remain as individual sleeves.
                    if (clusterList == null || clusterList.Count < 2) continue;
                    
                    try
                    {
                        var result = CalculateCluster(clusterList, batchId, comboId, filterId);
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

            return batchId;
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
                        foreach (var r in results)
                        {
                            try
                            {
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
                    int bucketX = (int)Math.Floor(z.IntersectionPointX / 32.8); // 10m
                    int bucketY = (int)Math.Floor(z.IntersectionPointY / 32.8);
                    int bucketZ = (int)Math.Floor(z.IntersectionPointZ / 32.8);
                    
                    // HostType/Category logic simplified
                    return new SleeveGroupKey(
                        z.StructuralElementDocumentTitle ?? "Unknown", 
                        z.MepElementCategory ?? "Unknown", 
                        z.HostOrientation ?? "Unknown",
                        bucketX, bucketY, bucketZ);
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

        private BatchClusterCalculationResult CalculateCluster(List<dynamic> clusterItems, string batchId, int comboId, int filterId)
        {
            // 1. Extract ClashZones
            // Ensure correct cast from dynamic
            var zones = clusterItems.Select(x => 
            {
                 if (x is ClashZone z) return z;
                 try { return (ClashZone)x.ClashZone; } catch { return (ClashZone)x; }
            }).ToList();

            if (zones.Count == 0) return null;

            // 2. Identify "Reference" Zone (First one) for Host info
            var first = zones[0];
            
            // ✅ DIAGNOSTIC: Log thickness values from first ClashZone to trace data flow
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📊 CLUSTER ZONE DATA (First Zone):\n" +
                $"    - ClashZoneGuid: {first.ClashZoneGuid}\n" +
                $"    - StructuralElementType: '{first.StructuralElementType}'\n" +
                $"    - StructuralElementThickness: {first.StructuralElementThickness * 304.8:F1}mm ({first.StructuralElementThickness:F6}ft)\n" +
                $"    - WallThickness: {first.WallThickness * 304.8:F1}mm ({first.WallThickness:F6}ft)\n" +
                $"    - FramingThickness: {first.FramingThickness * 304.8:F1}mm ({first.FramingThickness:F6}ft)\n" +
                $"    - HostOrientation: '{first.HostOrientation}'\n" +
                $"    - MepElementCategory: '{first.MepElementCategory}'\n" +
                $"    - Constituent Systems: {string.Join(", ", zones.Select(z => z.MepElementSystemAbbreviation).Distinct())}\n" +
                $"    - Zones in cluster: {zones.Count}\n");

            // 3. Calculate Geometry (Bounding Box, Rotation)
            // A. Calculate Rotation Angle
            // Logic: Use the shared Rotation Service to determine the authoritative rotation angle
            // This reuses the logic handling X-Wall, Y-Wall, and MEP Element Rotation from DB properties.
            double rotationAngle = _rotationService.DetermineRotationAngle(clusterItems);
            
            // B. Calculate Rotated Bounding Box
            // ✅ CRITICAL FIX: Use ClusterRotationService instead of RotatedBoundingBoxCalculator
            // ClusterRotationService can calculate dimensions from ClashZone data (no FamilyInstance needed)
            var bboxResult = _rotationService.CalculateRotatedBoundingBox(
                clusterItems, 
                null, // No actual sleeves in batch mode - will use ClashZone data
                rotationAngle
            );

            // Extract results
            double width = bboxResult.width;
            double height = bboxResult.height;
            double depth = bboxResult.depth;
            
            // ✅ CRITICAL FIX: Override depth with host thickness (SAME LOGIC AS INDIVIDUAL SLEEVES)
            // Bounding box depth is not reliable - use actual wall/floor thickness from ClashZone
            // MUST match SleeveParameterService.GetThickness() priority order!
            bool isWallHost = first.StructuralElementType == "Wall" || first.StructuralElementType == "Walls";
            bool isFramingHost = string.Equals(first.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            // Priority 1: Host-type specific thickness (same as individual sleeves)
            if (isWallHost && first.WallThickness > 0.001)
            {
                depth = first.WallThickness;
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📐 DEPTH FROM WallThickness: {depth * 304.8:F1}mm (HostType=Wall)\n");
            }
            else if (isFramingHost && first.FramingThickness > 0.001)
            {
                depth = first.FramingThickness;
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📐 DEPTH FROM FramingThickness: {depth * 304.8:F1}mm (HostType=Structural Framing)\n");
            }
            // Priority 2: General structural thickness (fallback)
            else if (first.StructuralElementThickness > 0.001)
            {
                depth = first.StructuralElementThickness;
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📐 DEPTH FROM StructuralElementThickness: {depth * 304.8:F1}mm (Fallback)\n");
            }
            else
            {
                // ⚠️ No valid thickness found - log warning but keep bounding box depth as last resort
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ NO THICKNESS DATA: WallThickness={first.WallThickness * 304.8:F1}mm, " +
                    $"FramingThickness={first.FramingThickness * 304.8:F1}mm, " +
                    $"StructuralElementThickness={first.StructuralElementThickness * 304.8:F1}mm, " +
                    $"HostType='{first.StructuralElementType}', Using BBox depth={depth * 304.8:F1}mm\n");
            }
            
            // ✅ FINAL DIAGNOSTIC: Log the final depth value being saved to DB
            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 💾 FINAL CLUSTER DEPTH: {depth * 304.8:F1}mm (saved to ClusterSleeves_v2.ClusterDepth)\n");
            
            var center = bboxResult.mid;
            
            // If center is zero/null, fallback to intersection centroid
            if (center == null || center.IsZeroLength())
            {
                // Fallback: Intersection Centroid
                double avgX = zones.Average(z => z.IntersectionPointX);
                double avgY = zones.Average(z => z.IntersectionPointY);
                double avgZ = zones.Average(z => z.IntersectionPointZ);
                center = new XYZ(avgX, avgY, avgZ);
            }

            // 3. Placement Calculation (Reuse ClusterPlacementCalculationService)
            // ✅ FIX: Use shared service for consistent logic (Damper Lateral Shift, etc.)
            // We create a local instance since we have the list of ClashZones here.
            
            var placementService = new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementCalculationService(
                (id, path) => zones.FirstOrDefault(z => z.SleeveInstanceId == id) ?? zones.FirstOrDefault(), // Simple lookup
                null // No existing cluster placement lookup needed for new batch
            );

            // Wrap zones into dynamic list for the service signature
            var dynamicCluster = zones.Cast<dynamic>().ToList();

            // Construct Min/Max from tuple (handling nulls)
            XYZ bboxMin = new XYZ(bboxResult.rotatedMinX ?? 0, bboxResult.rotatedMinY ?? 0, bboxResult.rotatedMinZ ?? 0);
            XYZ bboxMax = new XYZ(bboxResult.rotatedMaxX ?? 0, bboxResult.rotatedMaxY ?? 0, bboxResult.rotatedMaxZ ?? 0);

            // Calculate Placement Point (Handles Dampers, Walls, Centroids)
            XYZ placementPoint = placementService.CalculatePlacementPoint(
                dynamicCluster, 
                width, 
                height, 
                depth,
                bboxMin, 
                bboxMax,
                bboxResult.mid // ✅ Pass Geometric Center from Rotation Service
            );
            
            // Declare and assign properly (removed duplicate declaration above)
            double cX = placementPoint.X;
            double cY = placementPoint.Y;
            double cZ = placementPoint.Z;

            // 4. DEDUPLICATION (Location Based)
            string locKey = $"{batchId}_{cX:F1}_{cY:F1}_{cZ:F1}";
            if (!_placedLocations.TryAdd(locKey, 0))
            {
                if (!batchId.Contains("Test"))
                     SafeFileLogger.SafeAppendText("batch_v2.log", $"⚠️ SKIPPING DUPLICATE at ({cX:F1}, {cY:F1}, {cZ:F1})\n");
                return null;
            }
            
            // ... (GUID generation call above)
            return new BatchClusterCalculationResult
            {
                // ... (object initializer above)
                ClusterGUID = Guid.NewGuid().ToString(),
                ClusterBatchId = batchId,
                PlacementX = cX, PlacementY = cY, PlacementZ = cZ,
                ClusterWidth = width, ClusterHeight = height, ClusterDepth = depth,
                RotationAngleRad = rotationAngle, // ✅ FIX: Assign calculated rotation angle
                HostElementId = first.StructuralElementIdValue,
                // ✅ CRITICAL FIX: Use StructuralElementType (Wall/Floor) not DocumentTitle (empty for local elements)
                HostType = first.StructuralElementType,
                HostOrientation = first.HostOrientation,
                Category = first.MepElementCategory,
                FamilyName = JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement.ClusterPlacementService.GetFamilyName(
                    first.StructuralElementType ?? "Unknown", 
                    first.MepElementCategory ?? "Unknown", 
                    Math.Max(width, height), 
                    zones.Count > 1), // isCluster = true ONLY if multiple items. Single items behave like individual sleeves.
                ConstituentZoneGuids = string.Join(",", zones.Select(z => z.ClashZoneGuid)),
                ComboId = comboId, FilterId = filterId, Status = "Pending", ValidationStatus = "Valid"
            };
        }

    }

    public class BatchClusterCalculationResult
    {
        public string ClusterGUID { get; set; }
        public string ClusterBatchId { get; set; }
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
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string ConstituentZoneGuids { get; set; } // JSON or csv
        public int ComboId { get; set; }
        public int FilterId { get; set; }
        public string Status { get; set; }
        public string ValidationStatus { get; set; }
    }
}
