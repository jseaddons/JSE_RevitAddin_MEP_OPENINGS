using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm
{
    public class ClusterAlgorithmService : IClusterAlgorithmService
    {
        public Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel)
        {
            var result = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
            var lockObj = new object();

            Action<IGrouping<SleeveGroupKey, JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>> processGroup = group =>
            {
                var workItems = group.ToList();
                if (workItems.Count == 0)
                {
                    lock (lockObj) result[group.Key] = new List<List<dynamic>>();
                    return;
                }

                // ✅ PHASE 10: High-Performance Spatial Clustering
                List<List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>> clusters = FormClustersSpatial(workItems, toleranceDist);

                // Convert to dynamic for compatibility with orchestrator interface
                var dynamicClusters = clusters.Select(c => c.Cast<dynamic>().ToList()).ToList();

                lock (lockObj)
                {
                    result[group.Key] = dynamicClusters;
                }
            };

            if (enableParallel)
            {
                System.Threading.Tasks.Parallel.ForEach(sleeveGroups, processGroup);
            }
            else
            {
                foreach (var group in sleeveGroups)
                {
                    processGroup(group);
                }
            }
            return result;
        }

        public Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, dynamic>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel)
        {
            // Convert legacy dynamic grouping to typed ClashZoneWorkItem grouping
            var wrappedGroups = sleeveGroups.Select(g => 
            {
                var workItems = new List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>();
                foreach (var d in g)
                {
                    if (d == null) continue;
                    ClashZone cz = (d is ClashZone) ? (ClashZone)d : (ClashZone)d.ClashZone;
                    if (cz == null) continue;

                    workItems.Add(new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem
                    {
                        ClashZone = cz,
                        SleeveInstanceId = cz.SleeveInstanceId,
                        GroupKey = g.Key
                    });
                }
                // Return as IGrouping by grouping already grouped items by the same key
                return workItems.GroupBy(w => w.GroupKey).FirstOrDefault();
            }).Where(g => g != null).ToList();

            return FormClusters(wrappedGroups, toleranceDist, doc, enableParallel);
        }

        /// <summary>
        /// O(N^2) Standard Algorithm - Efficient for small N (<50).
        /// </summary>
        private List<List<dynamic>> FormClustersDirect(List<dynamic> sleeves, double toleranceDist)
        {
            var clusters = new List<List<dynamic>>();
            var visitedGuids = new HashSet<Guid>();

            for (int i = 0; i < sleeves.Count; i++)
            {
                var seed = sleeves[i];
                ClashZone seedCz = seed?.ClashZone as ClashZone;
                if (seedCz == null) continue;

                Guid seedGuid = seedCz.Id;
                if (visitedGuids.Contains(seedGuid)) continue;

                var cluster = new List<dynamic> { seed };
                visitedGuids.Add(seedGuid);

                var queue = new Queue<dynamic>();
                queue.Enqueue(seed);

                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    // O(N) scan for neighbors
                    foreach (var other in sleeves)
                    {
                        ClashZone otherCz = other?.ClashZone as ClashZone;
                        if (otherCz == null) continue;

                        Guid otherGuid = otherCz.Id;
                        if (visitedGuids.Contains(otherGuid)) continue;

                        if (ShouldClusterSleeves(current, other, toleranceDist))
                        {
                            visitedGuids.Add(otherGuid);
                            cluster.Add(other);
                            queue.Enqueue(other);
                        }
                    }
                }
                clusters.Add(cluster);
            }
            return clusters;
        }

        /// <summary>
        /// O(N) Spatial Hashing Algorithm - Optimized for large N (>50).
        /// Uses a generous cell size to partition sleeves and reduce comparisons.
        /// </summary>
        private List<List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>> FormClustersSpatial(
            List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem> sleeves, 
            double toleranceDist)
        {
            var clusters = new List<List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>>();
            var visitedGuids = new HashSet<Guid>();
            
            const double CELL_SIZE = 5.0; 
            var grid = new Dictionary<(int, int, int), List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>>();
            
            (int x, int y, int z) GetCell(ClashZone cz)
            {
                 double cx = (cz.SleeveBoundingBoxMinX + cz.SleeveBoundingBoxMaxX) / 2.0;
                 double cy = (cz.SleeveBoundingBoxMinY + cz.SleeveBoundingBoxMaxY) / 2.0;
                 double cz_z = (cz.SleeveBoundingBoxMinZ + cz.SleeveBoundingBoxMaxZ) / 2.0;
                 
                 return (
                     (int)Math.Floor(cx / CELL_SIZE),
                     (int)Math.Floor(cy / CELL_SIZE),
                     (int)Math.Floor(cz_z / CELL_SIZE)
                 );
            }

            foreach (var s in sleeves)
            {
                var cz = s.ClashZone;
                var key = GetCell(cz);
                if (!grid.ContainsKey(key)) grid[key] = new List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>();
                grid[key].Add(s);
            }

            // 3. Cluster with Grid Lookup
            for (int i = 0; i < sleeves.Count; i++)
            {
                var seed = sleeves[i];
                var seedCz = seed.ClashZone;
                Guid seedGuid = seedCz.Id;
                if (visitedGuids.Contains(seedGuid)) continue;

                var cluster = new List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem> { seed };
                visitedGuids.Add(seedGuid);

                var queue = new Queue<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>();
                queue.Enqueue(seed);

                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    var currentCz = current.ClashZone;
                    var centerKey = GetCell(currentCz);

                    // Check seed's cell AND all 26 neighbors
                    for (int dx = -1; dx <= 1; dx++)
                    {
                        for (int dy = -1; dy <= 1; dy++)
                        {
                            for (int dz = -1; dz <= 1; dz++)
                            {
                                var neighborKey = (centerKey.Item1 + dx, centerKey.Item2 + dy, centerKey.Item3 + dz);
                                if (!grid.TryGetValue(neighborKey, out var candidates)) continue;

                                foreach (var other in candidates)
                                {
                                    var otherCz = other.ClashZone;
                                    Guid otherGuid = otherCz.Id;
                                    if (visitedGuids.Contains(otherGuid)) continue;
                                    
                                    // Narrow Phase: Precise Check
                                    if (ShouldClusterSleeves(current, other, toleranceDist))
                                    {
                                        visitedGuids.Add(otherGuid);
                                        cluster.Add(other);
                                        queue.Enqueue(other);
                                    }
                                }
                            }
                        }
                    }
                }
                clusters.Add(cluster);
            }
            return clusters;
        }

        public Dictionary<(int x, int y, int z), List<FamilyInstance>> BuildSpatialGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize)
        {
            var grid = new Dictionary<(int x, int y, int z), List<FamilyInstance>>();
            foreach (var inst in familySleeves)
            {
                BoundingBoxXYZ bbox = null;
                if (bboxes.TryGetValue(inst, out var b)) bbox = b;
                else bbox = inst.get_BoundingBox(null);

                if (bbox != null)
                {
                    int minX = (int)Math.Floor(bbox.Min.X / cellSize);
                    int maxX = (int)Math.Floor(bbox.Max.X / cellSize);
                    int minY = (int)Math.Floor(bbox.Min.Y / cellSize);
                    int maxY = (int)Math.Floor(bbox.Max.Y / cellSize);
                    int minZ = (int)Math.Floor(bbox.Min.Z / cellSize);
                    int maxZ = (int)Math.Floor(bbox.Max.Z / cellSize);

                    for (int x = minX; x <= maxX; x++)
                        for (int y = minY; y <= maxY; y++)
                            for (int z = minZ; z <= maxZ; z++)
                            {
                                var key = (x, y, z);
                                if (!grid.ContainsKey(key)) grid[key] = new List<FamilyInstance>();
                                grid[key].Add(inst);
                            }
                }
                else if (centers.ContainsKey(inst))
                {
                    var c = centers[inst];
                    int x = (int)Math.Floor(c.X / cellSize);
                    int y = (int)Math.Floor(c.Y / cellSize);
                    int z = (int)Math.Floor(c.Z / cellSize);
                    var key = (x, y, z);
                    if (!grid.ContainsKey(key)) grid[key] = new List<FamilyInstance>();
                    grid[key].Add(inst);
                }
            }
            return grid;
        }

        public List<List<FamilyInstance>> FormClustersFromGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<(int x, int y, int z), List<FamilyInstance>> grid,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize,
            double toleranceDist,
            SleeveGroupKey groupKey)
        {
            var clusters = new List<List<FamilyInstance>>();
            var visited = new HashSet<FamilyInstance>();

            foreach (var seed in familySleeves)
            {
                if (visited.Contains(seed)) continue;

                var cluster = new List<FamilyInstance> { seed };
                visited.Add(seed);
                var queue = new Queue<FamilyInstance>();
                queue.Enqueue(seed);

                BoundingBoxXYZ seedBbox = null;
                if (bboxes.TryGetValue(seed, out var b)) seedBbox = b;
                else seedBbox = seed.get_BoundingBox(null);

                while (queue.Count > 0)
                {
                    var current = queue.Dequeue();
                    BoundingBoxXYZ currBbox = null;
                    if (bboxes.TryGetValue(current, out var cb)) currBbox = cb;
                    else currBbox = current.get_BoundingBox(null);

                    // Get candidates from grid
                    var candidates = GetCandidatesFromGrid(current, currBbox, centers, grid, cellSize, toleranceDist);

                    // Filter candidates
                    var neighbors = FilterNeighborsByBoundingBox(current, candidates, currBbox, bboxes, visited, toleranceDist, groupKey);

                    foreach (var neighbor in neighbors)
                    {
                        if (visited.Contains(neighbor)) continue;
                        visited.Add(neighbor);
                        cluster.Add(neighbor);
                        queue.Enqueue(neighbor);
                    }
                }
                clusters.Add(cluster);
            }
            return clusters;
        }
        private bool ShouldClusterSleeves(
            JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem s1, 
            JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem s2, 
            double toleranceDist)
        {
            var cz1 = s1.ClashZone;
            var cz2 = s2.ClashZone;
            
            if (cz1.Id == cz2.Id) return false;

            // 1. Check StructuralElementId (Host)
            long host1 = cz1.StructuralElementIdValue;
            long host2 = cz2.StructuralElementIdValue;
            
            if (host1 != host2)
            {
                 if (host1 > 0 && host2 > 0) return false;
            }

            // 2. Check MEP Category
            if (cz1.MepElementCategory != cz2.MepElementCategory) return false; 
            
            // ✅ PERSISTENCE CHECK: If bounding boxes are zero/null, log and skip (user requested no expensive fallback)
            bool s1HasBbox = Math.Abs(cz1.SleeveBoundingBoxMaxX - cz1.SleeveBoundingBoxMinX) > 1e-6;
            bool s2HasBbox = Math.Abs(cz2.SleeveBoundingBoxMaxX - cz2.SleeveBoundingBoxMinX) > 1e-6;

            if (!s1HasBbox || !s2HasBbox)
            {
                // Only log once per session/run to avoid log bloat if needed, but for now simple log
                if (!s1HasBbox) SafeFileLogger.SafeAppendText("clustering_warnings.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Zone {cz1.ClashZoneGuid} has zero BBox - skipping proximity check.\n");
                if (!s2HasBbox) SafeFileLogger.SafeAppendText("clustering_warnings.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Zone {cz2.ClashZoneGuid} has zero BBox - skipping proximity check.\n");
                return false;
            }


            // ✅ PHASE 10: Optimize Proximity Check
            // ✅ CRITICAL FIX: Check if EITHER sleeve is rotated (not just first one!)
            // If either is rotated, we must use rotated proximity checker
            double angle1 = cz1.MepElementRotationAngle;
            double angle2 = cz2.MepElementRotationAngle;
            bool isRotated1 = Math.Abs(angle1) > 1e-6 && !IsAxisAlignedAngle(angle1);
            bool isRotated2 = Math.Abs(angle2) > 1e-6 && !IsAxisAlignedAngle(angle2);
            
            // 🔍 DIAGNOSTIC: Log rotation check for debugging
            if (isRotated1 || isRotated2)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 ROTATION CHECK: Zone1={cz1.ClashZoneGuid.Substring(0,8)}, " +
                    $"Angle1={angle1:F4}, IsRot1={isRotated1}, Zone2={cz2.ClashZoneGuid.Substring(0,8)}, " +
                    $"Angle2={angle2:F4}, IsRot2={isRotated2}\n");
            }

            if (!isRotated1 && !isRotated2)  // Both must be straight for AABB fast path
            {
                // ✅ CRITICAL FIX: Use SAME bounding boxes as slow mode (CornerProximityChecker)
                // Slow mode uses SleeveCorner1-4 via GetBoundingBoxFromCorners(), NOT SleeveBoundingBoxMin/Max
                // SleeveBoundingBoxMin/Max are AABB and oversized for rotated sleeves!
                
                // Helper to compute 1D gap between two intervals (0 if overlapping)
                double Gap(double min1, double max1, double min2, double max2)
                {
                    if (max1 < min2) return min2 - max1;
                    if (max2 < min1) return min1 - max2;
                    return 0.0; // Overlapping
                }

                // Get bounding boxes from corners (SAME as SleeveCornerProximityHelper)
                var bbox1 = GetBoundingBoxFromCorners(cz1);
                var bbox2 = GetBoundingBoxFromCorners(cz2);
                
                // 🔍 VALIDATION: Check if corners are actually populated
                bool hasValidCorners1 = (bbox1.maxX - bbox1.minX) > 0.001 && (bbox1.maxY - bbox1.minY) > 0.001;
                bool hasValidCorners2 = (bbox2.maxX - bbox2.minX) > 0.001 && (bbox2.maxY - bbox2.minY) > 0.001;
                
                if (!hasValidCorners1 || !hasValidCorners2)
                {
                    // FALLBACK: Use SleeveBoundingBoxMin/Max if corners not available
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ CORNER_FALLBACK: Zone1={cz1.ClashZoneGuid?.Substring(0,8)}(corners={hasValidCorners1}), " +
                        $"Zone2={cz2.ClashZoneGuid?.Substring(0,8)}(corners={hasValidCorners2}) - Using AABB fallback\n");
                    
                    if (!hasValidCorners1)
                    {
                        bbox1 = (cz1.SleeveBoundingBoxMinX, cz1.SleeveBoundingBoxMaxX, cz1.SleeveBoundingBoxMinY, 
                                 cz1.SleeveBoundingBoxMaxY, cz1.SleeveBoundingBoxMinZ, cz1.SleeveBoundingBoxMaxZ);
                    }
                    if (!hasValidCorners2)
                    {
                        bbox2 = (cz2.SleeveBoundingBoxMinX, cz2.SleeveBoundingBoxMaxX, cz2.SleeveBoundingBoxMinY, 
                                 cz2.SleeveBoundingBoxMaxY, cz2.SleeveBoundingBoxMinZ, cz2.SleeveBoundingBoxMaxZ);
                    }
                }
                
                // ✅ CRITICAL: Use SAME tolerance adjustment as slow mode (SleeveCornerProximityHelper)
                double effectiveTolerance = Math.Ceiling(toleranceDist * 304.8) / 304.8;

                string hostType = cz1.StructuralElementType ?? "";
                string orientation = cz1.HostOrientation ?? "";
                double distance;

                // Floor: 2D distance in X,Y plane (ignore Z)
                if (hostType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double gapX = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                    double gapY = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                    distance = Math.Sqrt(gapX * gapX + gapY * gapY);
                }
                // Wall/Framing: Check orientation!
                // X-wall: distance in X,Z plane (ignore Y - through wall)
                // Y-wall: distance in Y,Z plane (ignore X - through wall)
                else if (hostType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0 || 
                         hostType.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (orientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        // X-wall: 2D distance in X,Z plane
                        double gapX = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                        double gapZ = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                        distance = Math.Sqrt(gapX * gapX + gapZ * gapZ);
                    }
                    else // Y-wall (or unknown - default to Y,Z)
                    {
                        // Y-wall: 2D distance in Y,Z plane
                        double gapY = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                        double gapZ = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                        distance = Math.Sqrt(gapY * gapY + gapZ * gapZ);
                    }
                }
                // Default: full 3D distance
                else
                {
                    double gapX = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                    double gapY = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                    double gapZ = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                    distance = Math.Sqrt(gapX * gapX + gapY * gapY + gapZ * gapZ);
                }

                bool shouldCluster = distance <= effectiveTolerance;
                
                // 🔍 DIAGNOSTIC: Log clustering decision for over-stretch analysis
                if (shouldCluster) // Log ALL clusters
                {
                    string source1 = hasValidCorners1 ? "corners" : "aabb";
                    string source2 = hasValidCorners2 ? "corners" : "aabb";
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] CLUSTER_DECISION: Zone1={cz1.ClashZoneGuid?.Substring(0,8)}({source1}), Zone2={cz2.ClashZoneGuid?.Substring(0,8)}({source2}), " +
                        $"Host={hostType}, Orient={orientation}, Distance={distance:F6}ft ({distance*304.8:F1}mm), " +
                        $"Tolerance={effectiveTolerance:F6}ft ({effectiveTolerance*304.8:F1}mm) [raw={toleranceDist:F6}ft], Result={shouldCluster}\n" +
                        $"  BBox1: X=[{bbox1.minX:F2},{bbox1.maxX:F2}], Y=[{bbox1.minY:F2},{bbox1.maxY:F2}], Z=[{bbox1.minZ:F2},{bbox1.maxZ:F2}]\n" +
                        $"  BBox2: X=[{bbox2.minX:F2},{bbox2.maxX:F2}], Y=[{bbox2.minY:F2},{bbox2.maxY:F2}], Z=[{bbox2.minZ:F2},{bbox2.maxZ:F2}]\n");
                }
                
                return shouldCluster;
            }

            // Fallback to Rotated Proximity (if either sleeve is rotated)
            // Use the angle from whichever sleeve is rotated (or angle1 if both are)
            double rotationAngle = isRotated1 ? angle1 : angle2;
            var checker = ProximityCheckerFactory.CreateChecker(s1.ClashZone, s2.ClashZone, rotationAngle, true);
            bool result = checker.CheckProximity(s1.ClashZone, s2.ClashZone, toleranceDist);
            
            // 🔍 DIAGNOSTIC: Log proximity check result for rotated elements
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] 🔍 ROTATED PROXIMITY: Zone1={cz1.ClashZoneGuid.Substring(0,8)}, " +
                $"Zone2={cz2.ClashZoneGuid.Substring(0,8)}, Angle={rotationAngle:F4}, Result={result}, " +
                $"Tolerance={toleranceDist:F3}\n");
            
            return result;
        }

        private static double NormalizeAngleDeg(double deg)
        {
            while (deg < 0) deg += 360.0; while (deg >= 360.0) deg -= 360.0; return deg;
        }
        private bool IsAxisAlignedAngle(double angle)
        {
            double deg = NormalizeAngleDeg(angle * 180.0 / Math.PI);
            double[] targets = { 0, 90, 180, 270 };
            return targets.Any(t => Math.Abs(deg - t) < 2.0);
        }

        // Simplified copies of legacy helpers (exact logic preserved)
        private bool CheckRotatedSleeveProximity(dynamic s1, dynamic s2, double axisAngle, double tolerance)
        {
            var cz1 = s1.ClashZone as ClashZone; var cz2 = s2.ClashZone as ClashZone;
            if (cz1?.RotatedBoundingBoxMinX == null || cz2?.RotatedBoundingBoxMinX == null) return false;
            double minX1 = cz1.RotatedBoundingBoxMinX.Value - tolerance;
            double maxX1 = cz1.RotatedBoundingBoxMaxX!.Value + tolerance;
            double minY1 = cz1.RotatedBoundingBoxMinY!.Value - tolerance;
            double maxY1 = cz1.RotatedBoundingBoxMaxY!.Value + tolerance;
            double minX2 = cz2.RotatedBoundingBoxMinX!.Value;
            double maxX2 = cz2.RotatedBoundingBoxMaxX!.Value;
            double minY2 = cz2.RotatedBoundingBoxMinY!.Value;
            double maxY2 = cz2.RotatedBoundingBoxMaxY!.Value;
            bool overlapX = maxX2 >= minX1 && minX2 <= maxX1;
            bool overlapY = maxY2 >= minY1 && minY2 <= maxY1;
            return overlapX && overlapY;
        }

        private bool BoundingBoxesOverlapFromXml(dynamic s1, dynamic s2, double tolerance)
        {
            var cz1 = s1.ClashZone as ClashZone; var cz2 = s2.ClashZone as ClashZone;
            if (cz1 == null || cz2 == null) return false;
            double minX1 = cz1.SleeveBoundingBoxMinX - tolerance;
            double minY1 = cz1.SleeveBoundingBoxMinY - tolerance;
            double minZ1 = cz1.SleeveBoundingBoxMinZ - tolerance;
            double maxX1 = cz1.SleeveBoundingBoxMaxX + tolerance;
            double maxY1 = cz1.SleeveBoundingBoxMaxY + tolerance;
            double maxZ1 = cz1.SleeveBoundingBoxMaxZ + tolerance;
            double minX2 = cz2.SleeveBoundingBoxMinX;
            double minY2 = cz2.SleeveBoundingBoxMinY;
            double minZ2 = cz2.SleeveBoundingBoxMinZ;
            double maxX2 = cz2.SleeveBoundingBoxMaxX;
            double maxY2 = cz2.SleeveBoundingBoxMaxY;
            double maxZ2 = cz2.SleeveBoundingBoxMaxZ;
            bool overlapX = maxX2 >= minX1 && minX2 <= maxX1;
            bool overlapY = maxY2 >= minY1 && minY2 <= maxY1;
            bool overlapZ = maxZ2 >= minZ1 && minZ2 <= maxZ1;
            return overlapX && overlapY && overlapZ;
        }

        /// <summary>
        /// Get bounding box from sleeve corners (SleeveCorner1-4).
        /// This matches the logic in SleeveCornerProximityHelper.GetBoundingBoxFromCorners()
        /// and ensures fast path uses the SAME bounding boxes as slow mode.
        /// </summary>
        private (double minX, double maxX, double minY, double maxY, double minZ, double maxZ) GetBoundingBoxFromCorners(ClashZone cz)
        {
            var minX = Math.Min(Math.Min(cz.SleeveCorner1X ?? 0, cz.SleeveCorner2X ?? 0),
                               Math.Min(cz.SleeveCorner3X ?? 0, cz.SleeveCorner4X ?? 0));
            var maxX = Math.Max(Math.Max(cz.SleeveCorner1X ?? 0, cz.SleeveCorner2X ?? 0),
                               Math.Max(cz.SleeveCorner3X ?? 0, cz.SleeveCorner4X ?? 0));

            var minY = Math.Min(Math.Min(cz.SleeveCorner1Y ?? 0, cz.SleeveCorner2Y ?? 0),
                               Math.Min(cz.SleeveCorner3Y ?? 0, cz.SleeveCorner4Y ?? 0));
            var maxY = Math.Max(Math.Max(cz.SleeveCorner1Y ?? 0, cz.SleeveCorner2Y ?? 0),
                               Math.Max(cz.SleeveCorner3Y ?? 0, cz.SleeveCorner4Y ?? 0));

            var minZ = Math.Min(Math.Min(cz.SleeveCorner1Z ?? 0, cz.SleeveCorner2Z ?? 0),
                               Math.Min(cz.SleeveCorner3Z ?? 0, cz.SleeveCorner4Z ?? 0));
            var maxZ = Math.Max(Math.Max(cz.SleeveCorner1Z ?? 0, cz.SleeveCorner2Z ?? 0),
                               Math.Max(cz.SleeveCorner3Z ?? 0, cz.SleeveCorner4Z ?? 0));

            // ✅ SAFETY CHECK: Ensure minimum thickness (e.g. 1mm ~ 0.0033 ft) to prevent 2D boxes
            const double minThickness = 0.0033; // ~1mm
            if (maxX - minX < minThickness) { minX -= minThickness/2; maxX += minThickness/2; }
            if (maxY - minY < minThickness) { minY -= minThickness/2; maxY += minThickness/2; }
            if (maxZ - minZ < minThickness) { minZ -= minThickness/2; maxZ += minThickness/2; }

            return (minX, maxX, minY, maxY, minZ, maxZ);
        }

        private List<FamilyInstance> GetCandidatesFromGrid(
            FamilyInstance inst,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, XYZ> centers,
            Dictionary<(int x, int y, int z), List<FamilyInstance>> grid,
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

            // ✅ CRITICAL FIX: Get HostElementId from PARAMETERS (not inst.Host which is null in batch mode)
            int host1Id = -1;
            try
            {
                // Individual sleeves store HostElementId in "Host Element ID" parameter
                var host1Param = inst.LookupParameter("Host Element ID");
                if (host1Param != null && host1Param.StorageType == StorageType.Integer)
                {
                    host1Id = host1Param.AsInteger();
                }
            }
            catch { }

            foreach (var candidate in candidates)
            {
                if (candidate == inst || !unprocessedSet.Contains(candidate)) continue;
                if (!MatchesGroupCriteria(candidate, groupKey)) continue;

                // ✅ FIX: Check HostElementId parameter
                if (host1Id > 0)
                {
                    int host2Id = -1;
                    try
                    {
                        var host2Param = candidate.LookupParameter("Host Element ID");
                        if (host2Param != null && host2Param.StorageType == StorageType.Integer)
                        {
                            host2Id = host2Param.AsInteger();
                        }
                    }
                    catch { }

                    // If both have valid IDs and are different -> REJECT
                    if (host2Id > 0 && host1Id != host2Id)
                    {
                        continue;
                    }
                }
                // Fallback to inst.Host check if parameters missing (legacy support)
                else if (groupKey.hostType == "Wall")
                {
                    var host1 = inst.Host; var host2 = candidate.Host;
                    if (host1 != null && host2 != null && host1.Id != host2.Id) continue;
                }

                BoundingBoxXYZ o2_bbox = bboxes.ContainsKey(candidate) ? bboxes[candidate] : candidate.get_BoundingBox(null);
                if (o2_bbox == null) continue;
                if (BoundingBoxesOverlap(o1_bbox, o2_bbox, toleranceDist)) neighbors.Add(candidate);
            }
            return neighbors;
        }

        private bool BoundingBoxesOverlap(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2, double tolerance)
        {
            if (bbox1 == null || bbox2 == null) return false;
            bool overlapX = bbox2.Max.X >= bbox1.Min.X - tolerance && bbox2.Min.X <= bbox1.Max.X + tolerance;
            bool overlapY = bbox2.Max.Y >= bbox1.Min.Y - tolerance && bbox2.Min.Y <= bbox1.Max.Y + tolerance;
            bool overlapZ = bbox2.Max.Z >= bbox1.Min.Z - tolerance && bbox2.Min.Z <= bbox1.Max.Z + tolerance;
            return overlapX && overlapY && overlapZ;
        }

        private bool MatchesGroupCriteria(FamilyInstance inst, SleeveGroupKey key)
        {
            try
            {
                var cz = inst.LookupParameter("Clash Zone Host Type")?.AsString();
                var orientation = inst.LookupParameter("Sleeve Orientation")?.AsString();
                return (cz == key.hostType) && (orientation == key.orientation);
            }
            catch { return false; }
        }
    }
}
