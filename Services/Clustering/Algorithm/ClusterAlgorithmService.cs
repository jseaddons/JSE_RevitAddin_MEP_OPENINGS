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
            IEnumerable<IGrouping<SleeveGroupKey, dynamic>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel)
        {
            var result = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
            var lockObj = new object();

            Action<IGrouping<SleeveGroupKey, dynamic>> processGroup = group =>
            {
                var clusters = new List<List<dynamic>>();
                var sleeves = group.ToList();
                var visited = new HashSet<dynamic>(); // Objects are assumed distinct references in the dynamic list

                for (int i = 0; i < sleeves.Count; i++)
                {
                    var seed = sleeves[i];
                    if (visited.Contains(seed)) continue;

                    var cluster = new List<dynamic> { seed };
                    visited.Add(seed);
                    var queue = new Queue<dynamic>();
                    queue.Enqueue(seed);

                    while (queue.Count > 0)
                    {
                        var current = queue.Dequeue();
                        // O(N^2) within group is acceptable as groups are usually small
                        foreach (var other in sleeves)
                        {
                            if (visited.Contains(other)) continue;
                            if (ShouldClusterSleeves(current, other, toleranceDist))
                            {
                                visited.Add(other);
                                cluster.Add(other);
                                queue.Enqueue(other);
                            }
                        }
                    }
                    clusters.Add(cluster);
                }

                if (enableParallel)
                {
                    lock (lockObj)
                    {
                        result[group.Key] = clusters;
                    }
                }
                else
                {
                    result[group.Key] = clusters;
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

        private bool ShouldClusterSleeves(dynamic s1, dynamic s2, double toleranceDist)
        {
            // ⚠️ DIAGNOSTIC 1: Prove method is called
            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🚨 ShouldClusterSleeves CALLED (tolerance={toleranceDist * 304.8:F0}mm)\n");
            
            // ✅ CRITICAL FIX: Check HostElementId FIRST, BEFORE proximity
            ClashZone cz1 = null;
            ClashZone cz2 = null;
            try
            {
                cz1 = s1?.ClashZone as ClashZone;
                cz2 = s2?.ClashZone as ClashZone;
                
                // ⚠️ DIAGNOSTIC 2: Check if ClashZones exist
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}]   ClashZones: cz1={(cz1 != null ? $"EXISTS (Id={cz1.Id.ToString().Substring(0, 8)})" : "NULL")}, cz2={(cz2 != null ? $"EXISTS (Id={cz2.Id.ToString().Substring(0, 8)})" : "NULL")}\n");
                
                if (cz1 == null || cz2 == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}]   ⚠️ WARNING: ClashZone is NULL! Cannot check HostElementId. Proceeding with proximity check only.\n");
                    // Continue to proximity check
                }
                else
                {
                    int host1 = cz1.StructuralElementIdValue;
                    int host2 = cz2.StructuralElementIdValue;
                    
                    // ⚠️ DIAGNOSTIC 3: Show HostElementId values
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}]   HostElementIds: host1={host1}, host2={host2}\n");
                    
                    // If both have valid IDs and they're different → CANNOT cluster
                    if (host1 > 0 && host2 > 0 && host1 != host2)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ❌ REJECT: Different walls (host1={host1} != host2={host2})\n");
                        return false; // Different walls/floors - stop immediately
                    }
                    else if (host1 > 0 && host2 > 0 && host1 == host2)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ✅ ACCEPT: Same wall (host={host1})\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ⚠️ WARNING: HostElementId is 0 or invalid (host1={host1}, host2={host2}). Cannot verify walls.\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}]   ❌ EXCEPTION in HostElementId check: {ex.Message}\n");
            }
            
            // ✅ Continue with existing proximity logic
            double angle1 = cz1?.MepElementRotationAngle ?? 0.0;
            bool isRotated = Math.Abs(angle1) > 1e-6 && !IsAxisAlignedAngle(angle1);
            var checker = ProximityCheckerFactory.CreateChecker(s1, s2, angle1, isRotated);
            
            bool proximityResult = checker.CheckProximity(s1, s2, toleranceDist);
            
            // ⚠️ DIAGNOSTIC 4: Show proximity result
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss}]   Proximity check result: {(proximityResult ? "PASS (will cluster)" : "FAIL (too far)")}\n");
            
            return proximityResult;
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
