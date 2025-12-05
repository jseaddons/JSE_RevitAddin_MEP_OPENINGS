using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm
{
    /// <summary>
    /// Phase 8: Extracted clustering algorithms (XML-based proximity + spatial grid flood-fill).
    /// Keeps logic identical to legacy UniversalClusterService methods; minimal surface changes.
    /// </summary>
    public class ClusterAlgorithmService : IClusterAlgorithmService
    {
        private readonly object _logLock = new object();
        private int _processedGroupsCount = 0;

        public Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, dynamic>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel)
        {
            var clustersByGroup = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
            var groupsList = sleeveGroups.ToList();
            var clustersConcurrent = new ConcurrentDictionary<SleeveGroupKey, List<List<dynamic>>>();
            _processedGroupsCount = 0;

            Action<IGrouping<SleeveGroupKey, dynamic>> processGroup = group =>
            {
                var xmlSleeves = group.ToList();
                var groupClusters = new List<List<dynamic>>();

                Interlocked.Increment(ref _processedGroupsCount);

                var clusters = CalculateClustersUsingXmlData(xmlSleeves, toleranceDist, group.Key.orientation, doc);
                if (clusters.Count > 0)
                {
                    groupClusters.AddRange(clusters);
                }
                clustersConcurrent[group.Key] = groupClusters;
            };

            try
            {
                // ✅ PERFORMANCE: Measure multi-threading benefit
                var singleThreadedTime = 0L;
                var multiThreadedTime = 0L;
                
                if (enableParallel)
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    System.Threading.Tasks.Parallel.ForEach(groupsList, new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, processGroup);
                    sw.Stop();
                    multiThreadedTime = sw.ElapsedMilliseconds;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_performance.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚡ MULTI-THREADING: Processed {groupsList.Count} groups in {multiThreadedTime}ms using {Environment.ProcessorCount} cores\n");
                    }
                }
                else
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    foreach (var g in groupsList) processGroup(g);
                    sw.Stop();
                    singleThreadedTime = sw.ElapsedMilliseconds;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_performance.log",
                            $"[{DateTime.Now:HH:mm:ss}] 🐌 SINGLE-THREADED: Processed {groupsList.Count} groups in {singleThreadedTime}ms\n");
                    }
                }

                clustersByGroup = clustersConcurrent.ToDictionary(k => k.Key, v => v.Value);
            }
            catch (AggregateException)
            {
                // Fallback to single-threaded if parallel processing fails
                clustersByGroup = new Dictionary<SleeveGroupKey, List<List<dynamic>>>();
                foreach (var g in groupsList) processGroup(g); // fallback single-threaded
                clustersByGroup = clustersConcurrent.ToDictionary(k => k.Key, v => v.Value);
            }

            return clustersByGroup;
        }

        public Dictionary<(int x, int y, int z), List<FamilyInstance>> BuildSpatialGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize)
        {
            var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();
            foreach (var s in familySleeves)
            {
                BoundingBoxXYZ? bb = null;
                try { bb = bboxes.ContainsKey(s) ? bboxes[s] : s.get_BoundingBox(null); } catch { }
                if (bb != null)
                {
                    int min_ix = (int)Math.Floor(bb.Min.X / cellSize);
                    int max_ix = (int)Math.Floor(bb.Max.X / cellSize);
                    int min_iy = (int)Math.Floor(bb.Min.Y / cellSize);
                    int max_iy = (int)Math.Floor(bb.Max.Y / cellSize);
                    int min_iz = (int)Math.Floor(bb.Min.Z / cellSize);
                    int max_iz = (int)Math.Floor(bb.Max.Z / cellSize);
                    for (int gx = min_ix; gx <= max_ix; gx++)
                        for (int gy = min_iy; gy <= max_iy; gy++)
                            for (int gz = min_iz; gz <= max_iz; gz++)
                            {
                                var key = (gx, gy, gz);
                                if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                                list.Add(s);
                            }
                }
                else
                {
                    var c = centers[s];
                    int ix = (int)Math.Floor(c.X / cellSize);
                    int iy = (int)Math.Floor(c.Y / cellSize);
                    int iz = (int)Math.Floor(c.Z / cellSize);
                    var key = (ix, iy, iz);
                    if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                    list.Add(s);
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
            var groupClusters = new List<List<FamilyInstance>>();
            var unprocessedSet = new HashSet<FamilyInstance>(familySleeves);
            while (unprocessedSet.Count > 0)
            {
                var start = unprocessedSet.First();
                var queue = new Queue<FamilyInstance>();
                var cluster = new List<FamilyInstance>();
                queue.Enqueue(start);
                unprocessedSet.Remove(start);
                while (queue.Count > 0)
                {
                    var inst = queue.Dequeue();
                    cluster.Add(inst);
                    BoundingBoxXYZ o1_bbox = bboxes.ContainsKey(inst) ? bboxes[inst] : inst.get_BoundingBox(null);
                    var candidates = GetCandidatesFromGrid(inst, o1_bbox, centers, grid, cellSize, toleranceDist);
                    var neighbors = FilterNeighborsByBoundingBox(inst, candidates, o1_bbox, bboxes, unprocessedSet, toleranceDist, groupKey);
                    foreach (var n in neighbors)
                    {
                        if (unprocessedSet.Remove(n)) queue.Enqueue(n);
                    }
                }
                groupClusters.Add(cluster);
            }
            return groupClusters;
        }

        // --- Private legacy logic copied from UniversalClusterService ---
        private List<List<dynamic>> CalculateClustersUsingXmlData(List<dynamic> xmlSleeves, double toleranceDist, string orientation, Document doc)
        {
            var clusters = new List<List<dynamic>>();
            var processed = new HashSet<int>();
            foreach (var sleeve in xmlSleeves)
            {
                if (processed.Contains(sleeve.SleeveInstanceId)) continue;
                var cluster = new List<dynamic> { sleeve };
                processed.Add(sleeve.SleeveInstanceId);
                bool foundNewNeighbors = true;
                while (foundNewNeighbors)
                {
                    foundNewNeighbors = false;
                    foreach (var clusterSleeve in cluster.ToList())
                    {
                        foreach (var otherSleeve in xmlSleeves)
                        {
                            if (processed.Contains(otherSleeve.SleeveInstanceId)) continue;
                            if (ShouldClusterSleeves(clusterSleeve, otherSleeve, toleranceDist))
                            {
                                cluster.Add(otherSleeve);
                                processed.Add(otherSleeve.SleeveInstanceId);
                                foundNewNeighbors = true;
                            }
                        }
                    }
                }
                if (cluster.Count > 1) clusters.Add(cluster);
            }
            return clusters;
        }

        private bool ShouldClusterSleeves(dynamic s1, dynamic s2, double toleranceDist)
        {
            // Wall host check (same wall id)
            if (s1.HostType == "Wall" && s2.HostType == "Wall")
            {
                var host1Id = s1.ClashZone?.StructuralElementIdValue ?? -1;
                var host2Id = s2.ClashZone?.StructuralElementIdValue ?? -1;
                if (host1Id > 0 && host2Id > 0 && host1Id != host2Id) return false;
            }
            // Round pipe/duct skip rotated path
            bool isRoundPipeOrDuct = false;
            var cz1 = s1.ClashZone as ClashZone; var cz2 = s2.ClashZone as ClashZone;
            if (cz1 != null)
            {
                bool isPipe = cz1.MepElementCategory != null && cz1.MepElementCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isRoundDuct = cz1.MepElementCategory != null && cz1.MepElementCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    (string.Equals(cz1.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) || (cz1.MepElementSizeData != null && (string.Equals(cz1.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) || string.Equals(cz1.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase))));
                if (isPipe || isRoundDuct) isRoundPipeOrDuct = true;
            }
            bool shouldUseRotated = false;
            double angle1 = cz1?.MepElementRotationAngle ?? 0.0;
            double angle2 = cz2?.MepElementRotationAngle ?? 0.0;
            if (!isRoundPipeOrDuct && cz1 != null && cz2 != null)
            {
                bool isRotated1 = Math.Abs(angle1) > 1e-6 && !IsAxisAlignedAngle(angle1);
                bool isRotated2 = Math.Abs(angle2) > 1e-6 && !IsAxisAlignedAngle(angle2);
                if (isRotated1 && isRotated2)
                {
                    double a1 = NormalizeAngleDeg(angle1 * 180.0 / Math.PI);
                    double a2 = NormalizeAngleDeg(angle2 * 180.0 / Math.PI);
                    double diff = Math.Abs(a1 - a2);
                    if (diff > 180.0) diff = 360.0 - diff;
                    bool sameAxis = diff <= 1.0 || Math.Abs(diff - 180.0) <= 1.0;
                    if (sameAxis) shouldUseRotated = true;
                }
            }
            if (shouldUseRotated)
            {
                return CheckRotatedSleeveProximity(s1, s2, angle1, toleranceDist);
            }
            return BoundingBoxesOverlapFromXml(s1, s2, toleranceDist);
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
            foreach (var candidate in candidates)
            {
                if (candidate == inst || !unprocessedSet.Contains(candidate)) continue;
                if (!MatchesGroupCriteria(candidate, groupKey)) continue;
                if (groupKey.hostType == "Wall")
                {
                    var host1 = inst.Host; var host2 = candidate.Host;
                    if (host1 == null || host2 == null || host1.Id != host2.Id) continue;
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
