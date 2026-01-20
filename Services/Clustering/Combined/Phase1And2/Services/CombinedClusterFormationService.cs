using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services
{
    /// <summary>
    /// Phase 2 grouping service: merges proximate cluster sleeves into combined cluster candidates.
    /// </summary>
    public class CombinedClusterFormationService : ICombinedClusterFormationService
    {
        public IReadOnlyList<CombinedClusterCandidate> BuildCombinedClusters(IReadOnlyList<ClusterSleeveInfo> clusterSleeves)
        {
            if (!OptimizationFlags.UseCombinedClustering || !OptimizationFlags.UseCombinedClusteringPhase1And2)
            {
                return Array.Empty<CombinedClusterCandidate>();
            }

            if (clusterSleeves == null || clusterSleeves.Count == 0)
            {
                return Array.Empty<CombinedClusterCandidate>();
            }

            var eligible = clusterSleeves.Where(s => s.HasValidClusterBounds && s.ClusterSleeveInstanceId > 0).ToList();
            if (eligible.Count == 0)
            {
                return Array.Empty<CombinedClusterCandidate>();
            }

            var tolerance = Math.Max(0, OptimizationFlags.CombinedClusterProximityTolerance);
            var adjacency = BuildAdjacency(eligible, tolerance);

            var visited = new bool[eligible.Count];
            var results = new List<CombinedClusterCandidate>();
            var nextId = 1;

            for (int i = 0; i < eligible.Count; i++)
            {
                if (visited[i])
                {
                    continue;
                }

                var groupIndexes = new List<int>();
                var stack = new Stack<int>();
                stack.Push(i);

                while (stack.Count > 0)
                {
                    var index = stack.Pop();
                    if (visited[index])
                    {
                        continue;
                    }

                    visited[index] = true;
                    groupIndexes.Add(index);

                    if (!adjacency.TryGetValue(index, out var neighbors))
                    {
                        continue;
                    }

                    foreach (var neighbor in neighbors)
                    {
                        if (!visited[neighbor])
                        {
                            stack.Push(neighbor);
                        }
                    }
                }

                var members = groupIndexes.Select(idx => eligible[idx]).ToList();
                var candidate = new CombinedClusterCandidate(nextId++, members);
                results.Add(candidate);
            }

            return results
                .OrderBy(r => r.CombinedClusterInstanceId)
                .ToList();
        }

        private static Dictionary<int, List<int>> BuildAdjacency(IReadOnlyList<ClusterSleeveInfo> sleeves, double tolerance)
        {
            var adjacency = new Dictionary<int, List<int>>();
            for (int i = 0; i < sleeves.Count; i++)
            {
                for (int j = i + 1; j < sleeves.Count; j++)
                {
                    if (IsWithinTolerance(sleeves[i], sleeves[j], tolerance))
                    {
                        AddEdge(adjacency, i, j);
                        AddEdge(adjacency, j, i);
                    }
                }
            }

            return adjacency;
        }

        private static void AddEdge(Dictionary<int, List<int>> adjacency, int from, int to)
        {
            if (!adjacency.TryGetValue(from, out var neighbors))
            {
                neighbors = new List<int>();
                adjacency[from] = neighbors;
            }

            neighbors.Add(to);
        }

        private static bool IsWithinTolerance(ClusterSleeveInfo a, ClusterSleeveInfo b, double tolerance)
        {
            var overlapsX = a.ClusterSleeveBoundingBoxMinX - tolerance <= b.ClusterSleeveBoundingBoxMaxX &&
                            a.ClusterSleeveBoundingBoxMaxX + tolerance >= b.ClusterSleeveBoundingBoxMinX;
            var overlapsY = a.ClusterSleeveBoundingBoxMinY - tolerance <= b.ClusterSleeveBoundingBoxMaxY &&
                            a.ClusterSleeveBoundingBoxMaxY + tolerance >= b.ClusterSleeveBoundingBoxMinY;
            var overlapsZ = a.ClusterSleeveBoundingBoxMinZ - tolerance <= b.ClusterSleeveBoundingBoxMaxZ &&
                            a.ClusterSleeveBoundingBoxMaxZ + tolerance >= b.ClusterSleeveBoundingBoxMinZ;

            return overlapsX && overlapsY && overlapsZ;
        }
    }
}
