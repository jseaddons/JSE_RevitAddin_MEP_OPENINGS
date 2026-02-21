using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services
{
    /// <summary>
    /// Service responsible for identifying and forming combined clusters.
    /// Phase 3 Implementation.
    /// </summary>
    public class CombinedClusterFormationService : ICombinedClusterFormation
    {
        private readonly ICombinedClusterRepository _repository;

        public CombinedClusterFormationService(ICombinedClusterRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public async Task<List<CombinedClusterCandidate>> FormCombinedClustersAsync(
            List<ClusterSleeveInfo> clustersInGroup, 
            double proximityTolerance = 100, 
            IProgress<string> progress = null)
        {
            if (clustersInGroup == null || clustersInGroup.Count < 2) 
                return new List<CombinedClusterCandidate>();

            // Tolerance is expected in Revit internal units (feet)
            double toleranceFt = proximityTolerance;

            var candidates = new List<CombinedClusterCandidate>();
            int processedCount = 0; // For generating unique local IDs if needed
            
            // 1. Group by Level Optimization
            // We process each level independently. This reduces O(N^2) complexity to Sum(n_i^2),
            // which is huge for multi-level projects, even without multithreading.
            var clustersByLevel = clustersInGroup.GroupBy(c => c.Level).ToList();

            // 2. Process each Level (Sequential)
            // Reverted Parallelism per user request ("if not effective for single level, forget it")
            foreach (var levelGroup in clustersByLevel)
            {
                var processedClusterIds = new HashSet<long>();
                int candidateIdBase = 1; 

                // Sort by size (largest first) - Local to this level
                var sortedClusters = levelGroup.OrderByDescending(c => c.Width * c.Height).ToList();

                foreach (var seed in sortedClusters)
                {
                    if (processedClusterIds.Contains(seed.ClusterSleeveInstanceId)) continue;
                    processedClusterIds.Add(seed.ClusterSleeveInstanceId);

                    var membersList = new List<ClusterSleeveInfo> { seed };
                    
                    // Create candidate
                    var currentCandidate = new CombinedClusterCandidate(processedCount + candidateIdBase++, membersList);
                    currentCandidate.HostType = seed.HostType ?? "Wall";
                    currentCandidate.Orientation = "X-Wall"; 
                    currentCandidate.Level = seed.Level;

                    // Grow candidate by finding nearby clusters
                    bool added;
                    do
                    {
                        added = false;
                        UpdateCandidateGeometry(currentCandidate);

                        foreach (var candidateMember in sortedClusters)
                        {
                            if (processedClusterIds.Contains(candidateMember.ClusterSleeveInstanceId)) continue;

                            var memberBox = new BoundingBoxXYZ();
                            memberBox.Min = new XYZ(candidateMember.ClusterSleeveBoundingBoxMinX, candidateMember.ClusterSleeveBoundingBoxMinY, candidateMember.ClusterSleeveBoundingBoxMinZ);
                            memberBox.Max = new XYZ(candidateMember.ClusterSleeveBoundingBoxMaxX, candidateMember.ClusterSleeveBoundingBoxMaxY, candidateMember.ClusterSleeveBoundingBoxMaxZ);

                            if (IsClusterNearby(currentCandidate.CombinedBoundingBox, memberBox, toleranceFt))
                            {
                                currentCandidate.MemberClusters.Add(candidateMember);
                                if (!currentCandidate.CategoriesInvolved.Contains(candidateMember.Category))
                                    currentCandidate.CategoriesInvolved.Add(candidateMember.Category);
                                
                                processedClusterIds.Add(candidateMember.ClusterSleeveInstanceId);
                                added = true;
                            }
                        }
                    } while (added);

                    // Add valid candidates
                    if (currentCandidate.MemberClusters.Count > 1)
                    {
                        UpdateCandidateGeometry(currentCandidate);
                        candidates.Add(currentCandidate);
                    }
                }
            }
            
            return candidates;
        }

        public List<ClashZone> FindIndividualSleevesNearCombinedCluster(
            CombinedClusterCandidate combinedCluster, 
            double incorporationTolerance = 100)
        {
            if (combinedCluster == null) return new List<ClashZone>();

            double toleranceFt = incorporationTolerance / 304.8;

            return _repository.FindIndividualSleevesNearBoundingBox(
                combinedCluster.CombinedBoundingBox,
                combinedCluster.HostType,
                combinedCluster.Level,
                toleranceFt
            );
        }

        private void UpdateCandidateGeometry(CombinedClusterCandidate candidate)
        {
            if (candidate.MemberClusters.Count == 0) return;

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            foreach (var member in candidate.MemberClusters)
            {
                // member.BoundingBox does not exist. Use scalars.
                minX = Math.Min(minX, member.ClusterSleeveBoundingBoxMinX);
                minY = Math.Min(minY, member.ClusterSleeveBoundingBoxMinY);
                minZ = Math.Min(minZ, member.ClusterSleeveBoundingBoxMinZ);

                maxX = Math.Max(maxX, member.ClusterSleeveBoundingBoxMaxX);
                maxY = Math.Max(maxY, member.ClusterSleeveBoundingBoxMaxY);
                maxZ = Math.Max(maxZ, member.ClusterSleeveBoundingBoxMaxZ);
            }

            // Also include already incorporated individual sleeves
            foreach (var indSleeve in candidate.IncorporatedIndividualSleeves)
            {
                // Treat point as small box
                minX = Math.Min(minX, indSleeve.SleevePlacementPointX);
                minY = Math.Min(minY, indSleeve.SleevePlacementPointY);
                minZ = Math.Min(minZ, indSleeve.SleevePlacementPointZ);
                maxX = Math.Max(maxX, indSleeve.SleevePlacementPointX);
                maxY = Math.Max(maxY, indSleeve.SleevePlacementPointY);
                maxZ = Math.Max(maxZ, indSleeve.SleevePlacementPointZ);
            }

            // Set internal properties directly. 'CombinedBoundingBox' is read-only derived.
            candidate.CombinedBoundingBoxMinX = minX;
            candidate.CombinedBoundingBoxMinY = minY;
            candidate.CombinedBoundingBoxMinZ = minZ;
            candidate.CombinedBoundingBoxMaxX = maxX;
            candidate.CombinedBoundingBoxMaxY = maxY;
            candidate.CombinedBoundingBoxMaxZ = maxZ;
            
            // Width/Height are also derived read-only. We don't need to set them. they are calculated from Max-Min.
        }

        private bool IsClusterNearby(BoundingBoxXYZ box1, BoundingBoxXYZ box2, double tolerance)
        {
            if (box1 == null || box2 == null) return false;
            
            // Should reuse ClashZone.IsBoundingBoxOverlapping logic if possible, 
            // but here we deal with BoundingBoxXYZ objects directly.
            // Check overlap with expansion
            
            bool xOverlap = (box1.Max.X + tolerance) >= box2.Min.X && (box1.Min.X - tolerance) <= box2.Max.X;
            bool yOverlap = (box1.Max.Y + tolerance) >= box2.Min.Y && (box1.Min.Y - tolerance) <= box2.Max.Y;
            bool zOverlap = (box1.Max.Z + tolerance) >= box2.Min.Z && (box1.Min.Z - tolerance) <= box2.Max.Z;

            return xOverlap && yOverlap && zOverlap;
        }
    }
}
