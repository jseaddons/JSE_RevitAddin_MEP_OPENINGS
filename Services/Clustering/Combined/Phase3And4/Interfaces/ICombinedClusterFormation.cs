using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Form combined clusters from candidates.
    /// CPU-bound (no Revit API), can be parallelized.
    /// Phase 3 Component.
    /// </summary>
    public interface ICombinedClusterFormation
    {
        /// <summary>
        /// Identify which clusters should be combined based on proximity.
        /// Uses spatial algorithm (same as regular clustering but for cluster sleeves).
        /// </summary>
        Task<List<CombinedClusterCandidate>> FormCombinedClustersAsync(
            List<ClusterSleeveInfo> clustersInGroup,
            double proximityTolerance = 100,  // mm
            IProgress<string> progress = null);
        
        /// <summary>
        /// Find individual sleeves that should be incorporated into combined cluster.
        /// Returns zones where: ClusterSleeveInstanceId == -1 AND within tolerance of combined bbox.
        /// </summary>
        List<ClashZone> FindIndividualSleevesNearCombinedCluster(
            CombinedClusterCandidate combinedCluster,
            double incorporationTolerance = 100);  // mm
    }
}
