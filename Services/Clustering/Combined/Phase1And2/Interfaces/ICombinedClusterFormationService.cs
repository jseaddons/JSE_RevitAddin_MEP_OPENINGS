using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces
{
    /// <summary>
    /// Groups cluster sleeves into combined cluster candidates.
    /// </summary>
    public interface ICombinedClusterFormationService
    {
        IReadOnlyList<CombinedClusterCandidate> BuildCombinedClusters(IReadOnlyList<ClusterSleeveInfo> clusterSleeves);
    }
}
