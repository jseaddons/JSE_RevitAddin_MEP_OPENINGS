using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Persist combined cluster updates to database + XML.
    /// Batching-aware (collects updates, returns list for batch transaction).
    /// Phase 4 Component.
    /// </summary>
    public interface ICombinedClusterPersistence
    {
        /// <summary>
        /// Prepare database updates for combined cluster (no write yet, just queue).
        /// Returns list of ClashZone updates to execute in batch transaction.
        /// </summary>
        List<ClashZone> QueueDatabaseUpdates(
            CombinedClusterCandidate combinedCluster,
            int combinedSleeveInstanceId);
        
        /// <summary>
        /// Update XML files with combined cluster information.
        /// Called AFTER Revit family instance created (has valid ID).
        /// </summary>
        void UpdateXmlWithCombinedClusterInfo(
            CombinedClusterCandidate combinedCluster,
            int combinedSleeveInstanceId);
    }
}
