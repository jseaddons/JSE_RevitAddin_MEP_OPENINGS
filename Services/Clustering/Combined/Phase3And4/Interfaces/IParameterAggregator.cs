using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces
{
    /// <summary>
    /// SOLID Interface: Single Responsibility - Aggregate parameters from all contributing sleeves.
    /// Append model: collect snapshots from each category and combine.
    /// Phase 3 Component.
    /// </summary>
    public interface IParameterAggregator
    {
        /// <summary>
        /// Collect parameter snapshots from all cluster sleeves in combined cluster.
        /// Returns: AggregatedParameterSnapshot with per-category breakdowns.
        /// </summary>
        AggregatedParameterSnapshot AggregateParameters(
            CombinedClusterCandidate combinedCluster,
            List<ClashZone> incorporatedIndividualSleeves);
        
        /// <summary>
        /// Create parameter set for combined sleeve family instance.
        /// Uses aggregated snapshot to populate parameters.
        /// </summary>
        Dictionary<string, object> CreateCombinedSleeveParameterSet(
            AggregatedParameterSnapshot aggregatedSnapshot,
            BoundingBoxXYZ combinedBoundingBox);
    }
}
