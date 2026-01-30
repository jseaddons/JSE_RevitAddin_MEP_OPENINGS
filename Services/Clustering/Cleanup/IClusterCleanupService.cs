using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup
{
    /// <summary>
    /// Interface for cluster cleanup operations: deleting individual sleeves inside clusters and resetting flags.
    /// </summary>
    public interface IClusterCleanupService
    {
        /// <summary>
        /// Delete individual sleeves that fall within placed cluster sleeves. Returns count deleted.
        /// </summary>
        /// <param name="deferredParameters">Optional deferred parameters dictionary to read correct dimensions when batching is enabled</param>
        int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> placedClusters, Dictionary<ElementId, Dictionary<string, object>> deferredParameters = null, string targetCategory = null);

        /// <summary>
        /// ✅ NEW: DB-ONLY cleanup - Uses bounding boxes from database (zero Revit queries).
        /// Finds individual sleeves within cluster bounding boxes using DB data only.
        /// Combines Stage 1 and Stage 2 cleanup into one operation.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="targetCategory">MEP category to filter (e.g., "Ducts", "Pipes")</param>
        /// <param name="clusterInstanceIds">Optional list of cluster instance IDs to check. If null, checks all clusters.</param>
        /// <returns>Number of sleeves deleted</returns>
        int CleanupSleevesWithinClustersFromDatabase(Document doc, string targetCategory = null, List<int> clusterInstanceIds = null);

        /// <summary>
        /// Reset database flags for cluster sleeves that were deleted from the model.
        /// </summary>
        void ResetClusterFlagsForDeletedSleeves(Document doc, string? xmlFilePath = null);
    }
}
