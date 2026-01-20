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
        /// Reset database flags for cluster sleeves that were deleted from the model.
        /// </summary>
        void ResetClusterFlagsForDeletedSleeves(Document doc, string? xmlFilePath = null);
    }
}
