using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup
{
    /// <summary>
    /// Interface for cluster cleanup operations
    /// </summary>
    public interface IClusterCleanupService
    {
        /// <summary>
        /// Cleanup individual sleeves within clusters using database-only approach (Stage 2 cleanup)
        /// Checks if individual sleeve placement points are inside cluster bounding boxes
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="targetCategory">Optional category filter (null = all categories)</param>
        /// <param name="clusterInstanceIds">Optional list of specific cluster IDs to check against (null = all clusters)</param>
        /// <returns>Number of individual sleeves deleted</returns>
        int CleanupSleevesWithinClustersFromDatabase(Document doc, string targetCategory = null, List<int> clusterInstanceIds = null);

        /// <summary>
        /// Cleanup individual sleeves within the bounding boxes of the provided cluster instances.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="clusters">List of placed cluster family instances</param>
        /// <param name="deferredParameters">Optional dictionary of deferred parameters (to ensure correct dimensions)</param>
        /// <param name="targetCategory">Optional category filter</param>
        /// <returns>Number of individual sleeves deleted</returns>
        int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> clusters, Dictionary<ElementId, Dictionary<string, object>> deferredParameters = null, string targetCategory = null);
    }
}
