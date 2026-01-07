using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository specifically for retrieving data needed for Marking (Prefix/Numbering).
    /// Handles the complexity of joining ClashZones with ClusterSleeves and Snapshots.
    /// </summary>
    public interface IMarkDataRepository
    {
        /// <summary>
        /// Retrieves all clash zones (Individual and Cluster) that match the given category.
        /// Handles the invisible cluster problem by joining with ClusterSleeves.
        /// </summary>
        /// <param name="category">The MEP category to filter by (e.g., "Pipes")</param>
        /// <returns>List of ClashZones populated with Snapshot data.</returns>
        List<ClashZone> GetMarkableClashZones(string category);
        List<ClashZone> GetSleevesForLevel(string levelName, string category);
        Dictionary<int, string> GetCategoryLookup(IEnumerable<int> instanceIds);
    }
}
