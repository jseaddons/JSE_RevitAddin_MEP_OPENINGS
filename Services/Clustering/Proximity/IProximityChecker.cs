using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Interface for proximity checking strategies.
    /// Phase 2: Extracted from UniversalClusterService for strategy pattern.
    /// </summary>
    public interface IProximityChecker
    {
        /// <summary>
        /// Check if two sleeves are within proximity tolerance.
        /// </summary>
        /// <param name="sleeve1">First sleeve (DTO with ClashZone)</param>
        /// <param name="sleeve2">Second sleeve (DTO with ClashZone)</param>
        /// <param name="tolerance">Tolerance distance in internal units</param>
        /// <returns>True if sleeves should cluster, false otherwise</returns>
        bool CheckProximity(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, double tolerance);

        /// <summary>
        /// Calculate the minimum distance between two sleeves.
        /// </summary>
        /// <param name="sleeve1">First sleeve (DTO with ClashZone)</param>
        /// <param name="sleeve2">Second sleeve (DTO with ClashZone)</param>
        /// <returns>Minimum distance in internal units, or null if calculation failed</returns>
        double? CalculateDistance(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2);
    }
}

