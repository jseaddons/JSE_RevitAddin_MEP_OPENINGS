using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Interface for clustering strategies that handle proximity checking and cluster formation.
    /// Phase 4A: Extracted from UniversalClusterService for strategy pattern implementation.
    /// </summary>
    public interface IClusteringStrategy
    {
        /// <summary>
        /// Check if two sleeves are in proximity based on strategy-specific logic.
        /// </summary>
        /// <param name="sleeve1">First sleeve (dynamic type with ClashZone and BoundingBox)</param>
        /// <param name="sleeve2">Second sleeve (dynamic type with ClashZone and BoundingBox)</param>
        /// <param name="tolerance">Maximum distance for sleeves to be considered in proximity</param>
        /// <param name="orientation">Orientation for context-specific checks (X, Y, Floor, etc.)</param>
        /// <param name="document">Document for host validation (e.g., wall host checking)</param>
        /// <returns>True if sleeves are in proximity, false otherwise</returns>
        bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance, string orientation, Autodesk.Revit.DB.Document document = null);

        /// <summary>
        /// Calculate proximity distance between two sleeves using strategy-specific logic.
        /// </summary>
        /// <param name="sleeve1">First sleeve</param>
        /// <param name="sleeve2">Second sleeve</param>
        /// <param name="orientation">Orientation for context-specific calculations</param>
        /// <returns>Distance between sleeves (or double.MaxValue if calculation failed)</returns>
        double CalculateProximityDistance(dynamic sleeve1, dynamic sleeve2, string orientation);

        /// <summary>
        /// Form clusters from list of sleeves using strategy-specific algorithm.
        /// Uses iterative expansion (flood-fill) to find all connected sleeves.
        /// </summary>
        /// <param name="sleeves">List of sleeves to cluster</param>
        /// <param name="tolerance">Maximum distance for sleeves to be considered in proximity</param>
        /// <param name="orientation">Orientation for context-specific checks</param>
        /// <param name="document">Document for host validation</param>
        /// <returns>List of clusters (each cluster is a list of sleeves)</returns>
        List<List<dynamic>> FormClusters(List<dynamic> sleeves, double tolerance, string orientation, Autodesk.Revit.DB.Document document = null);

        /// <summary>
        /// Check if this strategy can handle the given group key.
        /// </summary>
        /// <param name="groupKey">Group key containing hostType, systemType, and orientation</param>
        /// <param name="sleeves">List of sleeves in the group (for shape/rotation detection)</param>
        /// <returns>True if this strategy can handle the group, false otherwise</returns>
        bool CanHandle(SleeveGroupKey groupKey, List<dynamic> sleeves);

        /// <summary>
        /// Get strategy name for logging and diagnostics.
        /// </summary>
        string GetStrategyName();
    }
}

