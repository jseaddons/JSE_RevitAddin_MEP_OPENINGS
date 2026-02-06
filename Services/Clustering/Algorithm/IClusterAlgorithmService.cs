using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm
{
    /// <summary>
    /// Phase 8: Interface for clustering algorithms (spatial grid + XML-based proximity + flood expansion).
    /// Extracted from UniversalClusterService to reduce monolith size.
    /// </summary>
    public interface IClusterAlgorithmService
    {
        /// <summary>
        /// Optimized FormClusters for ClashZoneWorkItem (Batch Mode).
        /// </summary>
        Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.ClashZoneWorkItem>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel);

        /// <summary>
        /// Legacy FormClusters for dynamic grouping (RefactoredClusterService).
        /// </summary>
        Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, dynamic>> sleeveGroups,
            double toleranceDist,
            Document doc,
            bool enableParallel);

        /// <summary>
        /// Build spatial grid for family instances (used by Revit API based clustering path).
        /// </summary>
        Dictionary<(int x, int y, int z), List<FamilyInstance>> BuildSpatialGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize);

        /// <summary>
        /// Cluster sleeves from spatial grid (flood-fill expansion) – returns clusters including singletons.
        /// </summary>
        List<List<FamilyInstance>> FormClustersFromGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<(int x, int y, int z), List<FamilyInstance>> grid,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize,
            double toleranceDist,
            SleeveGroupKey groupKey);
    }
}
