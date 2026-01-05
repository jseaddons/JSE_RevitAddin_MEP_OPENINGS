using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm
{
    /// <summary>
    /// Phase 8: Interface for clustering algorithms (spatial grid + XML-based proximity + flood expansion).
    /// Extracted from UniversalClusterService to reduce monolith size.
    /// </summary>
    public interface IClusterAlgorithmService
    {
        /// <summary>
        /// Form clusters for already grouped sleeves (grouping done externally by host/system/orientation).
        /// Preserves multi-threading (parallel over groups). Returns dictionary keyed by group with cluster lists.
        /// </summary>
        Dictionary<SleeveGroupKey, List<List<ClusteringSleeveDto>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, ClusteringSleeveDto>> sleeveGroups,
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
