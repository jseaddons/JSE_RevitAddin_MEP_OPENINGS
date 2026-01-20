using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox
{
    /// <summary>
    /// Result structure for bounding box calculations.
    /// </summary>
    public struct BoundingBoxResult
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double Depth { get; set; }
        public XYZ Midpoint { get; set; }
        public double? RotatedMinX { get; set; }
        public double? RotatedMinY { get; set; }
        public double? RotatedMinZ { get; set; }
        public double? RotatedMaxX { get; set; }
        public double? RotatedMaxY { get; set; }
        public double? RotatedMaxZ { get; set; }
    }

    /// <summary>
    /// Interface for cluster bounding box calculations.
    /// Phase 3: Extracted from UniversalClusterService for calculator pattern.
    /// </summary>
    public interface IBoundingBoxCalculator
    {
        /// <summary>
        /// Calculate cluster bounding box for a list of sleeves.
        /// </summary>
        /// <param name="cluster">List of sleeves in the cluster (dynamic objects with ClashZone)</param>
        /// <param name="actualSleeves">List of actual Revit FamilyInstance sleeves</param>
        /// <param name="rotationAngle">Rotation angle in radians (0 for axis-aligned)</param>
        /// <param name="xmlFilePath">Optional XML file path for ClashZone lookup</param>
        /// <returns>BoundingBoxResult with width, height, depth, midpoint, and rotated coordinates</returns>
        BoundingBoxResult Calculate(
            List<dynamic> cluster,
            List<FamilyInstance> actualSleeves,
            double rotationAngle,
            string xmlFilePath = null);

        /// <summary>
        /// Calculate bounding box from pre-transformed corner points.
        /// Used for corner-based watertight algorithm.
        /// </summary>
        /// <param name="corners">List of corner points in rotated coordinate space</param>
        /// <returns>BoundingBoxResult with width, height, depth calculated from corner extents</returns>
        BoundingBoxResult CalculateFromCorners(List<XYZ> corners);
    }
}

