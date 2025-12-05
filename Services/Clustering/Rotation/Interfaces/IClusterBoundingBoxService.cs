using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for calculating cluster bounding boxes
    /// Single Responsibility: Calculate rotated bounding box dimensions and extents
    /// </summary>
    public interface IClusterBoundingBoxService
    {
        /// <summary>
        /// Calculate cluster bounding box using rotated coordinates from ClashZone data
        /// </summary>
        /// <param name="cluster">List of sleeve data</param>
        /// <param name="actualSleeves">List of actual family instances</param>
        /// <param name="rotationAngle">Rotation angle in radians</param>
        /// <param name="xmlFilePath">Optional XML file path for data access</param>
        /// <returns>Tuple containing width, height, depth, midpoint, and optional rotated bbox coordinates</returns>
        (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) 
        CalculateRotatedBoundingBox(List<dynamic> cluster, List<FamilyInstance> actualSleeves, double rotationAngle, string? xmlFilePath = null);
    }
}

