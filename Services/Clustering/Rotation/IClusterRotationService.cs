using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation
{
    /// <summary>
    /// Interface for cluster rotation service
    /// Handles rotation angle determination and rotated bounding box calculations
    /// </summary>
    public interface IClusterRotationService
    {
        /// <summary>
        /// Determine the dominant rotation angle for a cluster of sleeves
        /// </summary>
        /// <param name="cluster">List of sleeve data (dynamic objects with ClashZone property)</param>
        /// <param name="xmlFilePath">Optional XML file path for data access</param>
        /// <returns>Rotation angle in radians (0 for axis-aligned clusters)</returns>
        double DetermineRotationAngle(List<dynamic> cluster, string? xmlFilePath = null);

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

        /// <summary>
        /// Get rotation data for a specific cluster sleeve
        /// </summary>
        /// <param name="clusterInstanceId">Cluster sleeve instance ID</param>
        /// <returns>Tuple of rotation data or null if not found</returns>
        (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)? GetRotationData(int clusterInstanceId);

        /// <summary>
        /// Store rotation data for a cluster sleeve
        /// </summary>
        /// <param name="clusterInstanceId">Cluster sleeve instance ID</param>
        /// <param name="rotationAngleDeg">Rotation angle in degrees</param>
        /// <param name="isRotated">Whether the cluster is rotated</param>
        /// <param name="rotatedBboxMin">Rotated bounding box minimum point</param>
        /// <param name="rotatedBboxMax">Rotated bounding box maximum point</param>
        /// <param name="rotatedWidth">Rotated width</param>
        /// <param name="rotatedHeight">Rotated height</param>
        /// <param name="rotatedDepth">Rotated depth</param>
        void StoreRotationData(int clusterInstanceId, double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth);

        /// <summary>
        /// Clear all stored rotation data
        /// </summary>
        void ClearRotationData();
    }
}
