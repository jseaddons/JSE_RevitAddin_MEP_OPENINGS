using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation
{
    /// <summary>
    /// Service for handling cluster rotation calculations and rotated bounding box determinations
    /// Extracted from UniversalClusterService for better separation of concerns
    /// </summary>
    public class ClusterRotationService : IClusterRotationService
    {
        // Store rotation angle and rotated bounding box for each cluster sleeve
        // Key: ClusterInstanceId
        private readonly Dictionary<int, (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)> _clusterRotationData;

        // Delegate for getting ClashZone by sleeve instance ID (injected dependency)
        private readonly Func<int, string, ClashZone> _getClashZoneFunc;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="getClashZoneFunc">Function to retrieve ClashZone by sleeve instance ID</param>
        public ClusterRotationService(Func<int, string, ClashZone> getClashZoneFunc)
        {
            _clusterRotationData = new Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>();
            _getClashZoneFunc = getClashZoneFunc ?? throw new ArgumentNullException(nameof(getClashZoneFunc));
        }

        /// <summary>
        /// Determine the dominant rotation angle for a cluster of sleeves
        /// Returns 0 for axis-aligned clusters
        /// </summary>
        public double DetermineRotationAngle(List<dynamic> cluster, string? xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0)
                    return 0.0;

                var rotationAngles = new List<double>();

                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    var clashZone = sleeveData.ClashZone as ClashZone;
                    if (clashZone == null)
                        continue;

                    // ✅ ROTATION DATA FROM DB: Get rotation angle from clash zone (loaded from database).
                    // MepElementRotationAngle is saved to database during refresh (InsertOrUpdateClashZones)
                    // and loaded by GetClashZonesByCategory -> MepRotationAngleRad column.
                    double angle = clashZone.MepElementRotationAngle;
                    
                    // Normalize angle to 0-2π range
                    while (angle < 0) angle += 2 * Math.PI;
                    while (angle >= 2 * Math.PI) angle -= 2 * Math.PI;
                    
                    rotationAngles.Add(angle);
                }

                if (rotationAngles.Count == 0)
                    return 0.0;

                // ✅ STRATEGY: Use average angle (works well for similar angles)
                double averageAngle = rotationAngles.Average();
                
                // Check if angles wrap around 0°/360° boundary
                double minAngle = rotationAngles.Min();
                double maxAngle = rotationAngles.Max();
                if (maxAngle - minAngle > Math.PI)
                {
                    // Angles wrap around - adjust by adding 2π to angles < π
                    var adjustedAngles = rotationAngles.Select(a => a < Math.PI ? a + 2 * Math.PI : a).ToList();
                    averageAngle = adjustedAngles.Average();
                    if (averageAngle >= 2 * Math.PI)
                        averageAngle -= 2 * Math.PI;
                }

                // ✅ CRITICAL FIX: Check if angle is essentially axis-aligned (0°, 90°, 180°, 270°)
                // If so, return 0 to use axis-aligned bounding box logic
                double thresholdDegrees = 2.0; // 2 degree tolerance
                
                // Helper function to check if an angle is axis-aligned
                bool IsAxisAligned(double angleRad)
                {
                    double angleDeg = angleRad * 180 / Math.PI;
                    while (angleDeg < 0) angleDeg += 360;
                    while (angleDeg >= 360) angleDeg -= 360;
                    
                    double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
                    double distTo90 = Math.Abs(angleDeg - 90);
                    double distTo180 = Math.Abs(angleDeg - 180);
                    double distTo270 = Math.Abs(angleDeg - 270);
                    
                    return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                           distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
                }
                
                // Check if ALL individual angles are axis-aligned
                bool allAxisAligned = rotationAngles.All(IsAxisAligned);
                
                if (allAxisAligned)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] All {rotationAngles.Count} angles are axis-aligned: [{angleList}], using axis-aligned bounding box");
                    }
                    return 0.0;
                }
                
                // Check if average angle is axis-aligned
                if (IsAxisAligned(averageAngle))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double angleDegrees = averageAngle * 180 / Math.PI;
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] Average angle {angleDegrees:F1}° is axis-aligned (angles: [{angleList}]), using axis-aligned bounding box");
                    }
                    return 0.0;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ANGLE] Determined dominant rotation angle: {averageAngle * 180 / Math.PI:F1}° from {rotationAngles.Count} sleeves (non-axis-aligned)");
                }

                return averageAngle;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ClusterRotationService] Error determining dominant rotation angle: {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Calculate cluster bounding box using rotated coordinates from ClashZone data
        /// This is a simplified version - full implementation should be migrated from UniversalClusterService
        /// </summary>
        public (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)
        CalculateRotatedBoundingBox(List<dynamic> cluster, List<FamilyInstance> actualSleeves, double rotationAngle, string? xmlFilePath = null)
        {
            // Simplified rotated bounding box calculation (union + optional corner refinement) without external ambiguous loggers.
            if (cluster == null || cluster.Count == 0 || actualSleeves == null || actualSleeves.Count == 0)
                return (0,0,0,XYZ.Zero,null,null,null,null,null,null);

            var rotatedBboxes = new List<(XYZ min, XYZ max)>();
            foreach (var sleeveData in cluster)
            {
                var cz = _getClashZoneFunc(sleeveData.SleeveInstanceId, xmlFilePath);
                if (cz == null) continue;
                bool hasRot = cz.RotatedBoundingBoxMinX.HasValue && cz.RotatedBoundingBoxMaxX.HasValue && cz.RotatedBoundingBoxMinY.HasValue && cz.RotatedBoundingBoxMaxY.HasValue;
                if (hasRot)
                {
                    rotatedBboxes.Add((
                        new XYZ(cz.RotatedBoundingBoxMinX!.Value, cz.RotatedBoundingBoxMinY!.Value, cz.RotatedBoundingBoxMinZ ?? cz.SleeveBoundingBoxMinZ),
                        new XYZ(cz.RotatedBoundingBoxMaxX!.Value, cz.RotatedBoundingBoxMaxY!.Value, cz.RotatedBoundingBoxMaxZ ?? cz.SleeveBoundingBoxMaxZ)
                    ));
                }
            }

            if (rotatedBboxes.Count == 0)
            {
                // Fallback: axis-aligned union from Revit bounding boxes
                var revitBboxes = actualSleeves.Select(s => s.get_BoundingBox(null)).Where(b => b != null && b.Enabled).ToList();
                if (revitBboxes.Count == 0) return (0,0,0,XYZ.Zero,null,null,null,null,null,null);
                double minXf = revitBboxes.Min(b=>b.Min.X); double minYf = revitBboxes.Min(b=>b.Min.Y); double minZf = revitBboxes.Min(b=>b.Min.Z);
                double maxXf = revitBboxes.Max(b=>b.Max.X); double maxYf = revitBboxes.Max(b=>b.Max.Y); double maxZf = revitBboxes.Max(b=>b.Max.Z);
                double wf = maxXf - minXf; double hf = maxYf - minYf; double df = maxZf - minZf; XYZ midF = new XYZ((minXf+maxXf)/2,(minYf+maxYf)/2,(minZf+maxZf)/2);
                return (wf,hf,df,midF,null,null,null,null,null,null);
            }

            // Simple union of rotated boxes (immediate correctness for most cases). Corner refinement optional.
            double minX = rotatedBboxes.Min(b=>b.min.X);
            double minY = rotatedBboxes.Min(b=>b.min.Y);
            double minZ = rotatedBboxes.Min(b=>b.min.Z);
            double maxX = rotatedBboxes.Max(b=>b.max.X);
            double maxY = rotatedBboxes.Max(b=>b.max.Y);
            double maxZ = rotatedBboxes.Max(b=>b.max.Z);

            double width = maxX - minX;
            double height = maxY - minY;
            double depth = maxZ - minZ;
            XYZ mid = new XYZ((minX+maxX)/2,(minY+maxY)/2,(minZ+maxZ)/2);
            return (width,height,depth,mid,minX,minY,minZ,maxX,maxY,maxZ);
        }

        /// <summary>
        /// Get stored rotation data for a cluster sleeve
        /// </summary>
        public (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)? GetRotationData(int clusterInstanceId)
        {
            if (_clusterRotationData.TryGetValue(clusterInstanceId, out var data))
            {
                return data;
            }
            return null;
        }

        /// <summary>
        /// Store rotation data for a cluster sleeve
        /// </summary>
        public void StoreRotationData(int clusterInstanceId, double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)
        {
            _clusterRotationData[clusterInstanceId] = (rotationAngleDeg, isRotated, rotatedBboxMin, rotatedBboxMax, rotatedWidth, rotatedHeight, rotatedDepth);
        }

        /// <summary>
        /// Clear all stored rotation data
        /// </summary>
        public void ClearRotationData()
        {
            _clusterRotationData.Clear();
        }
    }
}
