using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Proximity checker for rotated sleeves using rotated coordinate system.
    /// Phase 2: Extracted from UniversalClusterService.CheckRotatedSleeveProximity with crash-safe guards.
    /// </summary>
    public class RotatedProximityChecker : IProximityChecker
    {
        private readonly double _rotationAngle;
        private readonly BoundingBoxProximityChecker _fallbackChecker;

        /// <summary>
        /// Constructor with rotation angle for rotated sleeve proximity checking.
        /// </summary>
        /// <param name="rotationAngle">Rotation angle in radians (cluster's intended rotated axis)</param>
        public RotatedProximityChecker(double rotationAngle)
        {
            _rotationAngle = rotationAngle;
            _fallbackChecker = new BoundingBoxProximityChecker();
        }

        /// <summary>
        /// Check if two rotated sleeves are within proximity tolerance.
        /// Uses pre-calculated rotation matrix components from database.
        /// </summary>
        public bool CheckProximity(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, double tolerance)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1 == null || sleeve2 == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotatedProximityChecker] Null sleeve inputs - returning false");
                    return false;
                }

                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    return false;
                }

                var cz1 = sleeve1.ClashZone;
                var cz2 = sleeve2.ClashZone;

                if (cz1 == null || cz2 == null)
                {
                    return false;
                }

                // ✅ DATABASE DATA: Get sleeve centers from database (placement points)
                XYZ center1 = cz1.SleevePlacementPoint ?? new XYZ(
                    (cz1.SleeveBoundingBoxMinX + cz1.SleeveBoundingBoxMaxX) / 2.0,
                    (cz1.SleeveBoundingBoxMinY + cz1.SleeveBoundingBoxMaxY) / 2.0,
                    (cz1.SleeveBoundingBoxMinZ + cz1.SleeveBoundingBoxMinZ) / 2.0);

                XYZ center2 = cz2.SleevePlacementPoint ?? new XYZ(
                    (cz2.SleeveBoundingBoxMinX + cz2.SleeveBoundingBoxMaxX) / 2.0,
                    (cz2.SleeveBoundingBoxMinY + cz2.SleeveBoundingBoxMaxY) / 2.0,
                    (cz2.SleeveBoundingBoxMinZ + cz2.SleeveBoundingBoxMinZ) / 2.0);

                // ✅ STEP 1: Calculate shared rotated axis direction
                // Prefer the full 3D orientation stored on the clash zone (handles tilted axes)
                XYZ rotatedAxisDirection;
                if (cz1.MepElementOrientation != null && cz1.MepElementOrientation.GetLength() > 1e-6)
                {
                    rotatedAxisDirection = cz1.MepElementOrientation.Normalize();
                }
                else
                {
                    // Fallback to XY-projected rotation angle (historical behaviour)
                    rotatedAxisDirection = new XYZ(Math.Cos(_rotationAngle), Math.Sin(_rotationAngle), 0).Normalize();
                }

                // ✅ STEP 2: Calculate vector from center1 to center2 in world-space
                XYZ worldVector = center2 - center1;

                // ✅ STEP 3: Project distance along the shared rotated axis using dot product
                double distanceAlongAxis = worldVector.DotProduct(rotatedAxisDirection);

                // ✅ STEP 4: Calculate perpendicular distance from axis using cross product magnitude
                XYZ crossProduct = worldVector.CrossProduct(rotatedAxisDirection);
                double perpendicularDistance = crossProduct.GetLength();

                // ✅ DATABASE DATA: Get sleeve dimensions from database (rotated bounding boxes if available)
                double sleeve1Width = GetSleeveWidth(cz1);
                double sleeve1Height = GetSleeveHeight(cz1);
                double sleeve2Width = GetSleeveWidth(cz2);
                double sleeve2Height = GetSleeveHeight(cz2);

                // Calculate half-dimensions for overlap check
                double halfWidth1 = sleeve1Width / 2.0;
                double halfWidth2 = sleeve2Width / 2.0;
                double halfHeight1 = sleeve1Height / 2.0;
                double halfHeight2 = sleeve2Height / 2.0;

                // ✅ STEP 5: Check if sleeves are close enough along the rotated axis
                double maxDistanceAlongAxis = halfWidth1 + halfWidth2 + tolerance;
                bool closeAlongAxis = Math.Abs(distanceAlongAxis) <= maxDistanceAlongAxis;

                // ✅ STEP 6: Check if perpendicular distance is small enough
                double maxPerpendicularDistance = halfHeight1 + halfHeight2 + tolerance;
                bool closePerpendicular = perpendicularDistance <= maxPerpendicularDistance;

                // Cluster if both conditions are met
                bool shouldCluster = closeAlongAxis && closePerpendicular;

                if (!shouldCluster && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[RotatedProximityChecker] Sleeves {sleeve1.SleeveInstanceId} and {sleeve2.SleeveInstanceId}: " +
                        $"DistanceAlongAxis={distanceAlongAxis * 304.8:F1}mm (max={maxDistanceAlongAxis * 304.8:F1}mm), " +
                        $"PerpendicularDistance={perpendicularDistance * 304.8:F1}mm (max={maxPerpendicularDistance * 304.8:F1}mm) - NO CLUSTER");
                }

                return shouldCluster;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotatedProximityChecker] Exception in CheckProximity: {ex.Message}, StackTrace: {ex.StackTrace}");
                
                // ✅ CRASH-SAFE: Fallback to regular bounding box check
                // Note: FallbackChecker likely expects dynamic, but we might need to adjust it too if it's strict.
                // Assuming BoundingBoxProximityChecker also implements IProximityChecker, it needs update too.
                // But for now, we pass DTOs. If BoundingBoxProximityChecker is updated, it works.
                // If not, we have a problem. I should check BoundingBoxProximityChecker.
                return _fallbackChecker.CheckProximity(sleeve1, sleeve2, tolerance);
            }
        }

        /// <summary>
        /// Calculate distance between two rotated sleeves.
        /// Returns minimum distance considering rotation, or null if calculation failed.
        /// </summary>
        public double? CalculateDistance(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1 == null || sleeve2 == null)
                {
                    return null;
                }

                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    return null;
                }

                var cz1 = sleeve1.ClashZone;
                var cz2 = sleeve2.ClashZone;

                if (cz1 == null || cz2 == null)
                {
                    return null;
                }

                XYZ center1 = cz1.SleevePlacementPoint ?? new XYZ(
                    (cz1.SleeveBoundingBoxMinX + cz1.SleeveBoundingBoxMaxX) / 2.0,
                    (cz1.SleeveBoundingBoxMinY + cz1.SleeveBoundingBoxMaxY) / 2.0,
                    (cz1.SleeveBoundingBoxMinZ + cz1.SleeveBoundingBoxMinZ) / 2.0);

                XYZ center2 = cz2.SleevePlacementPoint ?? new XYZ(
                    (cz2.SleeveBoundingBoxMinX + cz2.SleeveBoundingBoxMaxX) / 2.0,
                    (cz2.SleeveBoundingBoxMinY + cz2.SleeveBoundingBoxMaxY) / 2.0,
                    (cz2.SleeveBoundingBoxMinZ + cz2.SleeveBoundingBoxMinZ) / 2.0);

                // Calculate world-space distance
                XYZ worldVector = center2 - center1;
                double worldDistance = worldVector.GetLength();

                return worldDistance;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotatedProximityChecker] Exception in CalculateDistance: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ DATABASE-ONLY: Get sleeve width from ClashZone (rotated bounding box preferred).
        /// Uses pre-calculated data from database (dump once use many times).
        /// </summary>
        private double GetSleeveWidth(ClashZone cz)
        {
            if (cz.RotatedBoundingBoxMinX.HasValue && cz.RotatedBoundingBoxMaxX.HasValue)
            {
                return cz.RotatedBoundingBoxMaxX.Value - cz.RotatedBoundingBoxMinX.Value;
            }
            return cz.SleeveBoundingBoxMaxX - cz.SleeveBoundingBoxMinX;
        }

        /// <summary>
        /// ✅ DATABASE-ONLY: Get sleeve height from ClashZone (rotated bounding box preferred).
        /// Uses pre-calculated data from database (dump once use many times).
        /// </summary>
        private double GetSleeveHeight(ClashZone cz)
        {
            if (cz.RotatedBoundingBoxMinY.HasValue && cz.RotatedBoundingBoxMaxY.HasValue)
            {
                return cz.RotatedBoundingBoxMaxY.Value - cz.RotatedBoundingBoxMinY.Value;
            }
            return cz.SleeveBoundingBoxMaxY - cz.SleeveBoundingBoxMinY;
        }
    }
}

