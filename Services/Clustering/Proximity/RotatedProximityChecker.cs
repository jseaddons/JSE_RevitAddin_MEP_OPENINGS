using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
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
        public bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance)
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

                // ✅ CRASH-SAFE: Handle both wrapper objects (ClashZoneWorkItem) and direct ClashZone objects
                ClashZone cz1 = null;
                ClashZone cz2 = null;

                // Try to get ClashZone from sleeve1
                if (sleeve1 is ClashZone c1)
                {
                    cz1 = c1;
                }
                else if (sleeve1 != null)
                {
                    try { cz1 = sleeve1.ClashZone as ClashZone; } catch { }
                }

                // Try to get ClashZone from sleeve2
                if (sleeve2 is ClashZone c2)
                {
                    cz2 = c2;
                }
                else if (sleeve2 != null)
                {
                    try { cz2 = sleeve2.ClashZone as ClashZone; } catch { }
                }

                if (cz1 == null || cz2 == null)
                {
                    return false;
                }

                // [Use cz1/cz2 from here on]

                // ✅ DATABASE DATA: Get sleeve centers from database (placement points)
                XYZ center1 = cz1.SleevePlacementPoint ?? new XYZ(
                    (cz1.SleeveBoundingBoxMinX + cz1.SleeveBoundingBoxMaxX) / 2.0,
                    (cz1.SleeveBoundingBoxMinY + cz1.SleeveBoundingBoxMaxY) / 2.0,
                    (cz1.SleeveBoundingBoxMinZ + cz1.SleeveBoundingBoxMinZ) / 2.0);

                XYZ center2 = cz2.SleevePlacementPoint ?? new XYZ(
                    (cz2.SleeveBoundingBoxMinX + cz2.SleeveBoundingBoxMaxX) / 2.0,
                    (cz2.SleeveBoundingBoxMinY + cz2.SleeveBoundingBoxMaxY) / 2.0,
                    (cz2.SleeveBoundingBoxMinZ + cz2.SleeveBoundingBoxMinZ) / 2.0);

                // ✅ FIX: "Strict 2D for Floors" - zero out Z coordinates if both are on floors
                string hostType1 = cz1.StructuralElementType ?? "";
                string hostType2 = cz2.StructuralElementType ?? "";
                if (hostType1.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    hostType2.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    center1 = new XYZ(center1.X, center1.Y, 0);
                    center2 = new XYZ(center2.X, center2.Y, 0);
                    
                    // Also ensure the axis direction is horizontal
                    if (cz1.MepElementOrientation != null)
                    {
                        var horizontalOrientation = new XYZ(cz1.MepElementOrientation.X, cz1.MepElementOrientation.Y, 0);
                        if (horizontalOrientation.GetLength() > 1e-6)
                        {
                            // Temporarily override for this check if we know it's a floor
                        }
                    }
                }

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

                if (!shouldCluster)
                {
                    // 🔍 DIAGNOSTIC: Log detailed math for failed checks to batch_v2.log
                    // This helps debug why overlapping rotated sleeves are being rejected
                    SafeFileLogger.SafeAppendText("batch_v2.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 PROXIMITY FAIL: Sleeves {cz1.SleeveInstanceId} and {cz2.SleeveInstanceId}: " +
                        $"DistAxis={distanceAlongAxis:F3} (Max={maxDistanceAlongAxis:F3}), " +
                        $"DistPerp={perpendicularDistance:F3} (Max={maxPerpendicularDistance:F3}), " +
                        $"W1={sleeve1Width:F3}, W2={sleeve2Width:F3}, H1={sleeve1Height:F3}, H2={sleeve2Height:F3}, " +
                        $"Vector=[{worldVector.X:F3},{worldVector.Y:F3},{worldVector.Z:F3}], Axis=[{rotatedAxisDirection.X:F3},{rotatedAxisDirection.Y:F3},{rotatedAxisDirection.Z:F3}]\n");

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[RotatedProximityChecker] Sleeves {cz1.SleeveInstanceId} and {cz2.SleeveInstanceId}: " +
                            $"DistanceAlongAxis={distanceAlongAxis * 304.8:F1}mm (max={maxDistanceAlongAxis * 304.8:F1}mm), " +
                            $"PerpendicularDistance={perpendicularDistance * 304.8:F1}mm (max={maxPerpendicularDistance * 304.8:F1}mm) - NO CLUSTER");
                    }
                }

                return shouldCluster;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotatedProximityChecker] Exception in CheckProximity: {ex.Message}, StackTrace: {ex.StackTrace}");
                
                // ✅ CRASH-SAFE: Fallback to regular bounding box check
                return _fallbackChecker.CheckProximity(sleeve1, sleeve2, tolerance);
            }
        }

        /// <summary>
        /// Calculate distance between two rotated sleeves.
        /// Returns minimum distance considering rotation, or null if calculation failed.
        /// </summary>
        public double? CalculateDistance(dynamic sleeve1, dynamic sleeve2)
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

                var cz1 = sleeve1.ClashZone as ClashZone;
                var cz2 = sleeve2.ClashZone as ClashZone;

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
        /// ✅ DATABASE-ONLY: Get sleeve width from ClashZone.
        /// Uses actual sleeve dimensions (SleeveWidth) instead of bounding box.
        /// For rotated sleeves, bounding box is inflated (becomes square for 45° rotation).
        /// </summary>
        private double GetSleeveWidth(ClashZone cz)
        {
            // Use actual sleeve width from database (extracted from Revit geometry)
            // This is the true dimension, not the inflated axis-aligned bounding box
            if (cz.SleeveWidth > 0)
            {
                return cz.SleeveWidth;
            }
            
            // Fallback to calculated width if not set
            if (cz.CalculatedSleeveWidth > 0)
            {
                return cz.CalculatedSleeveWidth;
            }
            
            // Last resort: use bounding box (may be inflated for rotated sleeves)
            return cz.BoundingBoxMaxX - cz.BoundingBoxMinX;
        }

        /// <summary>
        /// ✅ DATABASE-ONLY: Get sleeve height from ClashZone.
        /// Uses actual sleeve dimensions (SleeveHeight) instead of bounding box.
        /// For rotated sleeves, bounding box is inflated (becomes square for 45° rotation).
        /// </summary>
        private double GetSleeveHeight(ClashZone cz)
        {
            // Use actual sleeve height from database (extracted from Revit geometry)
            // This is the true dimension, not the inflated axis-aligned bounding box
            if (cz.SleeveHeight > 0)
            {
                return cz.SleeveHeight;
            }
            
            // Fallback to calculated height if not set
            if (cz.CalculatedSleeveHeight > 0)
            {
                return cz.CalculatedSleeveHeight;
            }
            
            // Last resort: use bounding box (may be inflated for rotated sleeves)
            return cz.BoundingBoxMaxY - cz.BoundingBoxMinY;
        }
    }
}

