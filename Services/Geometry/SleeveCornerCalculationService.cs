using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Geometry
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for calculating sleeve corner coordinates.
    /// Pure mathematical operations - no database or Revit API dependencies.
    /// Can be called in parallel safely.
    /// </summary>
    public class SleeveCornerCalculationService
    {
        /// <summary>
        /// ✅ SRP: Calculates 4 corner coordinates in WORLD space from placement point, dimensions, and rotation.
        /// Pure math operation - thread-safe and can be parallelized.
        /// </summary>
        /// <param name="placementPoint">Placement point of the sleeve (center)</param>
        /// <param name="width">Width of the sleeve (in internal units, feet)</param>
        /// <param name="height">Height of the sleeve (in internal units, feet)</param>
        /// <param name="rotationAngleRad">Rotation angle in radians (around Z-axis)</param>
        /// <returns>Tuple of 4 corner coordinates (corner1, corner2, corner3, corner4) or null if invalid</returns>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCorners(
            XYZ placementPoint, double width, double height, double rotationAngleRad)
        {
            try
            {
                if (placementPoint == null)
                    return null;

                if (width <= 0 || height <= 0)
                    return null;

                // ✅ Calculate corner offsets in local coordinate system (before rotation)
                double halfWidth = width / 2.0;
                double halfHeight = height / 2.0;
                // Note: Z coordinate will be set to placementPoint.Z for all corners (2D opening)

                // ✅ CORNER ORDER (per methodology line 872): 1=Bottom-left, 2=Bottom-right, 3=Top-left, 4=Top-right
                // Local corner offsets (before rotation) - Z=0 in local space, will be translated to placementPoint.Z
                var localCorners = new[]
                {
                    new XYZ(-halfWidth, -halfHeight, 0),  // Corner 1: Bottom-left
                    new XYZ(halfWidth, -halfHeight, 0),    // Corner 2: Bottom-right
                    new XYZ(-halfWidth, halfHeight, 0),    // Corner 3: Top-left
                    new XYZ(halfWidth, halfHeight, 0)      // Corner 4: Top-right
                };

                // ✅ Apply rotation around Z-axis (if rotation angle is non-zero)
                var worldCorners = new XYZ[4];
                if (Math.Abs(rotationAngleRad) > 1e-6)
                {
                    // Rotation matrix for Z-axis rotation
                    double cos = Math.Cos(rotationAngleRad);
                    double sin = Math.Sin(rotationAngleRad);

                    for (int i = 0; i < 4; i++)
                    {
                        var local = localCorners[i];
                        // Rotate around Z-axis
                        double rotatedX = local.X * cos - local.Y * sin;
                        double rotatedY = local.X * sin + local.Y * cos;
                        // Translate to world coordinates (use placementPoint.Z for all corners - 2D opening)
                        worldCorners[i] = new XYZ(
                            placementPoint.X + rotatedX,
                            placementPoint.Y + rotatedY,
                            placementPoint.Z
                        );
                    }
                }
                else
                {
                    // No rotation - just translate to world coordinates (use placementPoint.Z for all corners - 2D opening)
                    for (int i = 0; i < 4; i++)
                    {
                        var local = localCorners[i];
                        worldCorners[i] = new XYZ(
                            placementPoint.X + local.X,
                            placementPoint.Y + local.Y,
                            placementPoint.Z
                        );
                    }
                }

                return (worldCorners[0], worldCorners[1], worldCorners[2], worldCorners[3]);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[SleeveCornerCalculationService] Error calculating corners: {ex.Message}");
                }
                return null;
            }
        }

        /// <summary>
        /// ✅ SRP: Calculates corners from a ClashZone (convenience method).
        /// </summary>
        public (XYZ corner1, XYZ corner2, XYZ corner3, XYZ corner4)? CalculateCornersFromZone(
            ClashZone zone, double width, double height)
        {
            if (zone == null || zone.SleevePlacementPoint == null)
                return null;

            return CalculateCorners(
                zone.SleevePlacementPoint,
                width,
                height,
                zone.MepElementRotationAngle);
        }
    }
}

