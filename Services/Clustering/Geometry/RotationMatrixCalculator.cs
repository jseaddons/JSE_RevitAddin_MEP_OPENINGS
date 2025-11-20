using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// Calculates rotation matrices and applies rotations for coordinate transformations.
    /// Phase 1: Extracted from UniversalClusterService with crash-safe guards.
    /// </summary>
    public static class RotationMatrixCalculator
    {
        /// <summary>
        /// Creates a rotation matrix from an angle (in radians).
        /// Returns cos and sin components.
        /// </summary>
        /// <param name="angleRadians">Rotation angle in radians</param>
        /// <returns>Tuple of (cos, sin), or null if angle is invalid</returns>
        public static (double cos, double sin)? CreateRotationMatrix(double angleRadians)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate angle (not NaN, not Infinity)
                if (!IsValidAngle(angleRadians))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Invalid rotation angle: {angleRadians} radians ({angleRadians * 180.0 / Math.PI:F2} degrees)");
                    return null;
                }

                double cos = Math.Cos(angleRadians);
                double sin = Math.Sin(angleRadians);

                // ✅ CRASH-SAFE: Validate calculated components
                if (!IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Calculated invalid rotation components: cos={cos}, sin={sin} for angle={angleRadians} radians");
                    return null;
                }

                return (cos, sin);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotationMatrixCalculator] Exception in CreateRotationMatrix: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Applies rotation to a 2D point using rotation matrix components.
        /// Forward rotation: rotatedX = x * cos - y * sin, rotatedY = x * sin + y * cos
        /// </summary>
        /// <param name="point">Input point (x, y)</param>
        /// <param name="cos">Cosine of rotation angle</param>
        /// <param name="sin">Sine of rotation angle</param>
        /// <returns>Rotated point, or null if inputs are invalid</returns>
        public static (double x, double y)? ApplyRotation((double x, double y) point, double cos, double sin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (!IsValidCoordinate(point.x) || !IsValidCoordinate(point.y) ||
                    !IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Invalid inputs in ApplyRotation: point=({point.x}, {point.y}), cos={cos}, sin={sin}");
                    return null;
                }

                double rotatedX = point.x * cos - point.y * sin;
                double rotatedY = point.x * sin + point.y * cos;

                // ✅ CRASH-SAFE: Validate results
                if (!IsValidCoordinate(rotatedX) || !IsValidCoordinate(rotatedY))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Calculated invalid rotated point: ({rotatedX}, {rotatedY})");
                    return null;
                }

                return (rotatedX, rotatedY);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotationMatrixCalculator] Exception in ApplyRotation: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Applies inverse rotation (transpose of rotation matrix) to a 2D point.
        /// Inverse rotation: worldX = rotatedX * cos + rotatedY * sin, worldY = -rotatedX * sin + rotatedY * cos
        /// </summary>
        /// <param name="rotatedPoint">Rotated point (x, y) in rotated coordinate space</param>
        /// <param name="cos">Cosine of rotation angle</param>
        /// <param name="sin">Sine of rotation angle</param>
        /// <returns>Point in world coordinates, or null if inputs are invalid</returns>
        public static (double x, double y)? InverseRotation((double x, double y) rotatedPoint, double cos, double sin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (!IsValidCoordinate(rotatedPoint.x) || !IsValidCoordinate(rotatedPoint.y) ||
                    !IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Invalid inputs in InverseRotation: point=({rotatedPoint.x}, {rotatedPoint.y}), cos={cos}, sin={sin}");
                    return null;
                }

                // ✅ INVERSE TRANSFORMATION: Transpose of rotation matrix
                double worldX = rotatedPoint.x * cos + rotatedPoint.y * sin;  // Note: + instead of -
                double worldY = -rotatedPoint.x * sin + rotatedPoint.y * cos;  // Note: -sin instead of sin

                // ✅ CRASH-SAFE: Validate results
                if (!IsValidCoordinate(worldX) || !IsValidCoordinate(worldY))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotationMatrixCalculator] Calculated invalid world point: ({worldX}, {worldY})");
                    return null;
                }

                return (worldX, worldY);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotationMatrixCalculator] Exception in InverseRotation: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Validates that an angle is a valid number (not NaN, not Infinity).
        /// Allows any real number (including negative and > 2π).
        /// </summary>
        private static bool IsValidAngle(double angle)
        {
            return !double.IsNaN(angle) && !double.IsInfinity(angle);
        }

        /// <summary>
        /// Validates that a coordinate is a valid number (not NaN, not Infinity).
        /// </summary>
        private static bool IsValidCoordinate(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}

