using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// Transforms coordinates between world space and local/rotated coordinate spaces.
    /// Phase 1: Extracted from UniversalClusterService with crash-safe guards.
    /// </summary>
    public static class CoordinateTransformer
    {
        /// <summary>
        /// Transforms a point from local coordinate space to world space using rotation and translation.
        /// </summary>
        /// <param name="localPoint">Point in local coordinate space</param>
        /// <param name="origin">Origin of local coordinate system in world space</param>
        /// <param name="cos">Cosine of rotation angle</param>
        /// <param name="sin">Sine of rotation angle</param>
        /// <returns>Point in world coordinates, or null if inputs are invalid</returns>
        public static XYZ? TransformToWorld(XYZ localPoint, XYZ origin, double cos, double sin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (localPoint == null || origin == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Null inputs in TransformToWorld: localPoint={localPoint}, origin={origin}");
                    return null;
                }

                if (!IsValidCoordinate(localPoint.X) || !IsValidCoordinate(localPoint.Y) || !IsValidCoordinate(localPoint.Z) ||
                    !IsValidCoordinate(origin.X) || !IsValidCoordinate(origin.Y) || !IsValidCoordinate(origin.Z) ||
                    !IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Invalid coordinates in TransformToWorld: " +
                        $"localPoint=({localPoint.X}, {localPoint.Y}, {localPoint.Z}), origin=({origin.X}, {origin.Y}, {origin.Z}), cos={cos}, sin={sin}");
                    return null;
                }

                // Apply rotation in 2D (X-Y plane)
                double rotatedX = localPoint.X * cos - localPoint.Y * sin;
                double rotatedY = localPoint.X * sin + localPoint.Y * cos;

                // Translate to world space
                double worldX = origin.X + rotatedX;
                double worldY = origin.Y + rotatedY;
                double worldZ = origin.Z + localPoint.Z;  // Z stays the same

                // ✅ CRASH-SAFE: Validate results
                if (!IsValidCoordinate(worldX) || !IsValidCoordinate(worldY) || !IsValidCoordinate(worldZ))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Calculated invalid world coordinates: ({worldX}, {worldY}, {worldZ})");
                    return null;
                }

                return new XYZ(worldX, worldY, worldZ);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CoordinateTransformer] Exception in TransformToWorld: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Transforms a point from world space to local coordinate space using inverse rotation and translation.
        /// </summary>
        /// <param name="worldPoint">Point in world coordinates</param>
        /// <param name="origin">Origin of local coordinate system in world space</param>
        /// <param name="cos">Cosine of rotation angle</param>
        /// <param name="sin">Sine of rotation angle</param>
        /// <returns>Point in local coordinates, or null if inputs are invalid</returns>
        public static XYZ? TransformToLocal(XYZ worldPoint, XYZ origin, double cos, double sin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (worldPoint == null || origin == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Null inputs in TransformToLocal: worldPoint={worldPoint}, origin={origin}");
                    return null;
                }

                if (!IsValidCoordinate(worldPoint.X) || !IsValidCoordinate(worldPoint.Y) || !IsValidCoordinate(worldPoint.Z) ||
                    !IsValidCoordinate(origin.X) || !IsValidCoordinate(origin.Y) || !IsValidCoordinate(origin.Z) ||
                    !IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Invalid coordinates in TransformToLocal: " +
                        $"worldPoint=({worldPoint.X}, {worldPoint.Y}, {worldPoint.Z}), origin=({origin.X}, {origin.Y}, {origin.Z}), cos={cos}, sin={sin}");
                    return null;
                }

                // Translate relative to origin
                double relX = worldPoint.X - origin.X;
                double relY = worldPoint.Y - origin.Y;
                double relZ = worldPoint.Z - origin.Z;

                // Apply inverse rotation in 2D (X-Y plane)
                double localX = relX * cos + relY * sin;  // Inverse: transpose of rotation matrix
                double localY = -relX * sin + relY * cos;

                // ✅ CRASH-SAFE: Validate results
                if (!IsValidCoordinate(localX) || !IsValidCoordinate(localY) || !IsValidCoordinate(relZ))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Calculated invalid local coordinates: ({localX}, {localY}, {relZ})");
                    return null;
                }

                return new XYZ(localX, localY, relZ);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CoordinateTransformer] Exception in TransformToLocal: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Transforms a 2D point using rotation matrix components.
        /// This is a simpler version that only handles X-Y plane transformations.
        /// </summary>
        /// <param name="point">Input point (x, y)</param>
        /// <param name="cos">Cosine of rotation angle</param>
        /// <param name="sin">Sine of rotation angle</param>
        /// <returns>Transformed point, or null if inputs are invalid</returns>
        public static (double x, double y)? TransformPoint((double x, double y) point, double cos, double sin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (!IsValidCoordinate(point.x) || !IsValidCoordinate(point.y) ||
                    !IsValidCoordinate(cos) || !IsValidCoordinate(sin))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Invalid inputs in TransformPoint: point=({point.x}, {point.y}), cos={cos}, sin={sin}");
                    return null;
                }

                double transformedX = point.x * cos - point.y * sin;
                double transformedY = point.x * sin + point.y * cos;

                // ✅ CRASH-SAFE: Validate results
                if (!IsValidCoordinate(transformedX) || !IsValidCoordinate(transformedY))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CoordinateTransformer] Calculated invalid transformed point: ({transformedX}, {transformedY})");
                    return null;
                }

                return (transformedX, transformedY);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CoordinateTransformer] Exception in TransformPoint: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
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

