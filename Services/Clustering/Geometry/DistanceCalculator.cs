using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// Calculates minimum distances between 2D and 3D bounding boxes.
    /// Phase 1: Extracted from UniversalClusterService with crash-safe guards.
    /// </summary>
    public static class DistanceCalculator
    {
        /// <summary>
        /// Calculate minimum distance between two 2D rectangles.
        /// Returns 0 if rectangles overlap.
        /// </summary>
        /// <param name="minX1">Minimum X of first rectangle</param>
        /// <param name="minY1">Minimum Y of first rectangle</param>
        /// <param name="maxX1">Maximum X of first rectangle</param>
        /// <param name="maxY1">Maximum Y of first rectangle</param>
        /// <param name="minX2">Minimum X of second rectangle</param>
        /// <param name="minY2">Minimum Y of second rectangle</param>
        /// <param name="maxX2">Maximum X of second rectangle</param>
        /// <param name="maxY2">Maximum Y of second rectangle</param>
        /// <returns>Minimum distance, or null if inputs are invalid (NaN, Infinity, or max < min)</returns>
        public static double? CalculateMinimumDistance2D(
            double minX1, double minY1, double maxX1, double maxY1,
            double minX2, double minY2, double maxX2, double maxY2)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs (fail-fast)
                if (!IsValidCoordinate(minX1) || !IsValidCoordinate(minY1) ||
                    !IsValidCoordinate(maxX1) || !IsValidCoordinate(maxY1) ||
                    !IsValidCoordinate(minX2) || !IsValidCoordinate(minY2) ||
                    !IsValidCoordinate(maxX2) || !IsValidCoordinate(maxY2))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log", 
                        $"[DistanceCalculator] Invalid coordinates detected in CalculateMinimumDistance2D: " +
                        $"Rect1=({minX1}, {minY1}, {maxX1}, {maxY1}), Rect2=({minX2}, {minY2}, {maxX2}, {maxY2})");
                    return null;
                }

                // ✅ CRASH-SAFE: Validate rectangle bounds (max >= min)
                if (maxX1 < minX1 || maxY1 < minY1 || maxX2 < minX2 || maxY2 < minY2)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[DistanceCalculator] Invalid rectangle bounds: Rect1(maxX={maxX1} < minX={minX1} or maxY={maxY1} < minY={minY1}), " +
                        $"Rect2(maxX={maxX2} < minX={minX2} or maxY={maxY2} < minY={minY2})");
                    return null;
                }

                // Check if rectangles overlap
                bool xOverlap = !(maxX1 < minX2 || maxX2 < minX1);
                bool yOverlap = !(maxY1 < minY2 || maxY2 < minY1);

                if (xOverlap && yOverlap)
                {
                    return 0; // Rectangles overlap
                }

                // Calculate minimum distance
                double dx = 0;
                double dy = 0;

                if (!xOverlap)
                {
                    dx = Math.Min(Math.Abs(maxX1 - minX2), Math.Abs(maxX2 - minX1));
                }

                if (!yOverlap)
                {
                    dy = Math.Min(Math.Abs(maxY1 - minY2), Math.Abs(maxY2 - minY1));
                }

                double distance = Math.Sqrt(dx * dx + dy * dy);

                // ✅ CRASH-SAFE: Validate result
                if (!IsValidCoordinate(distance))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[DistanceCalculator] Calculated invalid distance: {distance} from Rect1=({minX1}, {minY1}, {maxX1}, {maxY1}), Rect2=({minX2}, {minY2}, {maxX2}, {maxY2})");
                    return null;
                }

                return distance;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[DistanceCalculator] Exception in CalculateMinimumDistance2D: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Calculate minimum distance between two 3D bounding boxes.
        /// Returns 0 if bounding boxes overlap.
        /// </summary>
        /// <param name="minX1">Minimum X of first bounding box</param>
        /// <param name="minY1">Minimum Y of first bounding box</param>
        /// <param name="minZ1">Minimum Z of first bounding box</param>
        /// <param name="maxX1">Maximum X of first bounding box</param>
        /// <param name="maxY1">Maximum Y of first bounding box</param>
        /// <param name="maxZ1">Maximum Z of first bounding box</param>
        /// <param name="minX2">Minimum X of second bounding box</param>
        /// <param name="minY2">Minimum Y of second bounding box</param>
        /// <param name="minZ2">Minimum Z of second bounding box</param>
        /// <param name="maxX2">Maximum X of second bounding box</param>
        /// <param name="maxY2">Maximum Y of second bounding box</param>
        /// <param name="maxZ2">Maximum Z of second bounding box</param>
        /// <returns>Minimum distance, or null if inputs are invalid (NaN, Infinity, or max < min)</returns>
        public static double? CalculateMinimumDistance3D(
            double minX1, double minY1, double minZ1, double maxX1, double maxY1, double maxZ1,
            double minX2, double minY2, double minZ2, double maxX2, double maxY2, double maxZ2)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs (fail-fast)
                if (!IsValidCoordinate(minX1) || !IsValidCoordinate(minY1) || !IsValidCoordinate(minZ1) ||
                    !IsValidCoordinate(maxX1) || !IsValidCoordinate(maxY1) || !IsValidCoordinate(maxZ1) ||
                    !IsValidCoordinate(minX2) || !IsValidCoordinate(minY2) || !IsValidCoordinate(minZ2) ||
                    !IsValidCoordinate(maxX2) || !IsValidCoordinate(maxY2) || !IsValidCoordinate(maxZ2))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[DistanceCalculator] Invalid coordinates detected in CalculateMinimumDistance3D: " +
                        $"BBox1=({minX1}, {minY1}, {minZ1}, {maxX1}, {maxY1}, {maxZ1}), " +
                        $"BBox2=({minX2}, {minY2}, {minZ2}, {maxX2}, {maxY2}, {maxZ2})");
                    return null;
                }

                // ✅ CRASH-SAFE: Validate bounding box bounds (max >= min)
                if (maxX1 < minX1 || maxY1 < minY1 || maxZ1 < minZ1 ||
                    maxX2 < minX2 || maxY2 < minY2 || maxZ2 < minZ2)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[DistanceCalculator] Invalid bounding box bounds: BBox1(max < min) or BBox2(max < min)");
                    return null;
                }

                // Check if bounding boxes overlap
                bool xOverlap = !(maxX1 < minX2 || maxX2 < minX1);
                bool yOverlap = !(maxY1 < minY2 || maxY2 < minY1);
                bool zOverlap = !(maxZ1 < minZ2 || maxZ2 < minZ1);

                if (xOverlap && yOverlap && zOverlap)
                {
                    return 0; // Bounding boxes overlap
                }

                // Calculate minimum distance
                double dx = 0;
                double dy = 0;
                double dz = 0;

                if (!xOverlap)
                {
                    dx = Math.Min(Math.Abs(maxX1 - minX2), Math.Abs(maxX2 - minX1));
                }

                if (!yOverlap)
                {
                    dy = Math.Min(Math.Abs(maxY1 - minY2), Math.Abs(maxY2 - minY1));
                }

                if (!zOverlap)
                {
                    dz = Math.Min(Math.Abs(maxZ1 - minZ2), Math.Abs(maxZ2 - minZ1));
                }

                double distance = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                // ✅ CRASH-SAFE: Validate result
                if (!IsValidCoordinate(distance))
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[DistanceCalculator] Calculated invalid 3D distance: {distance}");
                    return null;
                }

                return distance;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[DistanceCalculator] Exception in CalculateMinimumDistance3D: {ex.Message}, StackTrace: {ex.StackTrace}");
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

