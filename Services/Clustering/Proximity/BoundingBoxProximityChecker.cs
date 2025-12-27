using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Proximity checker using bounding box minimum distance calculation.
    /// Phase 2: Extracted from UniversalClusterService.BoundingBoxesOverlapFromXml with crash-safe guards.
    /// </summary>
    public class BoundingBoxProximityChecker : IProximityChecker
    {
        /// <summary>
        /// Check if two sleeves are within proximity tolerance using bounding box distance.
        /// Uses pre-calculated bounding boxes from database.
        /// </summary>
        public bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1 == null || sleeve2 == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[BoundingBoxProximityChecker] Null sleeve inputs - returning false");
                    return false;
                }

                var bbox1 = sleeve1.BoundingBox;
                var bbox2 = sleeve2.BoundingBox;

                if (bbox1 == null || bbox2 == null)
                {
                    return false;
                }

                // ✅ CRASH-SAFE: Validate bounding boxes are enabled
                if (bbox1 is BoundingBoxXYZ bbox1XYZ && !bbox1XYZ.Enabled)
                {
                    return false;
                }
                if (bbox2 is BoundingBoxXYZ bbox2XYZ && !bbox2XYZ.Enabled)
                {
                    return false;
                }

                double? distance = CalculateDistance(sleeve1, sleeve2);
                
                if (!distance.HasValue)
                {
                    return false;
                }

                return distance.Value <= tolerance;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[BoundingBoxProximityChecker] Exception in CheckProximity: {ex.Message}, StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Calculate minimum distance between two sleeves using bounding box coordinates.
        /// Handles Floor, Wall, and Structural Framing host types with appropriate 2D/3D calculations.
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

                var bbox1 = sleeve1.BoundingBox;
                var bbox2 = sleeve2.BoundingBox;

                if (bbox1 == null || bbox2 == null)
                {
                    return null;
                }

                string hostType = sleeve1.HostType ?? "Unknown";
                string orientation = sleeve1.Orientation ?? "Unknown";

                double? minDistance;

                if (hostType == "Floor")
                {
                    // Floor sleeves: Calculate 2D distance in X,Y plane (ignore Z)
                    minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                        bbox1.Min.X, bbox1.Min.Y, bbox1.Max.X, bbox1.Max.Y,
                        bbox2.Min.X, bbox2.Min.Y, bbox2.Max.X, bbox2.Max.Y);
                }
                else if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    if (orientation == "X")
                    {
                        // X-oriented walls: Calculate 2D distance in X,Z plane (ignore Y/wall depth)
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.X, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Z,
                            bbox2.Min.X, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Z);
                    }
                    else if (orientation == "Y")
                    {
                        // Y-oriented walls: Calculate 2D distance in Y,Z plane (ignore X/wall depth)
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                            bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                    }
                    else
                    {
                        // Default for walls: Assume Y-oriented, use Y,Z distance
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                            bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                    }
                }
                else
                {
                    // Fallback for unknown host types: Use full 3D distance
                    minDistance = DistanceCalculator.CalculateMinimumDistance3D(
                        bbox1.Min.X, bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Y, bbox1.Max.Z,
                        bbox2.Min.X, bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Y, bbox2.Max.Z);
                }

                // ✅ DIAGNOSTIC LOGGING: Log distance for Duct Accessories to verify clustering
                if (sleeve1.Category?.ToString().IndexOf("Accessories", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double distMM = (minDistance ?? 0) * 304.8;
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[BBoxChecker] 📏 Duct Accessories Proximity: Dist={distMM:F1}mm vs Tol=??mm. Result: {(minDistance ?? 999) <= 0.328}"); // 0.328ft approx 100mm
                }

                return minDistance;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[BoundingBoxProximityChecker] Exception in CalculateDistance: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }
    }
}

