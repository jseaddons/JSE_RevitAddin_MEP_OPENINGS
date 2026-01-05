using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
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
        public bool CheckProximity(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, double tolerance)
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
                if (!bbox1.Enabled)
                {
                    return false;
                }
                if (!bbox2.Enabled)
                {
                    return false;
                }

                double? distance = CalculateDistance(sleeve1, sleeve2);
                
                if (!distance.HasValue)
                {
                    return false;
                }

                // ✅ DIAGNOSTIC LOGGING: Log distance for Duct Accessories to verify clustering
                if (sleeve1.Category?.IndexOf("Accessories", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double distMM = distance.Value * 304.8;
                    double toleranceMM = tolerance * 304.8;
                    bool isInRange = distance.Value <= tolerance;
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[BBoxChecker] 📏 Duct Accessories Proximity: Dist={distMM:F1}mm vs Tol={toleranceMM:F1}mm. Result: {isInRange} (MinDistFT={distance.Value:F6}ft)");
                    
                    // ✅ DETAILED DEBUG: Log actual bbox values used
                    string hostType = sleeve1.HostType ?? "Unknown";
                    string orientation = sleeve1.Orientation ?? "Unknown";
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $" [DEBUG] HostType={hostType}, Orientation={orientation}, RevitID1={sleeve1.SleeveInstanceId}, RevitID2={sleeve2.SleeveInstanceId}\n");
                        
                    if (orientation == "X")
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"  BBox1: X=[{bbox1.Min.X:F4},{bbox1.Max.X:F4}], Z=[{bbox1.Min.Z:F4},{bbox1.Max.Z:F4}]\n");
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"  BBox2: X=[{bbox2.Min.X:F4},{bbox2.Max.X:F4}], Z=[{bbox2.Min.Z:F4},{bbox2.Max.Z:F4}]\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"  BBox1: Y=[{bbox1.Min.Y:F4},{bbox1.Max.Y:F4}], Z=[{bbox1.Min.Z:F4},{bbox1.Max.Z:F4}]\n");
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"  BBox2: Y=[{bbox2.Min.Y:F4},{bbox2.Max.Y:F4}], Z=[{bbox2.Min.Z:F4},{bbox2.Max.Z:F4}]\n");
                    }
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
        public double? CalculateDistance(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2)
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
                        // X-oriented walls (Normal=X): Plane is YZ. Calculate 2D distance in Y,Z plane (ignore X/wall depth)
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                            bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                    }
                    else if (orientation == "Y")
                    {
                        // Y-oriented walls (Normal=Y): Plane is XZ. Calculate 2D distance in X,Z plane (ignore Y/wall depth)
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.X, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Z,
                            bbox2.Min.X, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Z);
                    }
                    else
                    {
                        // Default for walls: Assume Y-oriented, use X,Z distance
                        minDistance = DistanceCalculator.CalculateMinimumDistance2D(
                            bbox1.Min.X, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Z,
                            bbox2.Min.X, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Z);
                    }
                }
                else
                {
                    // Fallback for unknown host types: Use full 3D distance
                    minDistance = DistanceCalculator.CalculateMinimumDistance3D(
                        bbox1.Min.X, bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Y, bbox1.Max.Z,
                        bbox2.Min.X, bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Y, bbox2.Max.Z);
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

