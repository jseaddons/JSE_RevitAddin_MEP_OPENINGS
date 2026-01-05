using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Proximity checker using edge-to-edge distance for round pipes/ducts.
    /// Phase 2: Extracted from UniversalClusterService with crash-safe guards.
    /// Formula: edgeToEdgeDistance = centerToCenterDistance - (radius1 + radius2)
    /// </summary>
    public class EdgeToEdgeProximityChecker : IProximityChecker
    {
        /// <summary>
        /// Check if two round sleeves are within proximity tolerance using edge-to-edge distance.
        /// Uses SleeveDiameter from database (ClashZone).
        /// </summary>
        public bool CheckProximity(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, double tolerance)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1 == null || sleeve2 == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[EdgeToEdgeProximityChecker] Null sleeve inputs - returning false");
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
                    $"[EdgeToEdgeProximityChecker] Exception in CheckProximity: {ex.Message}, StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Calculate edge-to-edge distance between two round sleeves.
        /// Returns center-to-center distance minus the sum of radii.
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

                // ✅ DATABASE-ONLY: Get placement points from ClashZone (database data)
                XYZ placement1 = GetPlacementPointFromSleeve(sleeve1);
                XYZ placement2 = GetPlacementPointFromSleeve(sleeve2);

                if (placement1 == null || placement2 == null)
                {
                    return null;
                }

                // ✅ DATABASE-ONLY: Get sleeve radii from ClashZone (SleeveDiameter)
                (double radius1, double radius2) radiiResult = GetSleeveRadiiFromSleeves(sleeve1, sleeve2);
                double radius1 = radiiResult.radius1;
                double radius2 = radiiResult.radius2;

                // ✅ CRASH-SAFE: Validate sleeve diameter > 0 before calculation
                if (radius1 <= 0 || radius2 <= 0)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[EdgeToEdgeProximityChecker] Invalid radii: radius1={radius1}, radius2={radius2}");
                    return null;
                }

                string hostType = sleeve1.HostType ?? "Unknown";
                string orientation = sleeve1.Orientation ?? "Unknown";

                double centerToCenterDistance;

                if (hostType == "Floor")
                {
                    // Floor: 2D distance in X,Y plane (ignore Z)
                    double dx = placement2.X - placement1.X;
                    double dy = placement2.Y - placement1.Y;
                    centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy);
                }
                else if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    if (orientation == "X")
                    {
                        // X-oriented walls: 2D distance in X,Z plane (ignore Y)
                        double dx = placement2.X - placement1.X;
                        double dz = placement2.Z - placement1.Z;
                        centerToCenterDistance = Math.Sqrt(dx * dx + dz * dz);
                    }
                    else
                    {
                        // Y-oriented walls: 2D distance in Y,Z plane (ignore X)
                        double dy = placement2.Y - placement1.Y;
                        double dz = placement2.Z - placement1.Z;
                        centerToCenterDistance = Math.Sqrt(dy * dy + dz * dz);
                    }
                }
                else
                {
                    // Fallback: 3D distance
                    double dx = placement2.X - placement1.X;
                    double dy = placement2.Y - placement1.Y;
                    double dz = placement2.Z - placement1.Z;
                    centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }

                // ✅ CRITICAL: Calculate EDGE-TO-EDGE distance (accounts for sleeve diameter)
                // Formula: edgeToEdge = centerToCenter - (radius1 + radius2)
                // This is the actual gap between the two sleeves
                double edgeToEdgeDistance = centerToCenterDistance - (radius1 + radius2);

                // ✅ DIAGNOSTIC LOGGING: Log distance for pipes/ducts to verify clustering
                string category = sleeve1.Category;
                
                if (!string.IsNullOrEmpty(category) && 
                   (category.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    category.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    double distMM = edgeToEdgeDistance * 304.8;
                    double centerDistMM = centerToCenterDistance * 304.8;
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[EdgeToEdge] 📏 {category} Proximity: Dist={distMM:F1}mm (CenterDist={centerDistMM:F1}mm) vs Tol=??mm. Result: {edgeToEdgeDistance <= 0.328}"); // 0.328ft approx 100mm
                }

                // ✅ Ensure non-negative (sleeves can overlap)
                return Math.Max(0, edgeToEdgeDistance);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[EdgeToEdgeProximityChecker] Exception in CalculateDistance: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// ✅ DATABASE-ONLY: Helper method to get placement point from dynamic sleeve object.
        /// Uses ClashZone data from database (SleevePlacementPointActiveDocumentX/Y/Z).
        /// </summary>
        private XYZ GetPlacementPointFromSleeve(ClusteringSleeveDto sleeve)
        {
            try
            {
                // ✅ CRASH-SAFE: Check for null
                if (sleeve?.ClashZone == null)
                {
                    return null;
                }

                var cz = sleeve.ClashZone;
                if (cz == null)
                {
                    return null;
                }

                // ✅ Use active document placement point (where sleeve is actually placed)
                if (cz.SleevePlacementPointActiveDocumentX != 0 ||
                    cz.SleevePlacementPointActiveDocumentY != 0 ||
                    cz.SleevePlacementPointActiveDocumentZ != 0)
                {
                    return new XYZ(
                        cz.SleevePlacementPointActiveDocumentX,
                        cz.SleevePlacementPointActiveDocumentY,
                        cz.SleevePlacementPointActiveDocumentZ
                    );
                }

                // Fallback to regular placement point (from database)
                if (cz.SleevePlacementPointX != 0 ||
                    cz.SleevePlacementPointY != 0 ||
                    cz.SleevePlacementPointZ != 0)
                {
                    return new XYZ(
                        cz.SleevePlacementPointX,
                        cz.SleevePlacementPointY,
                        cz.SleevePlacementPointZ
                    );
                }

                // Fallback: Calculate center from bounding box (from database)
                var bbox = sleeve.BoundingBox;
                if (bbox != null)
                {
                    return new XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0
                    );
                }

                return null;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[EdgeToEdgeProximityChecker] Exception in GetPlacementPointFromSleeve: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// ✅ DATABASE-ONLY: Get sleeve radii from ClashZone (SleeveDiameter).
        /// NOTE: Rectangular sleeves should use bounding box logic, not this method.
        /// </summary>
        private (double radius1, double radius2) GetSleeveRadiiFromSleeves(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2)
        {
            double radius1 = 0;
            double radius2 = 0;

            try
            {
                // ✅ CRASH-SAFE: Check for null
                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    return (0, 0);
                }

                var cz1 = sleeve1.ClashZone;
                var cz2 = sleeve2.ClashZone;

                if (cz1 == null || cz2 == null)
                {
                    return (0, 0);
                }

                // ✅ DATABASE-ONLY: Get SleeveDiameter from ClashZone (stored in database)
                // SleeveDiameter is in internal units (feet)
                double? diameter1 = cz1.SleeveDiameter;
                double? diameter2 = cz2.SleeveDiameter;

                if (diameter1.HasValue && diameter1.Value > 0)
                {
                    radius1 = diameter1.Value / 2.0;
                }

                if (diameter2.HasValue && diameter2.Value > 0)
                {
                    radius2 = diameter2.Value / 2.0;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[EdgeToEdgeProximityChecker] Exception in GetSleeveRadiiFromSleeves: {ex.Message}");
            }

            return (radius1, radius2);
        }
    }
}

