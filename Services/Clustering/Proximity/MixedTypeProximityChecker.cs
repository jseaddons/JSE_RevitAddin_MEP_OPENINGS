using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Proximity checker for mixed sleeve types (e.g. one rectangular, one circular).
    /// Uses RCS (Relative Coordinate System) data for Walls/Framing to avoid WCS coordinate system issues.
    /// Phase 3: Robust path for mixed types in batch clustering.
    /// </summary>
    public class MixedTypeProximityChecker : IProximityChecker
    {
        public bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance)
        {
            try
            {
                double? distance = CalculateDistance(sleeve1, sleeve2);
                if (!distance.HasValue) return false;

                bool result = distance.Value <= tolerance;

                // DIAGNOSTIC LOGGING
                if (result)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[MixedChecker] ✅ MATCH: Mixed proximity detected. Dist={distance.Value * 304.8:F1}mm vs Tol={tolerance * 304.8:F1}mm\n");
                }

                return result;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[MixedTypeProximityChecker] Exception in CheckProximity: {ex.Message}");
                return false;
            }
        }

        public double? CalculateDistance(dynamic sleeve1, dynamic sleeve2)
        {
            try
            {
                ClashZone cz1 = sleeve1.ClashZone as ClashZone;
                ClashZone cz2 = sleeve2.ClashZone as ClashZone;

                if (cz1 == null || cz2 == null) return null;

                string hostType = cz1.StructuralElementType ?? "Unknown";

                // CASE 1: Walls or Structural Framing -> Use RCS for robustness if available
                if (hostType == "Wall" || hostType == "Structural Framing")
                {
                    // Check if BOTH have valid RCS data
                    if (HasValidRcsBoundingBox(cz1) && HasValidRcsBoundingBox(cz2))
                    {
                        // In RCS, X is along host, Z is vertical. We ignore Y (through host).
                        return DistanceCalculator.CalculateMinimumDistance2D(
                            cz1.SleeveBoundingBoxRCS_MinX, cz1.SleeveBoundingBoxRCS_MinZ,
                            cz1.SleeveBoundingBoxRCS_MaxX, cz1.SleeveBoundingBoxRCS_MaxZ,
                            cz2.SleeveBoundingBoxRCS_MinX, cz2.SleeveBoundingBoxRCS_MinZ,
                            cz2.SleeveBoundingBoxRCS_MaxX, cz2.SleeveBoundingBoxRCS_MaxZ);
                    }
                    
                    // Fallback to corners if RCS is missing but corners exist (Corners are more robust than WCS BBox)
                    var cornerHelper = new SleeveCornerProximityHelper();
                    if (cornerHelper.HasValidSleeveCorners(cz1) && cornerHelper.HasValidSleeveCorners(cz2))
                    {
                        var bbox1 = cornerHelper.GetBoundingBoxFromCorners(cz1);
                        var bbox2 = cornerHelper.GetBoundingBoxFromCorners(cz2);

                        // Walls/Framing: Determine orientation and calculate 2D distance
                        string orientation = cz1.HostOrientation ?? "X";
                        if (orientation == "X")
                        {
                            return DistanceCalculator.CalculateMinimumDistance2D(
                                bbox1.minX, bbox1.minZ, bbox1.maxX, bbox1.maxZ,
                                bbox2.minX, bbox2.minZ, bbox2.maxX, bbox2.maxZ);
                        }
                        else
                        {
                            return DistanceCalculator.CalculateMinimumDistance2D(
                                bbox1.minY, bbox1.minZ, bbox1.maxY, bbox1.maxZ,
                                bbox2.minY, bbox2.minZ, bbox2.maxY, bbox2.maxZ);
                        }
                    }
                }

                // CASE 2: Floors or Fallback -> Use WCS Bounding Box (axis-aligned on Floors is safe)
                // Use the BoundingBox property provided by BatchClusterCalculationService wrapper
                var wcsBbox1 = sleeve1.BoundingBox;
                var wcsBbox2 = sleeve2.BoundingBox;

                if (wcsBbox1 != null && wcsBbox2 != null)
                {
                    if (hostType == "Floor")
                    {
                        return DistanceCalculator.CalculateMinimumDistance2D(
                            wcsBbox1.Min.X, wcsBbox1.Min.Y, wcsBbox1.Max.X, wcsBbox1.Max.Y,
                            wcsBbox2.Min.X, wcsBbox2.Min.Y, wcsBbox2.Max.X, wcsBbox2.Max.Y);
                    }
                    
                    // General 3D fallback
                    return DistanceCalculator.CalculateMinimumDistance3D(
                        wcsBbox1.Min.X, wcsBbox1.Min.Y, wcsBbox1.Min.Z, wcsBbox1.Max.X, wcsBbox1.Max.Y, wcsBbox1.Max.Z,
                        wcsBbox2.Min.X, wcsBbox2.Min.Y, wcsBbox2.Min.Z, wcsBbox2.Max.X, wcsBbox2.Max.Y, wcsBbox2.Max.Z);
                }

                return null;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[MixedTypeProximityChecker] Exception in CalculateDistance: {ex.Message}");
                return null;
            }
        }

        private bool HasValidRcsBoundingBox(ClashZone cz)
        {
            // If Max > Min on any axis, it's valid.
            return cz.SleeveBoundingBoxRCS_MaxX > cz.SleeveBoundingBoxRCS_MinX ||
                   cz.SleeveBoundingBoxRCS_MaxZ > cz.SleeveBoundingBoxRCS_MinZ;
        }
    }
}
