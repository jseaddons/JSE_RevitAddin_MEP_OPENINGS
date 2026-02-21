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
                if (OptimizationFlags.EnableProximityDebugLog || result)
                {
                    string id1 = "unknown";
                    string id2 = "unknown";
                    try 
                    {
                        var cz1 = sleeve1 as ClashZone;
                        var cz2 = sleeve2 as ClashZone;
                        if (cz1 != null) id1 = cz1.SleeveInstanceId.ToString();
                        else id1 = ((dynamic)sleeve1).SleeveId?.ToString() ?? "dyn";
                        
                        if (cz2 != null) id2 = cz2.SleeveInstanceId.ToString();
                        else id2 = ((dynamic)sleeve2).SleeveId?.ToString() ?? "dyn";
                    }
                    catch { }

                    string status = result ? "✅ MATCH" : "❌ FAIL";
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[MixedChecker] {status}: IDs={id1}/{id2}, Dist={distance.Value * 304.8:F1}mm vs Tol={tolerance * 304.8:F1}mm\n");
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
                if (hostType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0 || hostType.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0)
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
                }

                // CASE 2: Mixed type handling (rectangular vs circular) using appropriate geometry
                var cornerHelper = new SleeveCornerProximityHelper();
                
                // Get bounding box for each sleeve (handles both rectangular and circular)
                var bbox1 = GetBoundingBoxForMixedType(cz1, cornerHelper);
                var bbox2 = GetBoundingBoxForMixedType(cz2, cornerHelper);
                
                if (!bbox1.HasValue || !bbox2.HasValue)
                {
                    return null;
                }
                
                var box1 = bbox1.Value;
                var box2 = bbox2.Value;

                if (hostType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Floor: 2D distance in X,Y plane (ignore Z)
                    double? dist = DistanceCalculator.CalculateMinimumDistance2D(
                        box1.minX, box1.minY, box1.maxX, box1.maxY,
                        box2.minX, box2.minY, box2.maxX, box2.maxY);
                    if (!dist.HasValue) return null;
                    SafeFileLogger.SafeAppendText("mixed_type_distances.log",
                        $"[{DateTime.Now:HH:mm:ss}] FLOOR: {cz1.ClashZoneGuid?.Substring(0,8)} vs {cz2.ClashZoneGuid?.Substring(0,8)}, Dist={dist.Value*304.8:F1}mm, BBox1=({box1.minX:F2},{box1.minY:F2})-({box1.maxX:F2},{box1.maxY:F2}), BBox2=({box2.minX:F2},{box2.minY:F2})-({box2.maxX:F2},{box2.maxY:F2})\n");
                    return dist.Value;
                }
                else if (hostType.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0 || hostType.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string orientation = cz1.HostOrientation ?? "X";
                    if (orientation == "X")
                    {
                        // X-wall: 2D distance in X,Z plane (ignore Y - through wall)
                        double? dist = DistanceCalculator.CalculateMinimumDistance2D(
                            box1.minX, box1.minZ, box1.maxX, box1.maxZ,
                            box2.minX, box2.minZ, box2.maxX, box2.maxZ);
                        if (!dist.HasValue) return null;
                        SafeFileLogger.SafeAppendText("mixed_type_distances.log",
                            $"[{DateTime.Now:HH:mm:ss}] X-WALL: {cz1.ClashZoneGuid?.Substring(0,8)} vs {cz2.ClashZoneGuid?.Substring(0,8)}, Dist={dist.Value*304.8:F1}mm, BBox1_XZ=({box1.minX:F2},{box1.minZ:F2})-({box1.maxX:F2},{box1.maxZ:F2}), BBox2_XZ=({box2.minX:F2},{box2.minZ:F2})-({box2.maxX:F2},{box2.maxZ:F2})\n");
                        return dist.Value;
                    }
                    else
                    {
                        // Y-wall: 2D distance in Y,Z plane (ignore X - through wall)
                        double? dist = DistanceCalculator.CalculateMinimumDistance2D(
                            box1.minY, box1.minZ, box1.maxY, box1.maxZ,
                            box2.minY, box2.minZ, box2.maxY, box2.maxZ);
                        if (!dist.HasValue) return null;
                        SafeFileLogger.SafeAppendText("mixed_type_distances.log",
                            $"[{DateTime.Now:HH:mm:ss}] Y-WALL: {cz1.ClashZoneGuid?.Substring(0,8)} vs {cz2.ClashZoneGuid?.Substring(0,8)}, Dist={dist.Value*304.8:F1}mm, BBox1_YZ=({box1.minY:F2},{box1.minZ:F2})-({box1.maxY:F2},{box1.maxZ:F2}), BBox2_YZ=({box2.minY:F2},{box2.minZ:F2})-({box2.maxY:F2},{box2.maxZ:F2})\n");
                        return dist.Value;
                    }
                }
                else 
                {
                    // 3D fallback
                    return DistanceCalculator.CalculateMinimumDistance3D(
                        box1.minX, box1.minY, box1.minZ, box1.maxX, box1.maxY, box1.maxZ,
                        box2.minX, box2.minY, box2.minZ, box2.maxX, box2.maxY, box2.maxZ);
                }
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
        
        /// <summary>
        /// Gets bounding box for mixed type proximity checking.
        /// - Rectangular sleeves: use corners from database
        /// - Circular sleeves: use WCS bounding box (SleeveBoundingBoxMinX/Y/Z, MaxX/Y/Z)
        /// </summary>
        private (double minX, double maxX, double minY, double maxY, double minZ, double maxZ)? 
            GetBoundingBoxForMixedType(ClashZone cz, SleeveCornerProximityHelper cornerHelper)
        {
            if (cz == null) return null;
            
            // Case 1: Rectangular sleeve - has corners
            if (cornerHelper.HasValidSleeveCorners(cz))
            {
                try
                {
                    return cornerHelper.GetBoundingBoxFromCorners(cz);
                }
                catch
                {
                    // Fall through to WCS bounding box
                }
            }
            
            // Case 2: Circular sleeve - use WCS bounding box from database
            if (HasValidWcsBoundingBox(cz))
            {
                return (
                    cz.BoundingBoxMinX, cz.BoundingBoxMaxX,
                    cz.BoundingBoxMinY, cz.BoundingBoxMaxY,
                    cz.BoundingBoxMinZ, cz.BoundingBoxMaxZ
                );
            }
            
            return null;
        }
        
        /// <summary>
        /// Checks if ClashZone has valid WCS bounding box data.
        /// </summary>
        private bool HasValidWcsBoundingBox(ClashZone cz)
        {
            if (cz == null) return false;
            
            // Valid if Max > Min on at least one axis
            return cz.BoundingBoxMaxX > cz.BoundingBoxMinX ||
                   cz.BoundingBoxMaxY > cz.BoundingBoxMinY ||
                   cz.BoundingBoxMaxZ > cz.BoundingBoxMinZ;
        }
    }
}
