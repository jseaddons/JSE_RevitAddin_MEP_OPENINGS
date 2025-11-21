using System;
using System.Collections.Concurrent;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// Transforms bounding boxes between World Coordinate System (WCS) and Wall-aligned Relative Coordinate System (RCS).
    /// 
    /// RCS Definition for Walls/Framing:
    /// - RCS X-axis = Along wall direction (WallDirection vector)
    /// - RCS Y-axis = Through wall (perpendicular to wall direction in XY plane)
    /// - RCS Z-axis = Vertical (same as WCS Z)
    /// 
    /// Features:
    /// - ✅ Crash-safe: Validates all inputs and outputs
    /// - ✅ Optimized: Caches transformation matrices (dump once, use many times)
    /// - ✅ Thread-safe: Uses ConcurrentDictionary for multi-threaded access
    /// - ✅ Performance: Batch transformations for multiple points
    /// </summary>
    public static class WallRcsTransformer
    {
        // ✅ OPTIMIZATION: Cache transformation matrices by wall direction (dump once, use many times)
        // Key: WallDirection hash (normalized vector), Value: Cached transformation basis vectors
        private static readonly ConcurrentDictionary<string, RcsBasisVectors> _transformationCache = new();
        private static readonly object _cacheLock = new object();
        private const int MAX_CACHE_SIZE = 1000; // Prevent unbounded cache growth

        /// <summary>
        /// Cached RCS basis vectors for a wall direction
        /// </summary>
        internal class RcsBasisVectors
        {
            public XYZ RcsX { get; set; } // Along wall
            public XYZ RcsY { get; set; } // Through wall
            public XYZ RcsZ { get; set; } // Vertical
            public DateTime LastUsed { get; set; } = DateTime.Now;
        }

        /// <summary>
        /// Transform WCS bounding box to wall-aligned RCS.
        /// 
        /// ✅ CRASH-SAFE: Validates all inputs and handles edge cases
        /// ✅ OPTIMIZED: Uses cached transformation matrices
        /// </summary>
        /// <param name="wcsBbox">Bounding box in World Coordinate System</param>
        /// <param name="wallDirection">Wall direction vector (normalized or not)</param>
        /// <returns>Bounding box in RCS, or null if transformation fails</returns>
        public static BoundingBoxXYZ? TransformToRcs(BoundingBoxXYZ? wcsBbox, XYZ? wallDirection)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (wcsBbox == null || !wcsBbox.Enabled)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid WCS bounding box: null or disabled");
                    }
                    return null;
                }

                if (wallDirection == null || wallDirection.IsZeroLength())
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid wall direction: null or zero length");
                    }
                    return null;
                }

                // ✅ CRASH-SAFE: Validate bounding box coordinates
                if (!IsValidBoundingBox(wcsBbox))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid WCS bounding box coordinates: Min=({wcsBbox.Min.X}, {wcsBbox.Min.Y}, {wcsBbox.Min.Z}), Max=({wcsBbox.Max.X}, {wcsBbox.Max.Y}, {wcsBbox.Max.Z})");
                    }
                    return null;
                }

                // ✅ OPTIMIZATION: Get cached transformation basis vectors
                var basis = GetOrCreateBasisVectors(wallDirection);
                if (basis == null)
                {
                    return null;
                }

                // ✅ PERFORMANCE: Transform all 8 corners of bounding box in batch
                var wcsCorners = GetBoundingBoxCorners(wcsBbox);
                var rcsCorners = TransformPointsToRcs(wcsCorners, basis);

                if (rcsCorners == null || rcsCorners.Length == 0)
                {
                    return null;
                }

                // Calculate min/max in RCS
                var rcsX_vals = rcsCorners.Select(c => c.X).ToArray();
                var rcsY_vals = rcsCorners.Select(c => c.Y).ToArray();
                var rcsZ_vals = rcsCorners.Select(c => c.Z).ToArray();

                var rcsBbox = new BoundingBoxXYZ
                {
                    Min = new XYZ(rcsX_vals.Min(), rcsY_vals.Min(), rcsZ_vals.Min()),
                    Max = new XYZ(rcsX_vals.Max(), rcsY_vals.Max(), rcsZ_vals.Max()),
                    Enabled = true
                };

                // ✅ CRASH-SAFE: Validate result
                if (!IsValidBoundingBox(rcsBbox))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid RCS bounding box result: Min=({rcsBbox.Min.X}, {rcsBbox.Min.Y}, {rcsBbox.Min.Z}), Max=({rcsBbox.Max.X}, {rcsBbox.Max.Y}, {rcsBbox.Max.Z})");
                    }
                    return null;
                }

                return rcsBbox;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[WallRcsTransformer] Exception in TransformToRcs: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Transform RCS point back to WCS.
        /// 
        /// ✅ CRASH-SAFE: Validates all inputs
        /// ✅ OPTIMIZED: Uses cached transformation matrices
        /// </summary>
        /// <param name="rcsPoint">Point in RCS</param>
        /// <param name="wallDirection">Wall direction vector</param>
        /// <param name="origin">Origin point in WCS (typically first sleeve center)</param>
        /// <returns>Point in WCS, or null if transformation fails</returns>
        public static XYZ? TransformToWcs(XYZ? rcsPoint, XYZ? wallDirection, XYZ? origin)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (rcsPoint == null || origin == null || wallDirection == null || wallDirection.IsZeroLength())
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid inputs in TransformToWcs: rcsPoint={rcsPoint}, origin={origin}, wallDirection={wallDirection}");
                    }
                    return null;
                }

                if (!IsValidCoordinate(rcsPoint.X) || !IsValidCoordinate(rcsPoint.Y) || !IsValidCoordinate(rcsPoint.Z) ||
                    !IsValidCoordinate(origin.X) || !IsValidCoordinate(origin.Y) || !IsValidCoordinate(origin.Z))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid coordinates in TransformToWcs: rcsPoint=({rcsPoint.X}, {rcsPoint.Y}, {rcsPoint.Z}), origin=({origin.X}, {origin.Y}, {origin.Z})");
                    }
                    return null;
                }

                // ✅ OPTIMIZATION: Get cached transformation basis vectors
                var basis = GetOrCreateBasisVectors(wallDirection);
                if (basis == null)
                {
                    return null;
                }

                // Inverse transformation: RCS → WCS
                // wcsPoint = origin + rcsX * RcsX + rcsY * RcsY + rcsZ * RcsZ
                var wcsPoint = origin + rcsPoint.X * basis.RcsX + rcsPoint.Y * basis.RcsY + rcsPoint.Z * basis.RcsZ;

                // ✅ CRASH-SAFE: Validate result
                if (!IsValidCoordinate(wcsPoint.X) || !IsValidCoordinate(wcsPoint.Y) || !IsValidCoordinate(wcsPoint.Z))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("geometry_errors.log",
                            $"[WallRcsTransformer] Invalid WCS point result: ({wcsPoint.X}, {wcsPoint.Y}, {wcsPoint.Z})");
                    }
                    return null;
                }

                return wcsPoint;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[WallRcsTransformer] Exception in TransformToWcs: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Batch transform multiple points from WCS to RCS.
        /// 
        /// ✅ PERFORMANCE: Optimized for batch operations
        /// ✅ MULTI-THREADING: Thread-safe transformation
        /// </summary>
        /// <param name="wcsPoints">Points in WCS</param>
        /// <param name="wallDirection">Wall direction vector</param>
        /// <param name="origin">Origin point in WCS (for relative transformation)</param>
        /// <returns>Transformed points in RCS, or null if transformation fails</returns>
        public static XYZ[]? TransformPointsToRcs(XYZ[]? wcsPoints, XYZ? wallDirection, XYZ? origin = null)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (wcsPoints == null || wcsPoints.Length == 0 || wallDirection == null || wallDirection.IsZeroLength())
                {
                    return null;
                }

                // ✅ OPTIMIZATION: Get cached transformation basis vectors
                var basis = GetOrCreateBasisVectors(wallDirection);
                if (basis == null)
                {
                    return null;
                }

                // ✅ PERFORMANCE: Use origin if provided, otherwise use first point as origin
                XYZ transformOrigin = origin ?? wcsPoints[0];

                // ✅ MULTI-THREADING: Parallel transformation for large batches
                if (wcsPoints.Length > 100)
                {
                    var rcsPoints = new XYZ[wcsPoints.Length];
                    System.Threading.Tasks.Parallel.For(0, wcsPoints.Length, i =>
                    {
                        var wcsPoint = wcsPoints[i];
                        if (wcsPoint != null && IsValidCoordinate(wcsPoint.X) && IsValidCoordinate(wcsPoint.Y) && IsValidCoordinate(wcsPoint.Z))
                        {
                            // Translate relative to origin
                            var relPoint = wcsPoint - transformOrigin;
                            
                            // Transform to RCS using dot product
                            var rcsX = relPoint.DotProduct(basis.RcsX);
                            var rcsY = relPoint.DotProduct(basis.RcsY);
                            var rcsZ = relPoint.DotProduct(basis.RcsZ);
                            
                            if (IsValidCoordinate(rcsX) && IsValidCoordinate(rcsY) && IsValidCoordinate(rcsZ))
                            {
                                rcsPoints[i] = new XYZ(rcsX, rcsY, rcsZ);
                            }
                        }
                    });
                    return rcsPoints;
                }
                else
                {
                    // Sequential for small batches (overhead of parallelization not worth it)
                    var rcsPoints = new XYZ[wcsPoints.Length];
                    for (int i = 0; i < wcsPoints.Length; i++)
                    {
                        var wcsPoint = wcsPoints[i];
                        if (wcsPoint != null && IsValidCoordinate(wcsPoint.X) && IsValidCoordinate(wcsPoint.Y) && IsValidCoordinate(wcsPoint.Z))
                        {
                            var relPoint = wcsPoint - transformOrigin;
                            var rcsX = relPoint.DotProduct(basis.RcsX);
                            var rcsY = relPoint.DotProduct(basis.RcsY);
                            var rcsZ = relPoint.DotProduct(basis.RcsZ);
                            
                            if (IsValidCoordinate(rcsX) && IsValidCoordinate(rcsY) && IsValidCoordinate(rcsZ))
                            {
                                rcsPoints[i] = new XYZ(rcsX, rcsY, rcsZ);
                            }
                        }
                    }
                    return rcsPoints;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[WallRcsTransformer] Exception in TransformPointsToRcs: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Overload: Transform points using pre-calculated basis vectors (for performance).
        /// </summary>
        internal static XYZ[]? TransformPointsToRcs(XYZ[]? wcsPoints, RcsBasisVectors basis)
        {
            if (wcsPoints == null || wcsPoints.Length == 0 || basis == null)
            {
                return null;
            }

            var rcsPoints = new XYZ[wcsPoints.Length];
            for (int i = 0; i < wcsPoints.Length; i++)
            {
                var wcsPoint = wcsPoints[i];
                if (wcsPoint != null && IsValidCoordinate(wcsPoint.X) && IsValidCoordinate(wcsPoint.Y) && IsValidCoordinate(wcsPoint.Z))
                {
                    var rcsX = wcsPoint.DotProduct(basis.RcsX);
                    var rcsY = wcsPoint.DotProduct(basis.RcsY);
                    var rcsZ = wcsPoint.DotProduct(basis.RcsZ);
                    
                    if (IsValidCoordinate(rcsX) && IsValidCoordinate(rcsY) && IsValidCoordinate(rcsZ))
                    {
                        rcsPoints[i] = new XYZ(rcsX, rcsY, rcsZ);
                    }
                }
            }
            return rcsPoints;
        }

        /// <summary>
        /// Get or create cached RCS basis vectors for a wall direction.
        /// 
        /// ✅ OPTIMIZATION: Caches transformation matrices (dump once, use many times)
        /// ✅ THREAD-SAFE: Uses ConcurrentDictionary
        /// </summary>
        private static RcsBasisVectors? GetOrCreateBasisVectors(XYZ wallDirection)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate and normalize wall direction
                if (wallDirection == null || wallDirection.IsZeroLength())
                {
                    return null;
                }

                var normalized = wallDirection.Normalize();
                if (normalized == null || normalized.IsZeroLength())
                {
                    return null;
                }

                // ✅ OPTIMIZATION: Create cache key from normalized direction (rounded to avoid floating point issues)
                // Round to 6 decimal places to handle floating point precision
                string cacheKey = $"{Math.Round(normalized.X, 6)},{Math.Round(normalized.Y, 6)},{Math.Round(normalized.Z, 6)}";

                // ✅ THREAD-SAFE: Get or create cached basis vectors
                return _transformationCache.GetOrAdd(cacheKey, key =>
                {
                    // ✅ PERFORMANCE: Limit cache size to prevent memory issues
                    lock (_cacheLock)
                    {
                        if (_transformationCache.Count >= MAX_CACHE_SIZE)
                        {
                            // Remove least recently used entry (simple LRU)
                            var oldest = _transformationCache.OrderBy(kvp => kvp.Value.LastUsed).First();
                            _transformationCache.TryRemove(oldest.Key, out _);
                        }
                    }

                    // Calculate RCS basis vectors
                    // RCS X = along wall (wall direction)
                    XYZ rcsX = normalized;
                    
                    // RCS Y = through wall (perpendicular to wall direction in XY plane)
                    // Rotate wall direction 90° clockwise in XY plane: (-Y, X, 0)
                    XYZ rcsY = new XYZ(-normalized.Y, normalized.X, 0).Normalize();
                    if (rcsY == null || rcsY.IsZeroLength())
                    {
                        // Fallback: If wall is vertical, use Y-axis
                        rcsY = XYZ.BasisY;
                    }
                    
                    // RCS Z = vertical (same as WCS Z)
                    XYZ rcsZ = XYZ.BasisZ;

                    return new RcsBasisVectors
                    {
                        RcsX = rcsX,
                        RcsY = rcsY,
                        RcsZ = rcsZ,
                        LastUsed = DateTime.Now
                    };
                });
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[WallRcsTransformer] Exception in GetOrCreateBasisVectors: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// Get all 8 corners of a bounding box.
        /// </summary>
        private static XYZ[] GetBoundingBoxCorners(BoundingBoxXYZ bbox)
        {
            return new XYZ[]
            {
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };
        }

        /// <summary>
        /// Validate that a coordinate is a valid number.
        /// </summary>
        private static bool IsValidCoordinate(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        /// <summary>
        /// Validate that a bounding box has valid coordinates.
        /// </summary>
        private static bool IsValidBoundingBox(BoundingBoxXYZ bbox)
        {
            if (bbox == null || !bbox.Enabled)
                return false;

            return IsValidCoordinate(bbox.Min.X) && IsValidCoordinate(bbox.Min.Y) && IsValidCoordinate(bbox.Min.Z) &&
                   IsValidCoordinate(bbox.Max.X) && IsValidCoordinate(bbox.Max.Y) && IsValidCoordinate(bbox.Max.Z) &&
                   bbox.Max.X >= bbox.Min.X && bbox.Max.Y >= bbox.Min.Y && bbox.Max.Z >= bbox.Min.Z;
        }

        /// <summary>
        /// Clear transformation cache (for testing or memory management).
        /// </summary>
        public static void ClearCache()
        {
            _transformationCache.Clear();
        }

        /// <summary>
        /// Get cache statistics (for diagnostics).
        /// </summary>
        public static (int count, long memoryBytes) GetCacheStats()
        {
            return (_transformationCache.Count, System.GC.GetTotalMemory(false));
        }
    }
}

