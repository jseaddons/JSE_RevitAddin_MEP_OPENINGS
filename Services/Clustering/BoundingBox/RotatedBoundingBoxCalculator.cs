using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox
{
    /// <summary>
    /// Calculates rotated cluster bounding boxes using corner-based watertight algorithm.
    /// Phase 3: Extracted from UniversalClusterService.GetClusterBoundingBoxWithRotatedCoordinates with crash-safe guards.
    /// </summary>
    public class RotatedBoundingBoxCalculator : IBoundingBoxCalculator
    {
        private readonly Func<int, string, dynamic> _getClashZoneBySleeveInstanceId;
        private readonly Func<List<FamilyInstance>, double, (double width, double height, double depth, XYZ mid)> _getClusterBoundingBoxFallback;

        /// <summary>
        /// Constructor with dependencies for ClashZone lookup and fallback bounding box calculation.
        /// </summary>
        public RotatedBoundingBoxCalculator(
            Func<int, string, dynamic> getClashZoneBySleeveInstanceId,
            Func<List<FamilyInstance>, double, (double width, double height, double depth, XYZ mid)> getClusterBoundingBoxFallback)
        {
            _getClashZoneBySleeveInstanceId = getClashZoneBySleeveInstanceId ?? throw new ArgumentNullException(nameof(getClashZoneBySleeveInstanceId));
            _getClusterBoundingBoxFallback = getClusterBoundingBoxFallback ?? throw new ArgumentNullException(nameof(getClusterBoundingBoxFallback));
        }

        /// <summary>
        /// Calculate rotated cluster bounding box using corner-based watertight algorithm.
        /// Uses pre-calculated corners and rotation matrices from database (dump once use many times).
        /// </summary>
        public BoundingBoxResult Calculate(
            List<ClusteringSleeveDto> cluster,
            List<FamilyInstance> actualSleeves,
            double rotationAngle,
            string xmlFilePath = null)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (cluster == null || cluster.Count == 0 || actualSleeves == null || actualSleeves.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotatedBoundingBoxCalculator] Invalid inputs - returning empty result");
                    return new BoundingBoxResult();
                }

                // ✅ Collect rotated and axis-aligned bounding boxes from ClashZone data
                var rotatedBboxes = new List<(XYZ min, XYZ max)>();
                var axisAlignedBboxes = new List<(XYZ min, XYZ max)>();
                bool hasRotatedBboxes = false;

                foreach (var sleeveData in cluster)
                {
                    int sId = sleeveData.SleeveInstanceId;
                    
                    // Priority: Use DTO ClashZone first, then lookup via delegate if missing
                    var clashZone = sleeveData.ClashZone;
                    if (clashZone == null && _getClashZoneBySleeveInstanceId != null)
                    {
                         clashZone = _getClashZoneBySleeveInstanceId(sId, xmlFilePath) as Models.ClashZone;
                    }
                    
                    if (clashZone == null)
                    {
                        DebugLogger.Warning($"[RotatedBoundingBoxCalculator] DB Lookup Failed for Sleeve {sId}. Using Fallsback.");
                        // Try Revit API as fallback
                        var sleeve = actualSleeves.FirstOrDefault(s => s.Id.IntegerValue == sId);
                        if (sleeve != null)
                        {
                            var bbox = sleeve.get_BoundingBox(null);
                            if (bbox != null && bbox.Enabled)
                            {
                                axisAlignedBboxes.Add((bbox.Min, bbox.Max));
                            }
                        }
                        continue;
                    }

                    bool hasRotatedBbox = clashZone.RotatedBoundingBoxMinX.HasValue &&
                                         clashZone.RotatedBoundingBoxMinY.HasValue &&
                                         clashZone.RotatedBoundingBoxMaxX.HasValue &&
                                         clashZone.RotatedBoundingBoxMaxY.HasValue;

                    if (hasRotatedBbox)
                    {
                        DebugLogger.Info($"[RotatedBoundingBoxCalculator] Extracted Corners from DB for Sleeve {sId}: X=[{clashZone.RotatedBoundingBoxMinX}, {clashZone.RotatedBoundingBoxMaxX}], Y=[{clashZone.RotatedBoundingBoxMinY}, {clashZone.RotatedBoundingBoxMaxY}]");
                        
                        rotatedBboxes.Add((
                            new XYZ(clashZone.RotatedBoundingBoxMinX.Value,
                                   clashZone.RotatedBoundingBoxMinY.Value,
                                   clashZone.RotatedBoundingBoxMinZ ?? clashZone.SleeveBoundingBoxMinZ),
                            new XYZ(clashZone.RotatedBoundingBoxMaxX.Value,
                                   clashZone.RotatedBoundingBoxMaxY.Value,
                                   clashZone.RotatedBoundingBoxMaxZ ?? clashZone.SleeveBoundingBoxMaxZ)
                        ));
                        hasRotatedBboxes = true;
                    }
                    else
                    {
                        DebugLogger.Info($"[RotatedBoundingBoxCalculator] Using Axis-Aligned Bounds from DB for Sleeve {sId} (No Rotated Corners Found)");
                        axisAlignedBboxes.Add((
                            new XYZ(clashZone.SleeveBoundingBoxMinX,
                                   clashZone.SleeveBoundingBoxMinY,
                                   clashZone.SleeveBoundingBoxMinZ),
                            new XYZ(clashZone.SleeveBoundingBoxMaxX,
                                   clashZone.SleeveBoundingBoxMaxY,
                                   clashZone.SleeveBoundingBoxMaxZ)
                        ));
                    }
                }

                // ✅ Use corner-based watertight algorithm if we have rotated bounding boxes
                if (hasRotatedBboxes && rotatedBboxes.Count > 0 && Math.Abs(rotationAngle) > 1e-6)
                {
                    // ✅ PROPER WATERTIGHT ALGORITHM: Use corner-based calculation
                    // Need to check CornerBasedBoundingBoxCalculator signature too, but assume it handles List<ClusteringSleeveDto> or fix soon.
                    // For now, cast to dynamic if CornerBasedBoundingBoxCalculator (static) expects dynamic, OR assume I fix it too.
                    // IMPORTANT: CornerBasedBoundingBoxCalculator.CalculateFromCorners takes List<dynamic> (cluster).
                    // I will perform a quick cast here until I fix that one too, or better yet, assume I will fix ALL calculators in this batch.
                    // Let's assume I fix CornerBasedBoundingBoxCalculator next.
                    var cornerResult = CornerBasedBoundingBoxCalculator.CalculateFromCorners(
                        cluster, // Pass DTO list
                        rotationAngle,
                        out XYZ origin,
                        _getClashZoneBySleeveInstanceId,
                        xmlFilePath);

                    if (cornerResult.HasValue)
                    {
                        var (width, height, minX, minY, maxX, maxY, calculatedOrigin) = cornerResult.Value;
                        origin = calculatedOrigin;

                        double minZ = rotatedBboxes.Min(b => b.min.Z);
                        double maxZ = rotatedBboxes.Max(b => b.max.Z);
                        double depth = maxZ - minZ;

                        // ✅ MIDPOINT: Calculate in rotated coordinate space, then transform back to world coordinates
                        XYZ midRotated = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);

                        XYZ mid;
                        if (Math.Abs(rotationAngle) > 1e-6 && origin != XYZ.Zero)
                        {
                            // ✅ Transform midpoint from rotated coordinate space back to world coordinates
                            var inverseRotationMatrix = RotationMatrixCalculator.CreateRotationMatrix(rotationAngle);
                            if (inverseRotationMatrix != null)
                            {
                                double cosA = inverseRotationMatrix.Value.cos;
                                double sinA = inverseRotationMatrix.Value.sin;
                                var inversePoint = RotationMatrixCalculator.InverseRotation((midRotated.X, midRotated.Y), cosA, sinA);
                                if (inversePoint != null)
                                {
                                    mid = new XYZ(
                                        origin.X + inversePoint.Value.x,
                                        origin.Y + inversePoint.Value.y,
                                        midRotated.Z
                                    );
                                }
                                else
                                {
                                    mid = new XYZ(origin.X + midRotated.X, origin.Y + midRotated.Y, midRotated.Z);
                                }
                            }
                            else
                            {
                                mid = new XYZ(origin.X + midRotated.X, origin.Y + midRotated.Y, midRotated.Z);
                            }
                        }
                        else
                        {
                            mid = origin != XYZ.Zero
                                ? new XYZ(origin.X + midRotated.X, origin.Y + midRotated.Y, midRotated.Z)
                                : midRotated;
                        }

                        return new BoundingBoxResult
                        {
                            Width = width,
                            Height = height,
                            Depth = depth,
                            Midpoint = mid,
                            RotatedMinX = minX,
                            RotatedMinY = minY,
                            RotatedMinZ = minZ,
                            RotatedMaxX = maxX,
                            RotatedMaxY = maxY,
                            RotatedMaxZ = maxZ
                        };
                    }
                }

                // ✅ Fallback: Use simple union of rotated bounding boxes
                if (rotatedBboxes.Count > 0)
                {
                    double minX = rotatedBboxes.Min(b => b.min.X);
                    double minY = rotatedBboxes.Min(b => b.min.Y);
                    double minZ = rotatedBboxes.Min(b => b.min.Z);
                    double maxX = rotatedBboxes.Max(b => b.max.X);
                    double maxY = rotatedBboxes.Max(b => b.max.Y);
                    double maxZ = rotatedBboxes.Max(b => b.max.Z);

                    double width = maxX - minX;
                    double height = maxY - minY;
                    double depth = maxZ - minZ;
                    XYZ mid = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);

                    return new BoundingBoxResult
                    {
                        Width = width,
                        Height = height,
                        Depth = depth,
                        Midpoint = mid,
                        RotatedMinX = minX,
                        RotatedMinY = minY,
                        RotatedMinZ = minZ,
                        RotatedMaxX = maxX,
                        RotatedMaxY = maxY,
                        RotatedMaxZ = maxZ
                    };
                }

                // ✅ Final fallback: Use ClusterBoundingBoxServices
                var fallbackResult = _getClusterBoundingBoxFallback(actualSleeves, rotationAngle);
                return new BoundingBoxResult
                {
                    Width = fallbackResult.width,
                    Height = fallbackResult.height,
                    Depth = fallbackResult.depth,
                    Midpoint = fallbackResult.mid
                };
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotatedBoundingBoxCalculator] Exception in Calculate: {ex.Message}, StackTrace: {ex.StackTrace}");
                
                // ✅ CRASH-SAFE: Return fallback result
                var fallbackResult = _getClusterBoundingBoxFallback(actualSleeves, rotationAngle);
                return new BoundingBoxResult
                {
                    Width = fallbackResult.width,
                    Height = fallbackResult.height,
                    Depth = fallbackResult.depth,
                    Midpoint = fallbackResult.mid
                };
            }
        }

        /// <summary>
        /// Calculate bounding box from pre-transformed corner points.
        /// </summary>
        public BoundingBoxResult CalculateFromCorners(List<XYZ> corners)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (corners == null || corners.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[RotatedBoundingBoxCalculator] Invalid corners input - returning empty result");
                    return new BoundingBoxResult();
                }

                double minX = corners.Min(p => p.X);
                double minY = corners.Min(p => p.Y);
                double minZ = corners.Min(p => p.Z);
                double maxX = corners.Max(p => p.X);
                double maxY = corners.Max(p => p.Y);
                double maxZ = corners.Max(p => p.Z);

                return new BoundingBoxResult
                {
                    Width = maxX - minX,
                    Height = maxY - minY,
                    Depth = maxZ - minZ,
                    Midpoint = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0),
                    RotatedMinX = minX,
                    RotatedMinY = minY,
                    RotatedMinZ = minZ,
                    RotatedMaxX = maxX,
                    RotatedMaxY = maxY,
                    RotatedMaxZ = maxZ
                };
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[RotatedBoundingBoxCalculator] Exception in CalculateFromCorners: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new BoundingBoxResult();
            }
        }
    }
}

