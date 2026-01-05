using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox
{
    /// <summary>
    /// Calculates axis-aligned cluster bounding boxes from database data.
    /// Phase 3: Extracted from UniversalClusterService.GetClusterBoundingBoxFromXml with crash-safe guards.
    /// </summary>
    public class AxisAlignedBoundingBoxCalculator : IBoundingBoxCalculator
    {
        private readonly Func<int, string, dynamic> _getClashZoneBySleeveInstanceId;
        private readonly Func<List<ClusteringSleeveDto>, string, double> _determineDominantRotationAngle;

        /// <summary>
        /// Constructor with dependencies for ClashZone lookup and rotation angle determination.
        /// </summary>
        public AxisAlignedBoundingBoxCalculator(
            Func<int, string, dynamic> getClashZoneBySleeveInstanceId,
            Func<List<ClusteringSleeveDto>, string, double> determineDominantRotationAngle)
        {
            _getClashZoneBySleeveInstanceId = getClashZoneBySleeveInstanceId ?? throw new ArgumentNullException(nameof(getClashZoneBySleeveInstanceId));
            _determineDominantRotationAngle = determineDominantRotationAngle ?? throw new ArgumentNullException(nameof(determineDominantRotationAngle));
        }

        /// <summary>
        /// Calculate axis-aligned cluster bounding box from sleeve data.
        /// Uses pre-calculated bounding boxes from database.
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
                if (cluster == null || cluster.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[AxisAlignedBoundingBoxCalculator] Invalid cluster input - returning empty result");
                    return new BoundingBoxResult();
                }

                // ✅ Determine dominant rotation angle if not provided
                if (Math.Abs(rotationAngle) < 1e-6)
                {
                    rotationAngle = _determineDominantRotationAngle(cluster, xmlFilePath);
                }

                // Get all bounding boxes from cluster data
                // Accessing BoundingBox directly from DTO (ClashZone properties) may require reconstruction if not already on DTO.
                // DTO has ClashZone. We can get BoundingBox from ClashZone properties: MinX, MinY, MinZ, MaxX, MaxY, MaxZ.
                var boundingBoxes = new List<BoundingBoxXYZ>();
                
                foreach (var item in cluster)
                {
                    var cz = item.ClashZone;
                    if (cz != null)
                    {
                        var bbox = new BoundingBoxXYZ();
                        bbox.Min = new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ);
                        bbox.Max = new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ);
                        boundingBoxes.Add(bbox);
                    }
                }

                if (boundingBoxes.Count == 0)
                {
                    return new BoundingBoxResult();
                }

                double minX, minY, minZ, maxX, maxY, maxZ;
                double width, height, depth;
                XYZ mid;

                // ✅ Calculate bounding box in rotated coordinate system if rotation angle is significant
                if (Math.Abs(rotationAngle) > 1e-6)
                {
                    // Calculate center point (midpoint of all sleeve centers) for rotation
                    var sleeveCenters = new List<XYZ>();
                    foreach (var bbox in boundingBoxes)
                    {
                        sleeveCenters.Add((bbox.Min + bbox.Max) / 2.0);
                    }

                    if (sleeveCenters.Count > 0)
                    {
                        // Use average center as rotation origin
                        XYZ rotationOrigin = new XYZ(
                            sleeveCenters.Average(p => p.X),
                            sleeveCenters.Average(p => p.Y),
                            sleeveCenters.Average(p => p.Z)
                        );

                        // Create rotation transform (rotate around Z-axis)
                        Transform rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, rotationAngle, rotationOrigin);
                        Transform inverseTransform = rotationTransform.Inverse;

                        // Transform all bounding box corners to rotated coordinate system
                        var transformedPoints = new List<XYZ>();

                        foreach (var bbox in boundingBoxes)
                        {
                            // Get all 8 corners of the bounding box
                            var corners = new[]
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

                            // Transform each corner to rotated coordinate system
                            foreach (var corner in corners)
                            {
                                var transformed = inverseTransform.OfPoint(corner);
                                transformedPoints.Add(transformed);
                            }
                        }

                        if (transformedPoints.Count > 0)
                        {
                            // Calculate min/max in rotated coordinate system
                            minX = transformedPoints.Min(p => p.X);
                            minY = transformedPoints.Min(p => p.Y);
                            minZ = transformedPoints.Min(p => p.Z);
                            maxX = transformedPoints.Max(p => p.X);
                            maxY = transformedPoints.Max(p => p.Y);
                            maxZ = transformedPoints.Max(p => p.Z);

                            // Dimensions in rotated coordinate system
                            width = maxX - minX;
                            height = maxY - minY;
                            depth = maxZ - minZ;

                            // Midpoint in rotated coordinate system (transform back to model coordinates)
                            XYZ rotatedMid = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);
                            mid = rotationTransform.OfPoint(rotatedMid);

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
                }

                // Fallback: Calculate overall bounding box (union of all boxes) - axis-aligned
                minX = boundingBoxes.Min(bbox => bbox.Min.X);
                minY = boundingBoxes.Min(bbox => bbox.Min.Y);
                minZ = boundingBoxes.Min(bbox => bbox.Min.Z);
                maxX = boundingBoxes.Max(bbox => bbox.Max.X);
                maxY = boundingBoxes.Max(bbox => bbox.Max.Y);
                maxZ = boundingBoxes.Max(bbox => bbox.Max.Z);

                width = maxX - minX;
                height = maxY - minY;
                depth = maxZ - minZ;
                mid = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);

                return new BoundingBoxResult
                {
                    Width = width,
                    Height = height,
                    Depth = depth,
                    Midpoint = mid
                };
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[AxisAlignedBoundingBoxCalculator] Exception in Calculate: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new BoundingBoxResult();
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
                        $"[AxisAlignedBoundingBoxCalculator] Invalid corners input - returning empty result");
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
                    $"[AxisAlignedBoundingBoxCalculator] Exception in CalculateFromCorners: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new BoundingBoxResult();
            }
        }
    }
}

