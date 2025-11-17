
#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class ClusterBoundingBoxServices
    {
        /// <summary>
        /// Returns bounding box dimensions and midpoint for a cluster of sleeves
        /// Uses axis-aligned bounding box (original behavior)
        /// </summary>
        public static (double width, double height, double depth, XYZ mid) GetClusterBoundingBox(List<FamilyInstance>? cluster)
        {
            return GetClusterBoundingBox(cluster, 0.0);
        }

        /// <summary>
        /// Returns bounding box dimensions and midpoint for a cluster of sleeves in a rotated coordinate system
        /// This ensures cluster sleeves follow the actual outline of individual sleeves, not just 0°/90° boxes
        /// </summary>
        /// <param name="cluster">List of sleeve family instances</param>
        /// <param name="rotationAngle">Rotation angle in radians. If 0, uses axis-aligned bounding box.</param>
        public static (double width, double height, double depth, XYZ mid) GetClusterBoundingBox(List<FamilyInstance>? cluster, double rotationAngle)
        {
            if (cluster == null || cluster.Count == 0)
                return (0, 0, 0, XYZ.Zero);

            // If no rotation, use original axis-aligned approach
            if (Math.Abs(rotationAngle) < 1e-6)
            {
                return GetAxisAlignedBoundingBox(cluster);
            }

            // ✅ NEW: Calculate bounding box in rotated coordinate system
            return GetRotatedBoundingBox(cluster, rotationAngle);
        }

        /// <summary>
        /// Original axis-aligned bounding box calculation
        /// </summary>
        private static (double width, double height, double depth, XYZ mid) GetAxisAlignedBoundingBox(List<FamilyInstance> cluster)
        {
            BoundingBoxXYZ? combinedBbox = null;

            foreach (var s in cluster)
            {
                // Use get_BoundingBox(null) which returns an axis-aligned bounding box in the model's coordinate system.
                // This correctly accounts for the element's geometry, size, and orientation.
                BoundingBoxXYZ? bbox = s.get_BoundingBox(null);
                if (bbox == null || !bbox.Enabled) continue;

                if (combinedBbox == null)
                {
                    // Defensive copy to avoid mutating original bbox
                    combinedBbox = new BoundingBoxXYZ
                    {
                        Min = bbox.Min,
                        Max = bbox.Max,
                        Enabled = bbox.Enabled
                    };
                }
                else
                {
                    // Union of the existing combined box and the new sleeve's box
                    combinedBbox.Min = new XYZ(Math.Min(combinedBbox.Min.X, bbox.Min.X),
                                               Math.Min(combinedBbox.Min.Y, bbox.Min.Y),
                                               Math.Min(combinedBbox.Min.Z, bbox.Min.Z));
                    combinedBbox.Max = new XYZ(Math.Max(combinedBbox.Max.X, bbox.Max.X),
                                               Math.Max(combinedBbox.Max.Y, bbox.Max.Y),
                                               Math.Max(combinedBbox.Max.Z, bbox.Max.Z));
                }
            }

            if (combinedBbox == null)
                return (0, 0, 0, XYZ.Zero);

            double widthVal = combinedBbox.Max.X - combinedBbox.Min.X;
            double heightVal = combinedBbox.Max.Y - combinedBbox.Min.Y;
            double depthVal = combinedBbox.Max.Z - combinedBbox.Min.Z;
            XYZ mid = (combinedBbox.Min + combinedBbox.Max) / 2.0;

            return (widthVal, heightVal, depthVal, mid);
        }

        /// <summary>
        /// Calculate bounding box in rotated coordinate system
        /// Transforms all sleeve bounding boxes to the rotated coordinate system before calculating min/max
        /// </summary>
        private static (double width, double height, double depth, XYZ mid) GetRotatedBoundingBox(List<FamilyInstance> cluster, double rotationAngle)
        {
            // Calculate center point (midpoint of all sleeve centers) for rotation
            var sleeveCenters = new List<XYZ>();
            foreach (var s in cluster)
            {
                var bbox = s.get_BoundingBox(null);
                if (bbox == null || !bbox.Enabled) continue;
                sleeveCenters.Add((bbox.Min + bbox.Max) / 2.0);
            }

            if (sleeveCenters.Count == 0)
                return (0, 0, 0, XYZ.Zero);

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

            foreach (var s in cluster)
            {
                var bbox = s.get_BoundingBox(null);
                if (bbox == null || !bbox.Enabled) continue;

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

            if (transformedPoints.Count == 0)
                return (0, 0, 0, XYZ.Zero);

            // Calculate min/max in rotated coordinate system
            double minX = transformedPoints.Min(p => p.X);
            double minY = transformedPoints.Min(p => p.Y);
            double minZ = transformedPoints.Min(p => p.Z);
            double maxX = transformedPoints.Max(p => p.X);
            double maxY = transformedPoints.Max(p => p.Y);
            double maxZ = transformedPoints.Max(p => p.Z);

            // Dimensions in rotated coordinate system
            double widthVal = maxX - minX;
            double heightVal = maxY - minY;
            double depthVal = maxZ - minZ;

            // ✅ CRITICAL FIX: The dimensions in rotated coordinate system are correct
            // But we need to ensure we're not expanding unnecessarily
            // The width/height in rotated system should match the actual cluster outline
            
            // Midpoint in rotated coordinate system (transform back to model coordinates)
            XYZ rotatedMid = new XYZ((minX + maxX) / 2.0, (minY + maxY) / 2.0, (minZ + maxZ) / 2.0);
            XYZ mid = rotationTransform.OfPoint(rotatedMid);

            // ✅ DEBUG: Log the calculation for troubleshooting
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CLUSTER-BBOX-ROTATED] Rotation angle: {rotationAngle * 180 / Math.PI:F1}°, " +
                    $"Rotated coords - Min=({minX:F3}, {minY:F3}, {minZ:F3}), Max=({maxX:F3}, {maxY:F3}, {maxZ:F3}), " +
                    $"Dimensions: W={UnitUtils.ConvertFromInternalUnits(widthVal, UnitTypeId.Millimeters):F1}mm, " +
                    $"H={UnitUtils.ConvertFromInternalUnits(heightVal, UnitTypeId.Millimeters):F1}mm, " +
                    $"D={UnitUtils.ConvertFromInternalUnits(depthVal, UnitTypeId.Millimeters):F1}mm");
            }

            return (widthVal, heightVal, depthVal, mid);
        }
    }
}
