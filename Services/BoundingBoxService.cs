using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized service for bounding box operations.
    /// This eliminates code duplication across MepIntersectionService, ClashZoneService, and other services.
    /// </summary>
    public static class BoundingBoxService
    {
        /// <summary>
        /// Checks if two bounding boxes intersect.
        /// </summary>
        /// <param name="bbox1">First bounding box</param>
        /// <param name="bbox2">Second bounding box</param>
        /// <returns>True if bounding boxes intersect, false otherwise</returns>
        public static bool BoundingBoxesIntersect(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            if (bbox1 == null || bbox2 == null)
                return false;

            return BoundingBoxesIntersect(bbox1.Min, bbox1.Max, bbox2.Min, bbox2.Max);
        }

        /// <summary>
        /// Checks if two bounding boxes intersect using min/max points.
        /// </summary>
        /// <param name="min1">Minimum point of first bounding box</param>
        /// <param name="max1">Maximum point of first bounding box</param>
        /// <param name="min2">Minimum point of second bounding box</param>
        /// <param name="max2">Maximum point of second bounding box</param>
        /// <returns>True if bounding boxes intersect, false otherwise</returns>
        public static bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            // Method 1: Check if boxes don't overlap (more intuitive)
            // Boxes intersect if they overlap on all axes
            return min1.X <= max2.X && max1.X >= min2.X &&
                   min1.Y <= max2.Y && max1.Y >= min2.Y &&
                   min1.Z <= max2.Z && max1.Z >= min2.Z;
        }

        /// <summary>
        /// Checks if two bounding boxes intersect with tolerance.
        /// </summary>
        /// <param name="box1Min">Minimum point of first bounding box</param>
        /// <param name="box1Max">Maximum point of first bounding box</param>
        /// <param name="box2Min">Minimum point of second bounding box</param>
        /// <param name="box2Max">Maximum point of second bounding box</param>
        /// <param name="tolerance">Tolerance for intersection check (default: 0.01)</param>
        /// <returns>True if bounding boxes intersect within tolerance, false otherwise</returns>
        public static bool BoundingBoxesIntersectWithTolerance(XYZ box1Min, XYZ box1Max, XYZ box2Min, XYZ box2Max, double tolerance = 0.01)
        {
            // Expand boxes by tolerance before checking intersection
            var expandedBox1Min = new XYZ(box1Min.X - tolerance, box1Min.Y - tolerance, box1Min.Z - tolerance);
            var expandedBox1Max = new XYZ(box1Max.X + tolerance, box1Max.Y + tolerance, box1Max.Z + tolerance);
            
            var expandedBox2Min = new XYZ(box2Min.X - tolerance, box2Min.Y - tolerance, box2Min.Z - tolerance);
            var expandedBox2Max = new XYZ(box2Max.X + tolerance, box2Max.Y + tolerance, box2Max.Z + tolerance);

            return BoundingBoxesIntersect(expandedBox1Min, expandedBox1Max, expandedBox2Min, expandedBox2Max);
        }

        /// <summary>
        /// Gets the center point of a bounding box.
        /// </summary>
        /// <param name="bbox">The bounding box</param>
        /// <returns>The center point XYZ</returns>
        public static XYZ GetBoundingBoxCenter(BoundingBoxXYZ bbox)
        {
            if (bbox == null)
                return XYZ.Zero;

            return new XYZ(
                (bbox.Min.X + bbox.Max.X) / 2,
                (bbox.Min.Y + bbox.Max.Y) / 2,
                (bbox.Min.Z + bbox.Max.Z) / 2
            );
        }

        /// <summary>
        /// Transforms a bounding box using the provided transform.
        /// Transforms all 8 corners of the box and computes a new bounding box.
        /// </summary>
        /// <param name="bbox">The bounding box to transform</param>
        /// <param name="transform">The transform to apply</param>
        /// <returns>The transformed bounding box, or null if transformation fails</returns>
        public static BoundingBoxXYZ? TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            if (bbox == null || transform == null)
                return null;

            try
            {
                // Transform all 8 corners of the bounding box
                var pts = new[]
                {
                    transform.OfPoint(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z)),
                    transform.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z)),
                    transform.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z)),
                    transform.OfPoint(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z)),
                    transform.OfPoint(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)),
                    transform.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z)),
                    transform.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z)),
                    transform.OfPoint(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z))
                };

                // Compute new min/max from transformed points
                var newMin = new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Min(p => p.Z));
                var newMax = new XYZ(pts.Max(p => p.X), pts.Max(p => p.Y), pts.Max(p => p.Z));

                return new BoundingBoxXYZ
                {
                    Min = newMin,
                    Max = newMax
                };
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[BoundingBoxService] Error transforming bounding box: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets an element's bounding box with optional transform for linked elements.
        /// </summary>
        /// <param name="element">The element</param>
        /// <param name="linkTransform">Optional transform for linked elements</param>
        /// <returns>The bounding box, or null if not available</returns>
        public static BoundingBoxXYZ? GetElementBoundingBox(Element element, Transform? linkTransform = null)
        {
            if (element == null)
                return null;

            var bbox = element.get_BoundingBox(null);
            if (bbox == null)
                return null;

            if (linkTransform != null && !linkTransform.IsIdentity)
            {
                return TransformBoundingBox(bbox, linkTransform);
            }

            return bbox;
        }
    }
}

