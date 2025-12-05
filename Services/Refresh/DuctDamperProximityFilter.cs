using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Duct-Damper Proximity Filter
    /// 
    /// Implements: IMepIntersectionFilter, IDuctDamperProximityFilter
    /// Responsibility: ONLY filter ducts near dampers on same wall
    /// 
    /// This service implements the duct-damper combo logic where:
    /// - Ducts that have dampers within proximity on the same wall should be SKIPPED
    /// - Dampers are always processed (never skipped)
    /// - Both ducts and dampers are detected separately, but ducts near dampers are filtered OUT after detection
    /// 
    /// SOLID Principles:
    /// ✅ SRP: Single Responsibility = filter ducts near dampers
    /// ✅ OCP: Open/Closed = can be extended via interface inheritance
    /// ✅ LSP: Liskov Substitution = implements IMepIntersectionFilter interface
    /// ✅ ISP: Interface Segregation = has specialized IDuctDamperProximityFilter
    /// ✅ DIP: Dependency Inversion = depends on interfaces, injected via constructor
    /// </summary>
    public class DuctDamperProximityFilter : IDuctDamperProximityFilter
    {
        private const double INTERSECTION_TOLERANCE_FEET = 0.164;  // 50mm tolerance for damper at intersection point
        private const double BBOX_PROXIMITY_TOLERANCE_FEET = 0.164; // 50mm tolerance for connected duct-damper pairs

        /// <summary>
        /// ✅ IMepIntersectionFilter implementation
        /// Filter intersections by removing ducts that are in close proximity to dampers on the same wall.
        /// </summary>
        public List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> FilterIntersections(
            List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> allIntersections,
            Document doc,
            Action<string> logger)
        {
            return FilterDuctsNearDampers(allIntersections, doc, logger);
        }

        /// <summary>
        /// ✅ IDuctDamperProximityFilter implementation
        /// Filters out ducts that are in close proximity to dampers on the same wall.
        /// Returns list of zones to KEEP (i.e., ducts NOT near dampers, plus all dampers).
        /// </summary>
        public List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> FilterDuctsNearDampers(
            List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> allIntersections,
            Document doc,
            Action<string> logger)
        {
            if (allIntersections == null || allIntersections.Count == 0)
                return allIntersections;

            logger("[DUCT-DAMPER-FILTER] Starting duct-damper proximity filtering...");

            // ✅ STEP 1: Separate ducts from dampers
            var ducts = allIntersections
                .Where(i => !IsDamper(i.mepElement))
                .ToList();

            var dampers = allIntersections
                .Where(i => IsDamper(i.mepElement))
                .ToList();

            logger($"[DUCT-DAMPER-FILTER] Found {ducts.Count} ducts and {dampers.Count} dampers");

            if (dampers.Count == 0)
            {
                logger("[DUCT-DAMPER-FILTER] No dampers found - all ducts will be kept");
                return allIntersections; // Keep all if no dampers
            }

            // ✅ STEP 2: Build damper location cache (group by host wall for fast lookup)
            var dampersByWall = BuildDamperCache(dampers, logger);

            // ✅ STEP 3: Filter ducts - keep only those NOT near dampers on same wall
            var filteredDucts = new List<(Element mepElement, Element hostElement, BoundingBoxXYZ damperBBox, XYZ intersectionPoint)>();

            foreach (var duct in ducts)
            {
                if (!IsDuctNearDamperOnSameWall(duct, dampersByWall, doc, logger))
                {
                    filteredDucts.Add(duct);
                }
            }

            // ✅ STEP 4: Combine filtered ducts with all dampers
            var result = filteredDucts;
            result.AddRange(dampers);

            logger($"[DUCT-DAMPER-FILTER] ✅ Filtering complete: Removed {ducts.Count - filteredDucts.Count} ducts near dampers, keeping {filteredDucts.Count} ducts + {dampers.Count} dampers = {result.Count} total zones");

            return result;
        }

        /// <summary>
        /// Check if a duct is in close proximity to any damper on the same wall.
        /// Returns true if duct should be SKIPPED (i.e., damper found nearby).
        /// </summary>
        private bool IsDuctNearDamperOnSameWall(
            (Element mepElement, Element hostElement, BoundingBoxXYZ damperBBox, XYZ intersectionPoint) ductInfo,
            Dictionary<ElementId, List<(Element damperElement, BoundingBoxXYZ bbox, XYZ intersectionPoint)>> dampersByWall,
            Document doc,
            Action<string> logger)
        {
            try
            {
                var (ductElement, wallElement, hostBBox, ductIntersectionPoint) = ductInfo;
                var wallId = wallElement.Id;

                // ✅ OPTIMIZATION: Check if wall has any dampers first
                if (!dampersByWall.ContainsKey(wallId))
                {
                    return false; // No dampers on this wall
                }

                var dampersOnWall = dampersByWall[wallId];
                logger($"[DUCT-DAMPER-FILTER] 🔍 Checking DUCT(id={ductElement.Id}) on WALL(id={wallId}) against {dampersOnWall.Count} damper(s)");

                // Get duct bounding box directly from the duct element
                var ductBbox = ductElement.get_BoundingBox(null);
                if (ductBbox == null)
                {
                    logger($"[DUCT-DAMPER-FILTER] ⚠️ Duct {ductElement.Id} has no bounding box - cannot check damper proximity, keeping duct");
                    return false;
                }

                // Get duct center for distance calculation
                var ductCenter = (ductBbox.Min + ductBbox.Max) * 0.5;

                // Check each damper on the same wall
                foreach (var (damperElement, damperBbox, damperIntersectionPoint) in dampersOnWall)
                {
                    logger($"[DUCT-DAMPER-FILTER]   Testing DAMPER(id={damperElement.Id}) - intersection at ({ductIntersectionPoint.X:F3}, {ductIntersectionPoint.Y:F3}, {ductIntersectionPoint.Z:F3})");

                    // ✅ METHOD 1: Check if damper intersection point is near duct intersection point (same wall crossing)
                    double distanceBetweenIntersections = ductIntersectionPoint.DistanceTo(damperIntersectionPoint);
                    logger($"[DUCT-DAMPER-FILTER]     Distance between intersection points: {distanceBetweenIntersections:F4}ft");

                    if (distanceBetweenIntersections <= INTERSECTION_TOLERANCE_FEET)
                    {
                        logger($"[DUCT-DAMPER-FILTER] ✓ SKIP DUCT: Intersection points are {distanceBetweenIntersections:F4}ft apart (tolerance: {INTERSECTION_TOLERANCE_FEET}ft = 10mm) → DAMPER AT SAME LOCATION");
                        return true;
                    }

                    // ✅ METHOD 2: Check if duct center is near damper center (both close to wall)
                    XYZ damperCenter = (damperBbox.Min + damperBbox.Max) * 0.5;
                    double distanceBetweenCenters = ductCenter.DistanceTo(damperCenter);
                    logger($"[DUCT-DAMPER-FILTER]     Distance between element centers: {distanceBetweenCenters:F4}ft");

                    if (distanceBetweenCenters <= BBOX_PROXIMITY_TOLERANCE_FEET)
                    {
                        logger($"[DUCT-DAMPER-FILTER] ✓ SKIP DUCT: Element centers are {distanceBetweenCenters:F4}ft apart (tolerance: {BBOX_PROXIMITY_TOLERANCE_FEET}ft = 10mm) → ELEMENTS CLOSE TOGETHER");
                        return true;
                    }

                    // ✅ METHOD 3: Check if duct and damper bounding boxes are within proximity (for connected pairs)
                    var bboxDistance = GetMinimumDistanceBetweenBoundingBoxes(ductBbox, damperBbox);
                    logger($"[DUCT-DAMPER-FILTER]     Minimum bounding box distance: {bboxDistance:F4}ft");

                    if (bboxDistance <= BBOX_PROXIMITY_TOLERANCE_FEET)
                    {
                        logger($"[DUCT-DAMPER-FILTER] ✓ SKIP DUCT: Bounding boxes are {bboxDistance:F4}ft apart (tolerance: {BBOX_PROXIMITY_TOLERANCE_FEET}ft = 10mm) → CONNECTED PAIR");
                        return true;
                    }
                }

                logger($"[DUCT-DAMPER-FILTER] ✗ KEEP DUCT: {ductElement.Id} NOT near any dampers on wall {wallId}");
                return false;
            }
            catch (Exception ex)
            {
                logger($"[DUCT-DAMPER-FILTER] ⚠️ Error checking duct-damper proximity: {ex.Message} - keeping duct");
                return false;
            }
        }

        /// <summary>
        /// Build a cache of damper locations grouped by wall for efficient lookup.
        /// </summary>
        private Dictionary<ElementId, List<(Element damperElement, BoundingBoxXYZ bbox, XYZ intersectionPoint)>> BuildDamperCache(
            List<(Element mepElement, Element hostElement, BoundingBoxXYZ damperBBox, XYZ intersectionPoint)> dampers,
            Action<string> logger)
        {
            var cache = new Dictionary<ElementId, List<(Element, BoundingBoxXYZ, XYZ)>>();

            foreach (var (damperElement, wallElement, damperBBox, intersectionPoint) in dampers)
            {
                var wallId = wallElement.Id;
                if (!cache.ContainsKey(wallId))
                {
                    cache[wallId] = new List<(Element, BoundingBoxXYZ, XYZ)>();
                }

                cache[wallId].Add((damperElement, damperBBox, intersectionPoint));
            }

            logger($"[DUCT-DAMPER-FILTER] Built cache with dampers on {cache.Count} walls");
            return cache;
        }

        /// <summary>
        /// Check if an element is a damper (Duct Accessory category).
        /// </summary>
        private bool IsDamper(Element element)
        {
            try
            {
                return element?.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Check if a point is within tolerance distance of a bounding box.
        /// </summary>
        private bool IsPointNearBoundingBox(XYZ point, BoundingBoxXYZ bbox, double tolerance)
        {
            try
            {
                // Expand bbox by tolerance
                var expandedMin = new XYZ(bbox.Min.X - tolerance, bbox.Min.Y - tolerance, bbox.Min.Z - tolerance);
                var expandedMax = new XYZ(bbox.Max.X + tolerance, bbox.Max.Y + tolerance, bbox.Max.Z + tolerance);

                // Check if point is within expanded bbox
                return point.X >= expandedMin.X && point.X <= expandedMax.X &&
                       point.Y >= expandedMin.Y && point.Y <= expandedMax.Y &&
                       point.Z >= expandedMin.Z && point.Z <= expandedMax.Z;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Calculate minimum distance between two bounding boxes.
        /// </summary>
        private double GetMinimumDistanceBetweenBoundingBoxes(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            try
            {
                var min1 = bbox1.Min;
                var max1 = bbox1.Max;
                var min2 = bbox2.Min;
                var max2 = bbox2.Max;

                // Find closest points on each bbox
                var closest1 = new XYZ(
                    Math.Max(min1.X, Math.Min(max1.X, (min2.X + max2.X) / 2)),
                    Math.Max(min1.Y, Math.Min(max1.Y, (min2.Y + max2.Y) / 2)),
                    Math.Max(min1.Z, Math.Min(max1.Z, (min2.Z + max2.Z) / 2))
                );

                var closest2 = new XYZ(
                    Math.Max(min2.X, Math.Min(max2.X, (min1.X + max1.X) / 2)),
                    Math.Max(min2.Y, Math.Min(max2.Y, (min1.Y + max1.Y) / 2)),
                    Math.Max(min2.Z, Math.Min(max2.Z, (min1.Z + max1.Z) / 2))
                );

                return closest1.DistanceTo(closest2);
            }
            catch
            {
                return double.MaxValue;
            }
        }
    }
}
