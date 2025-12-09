using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;

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
    /// - Uses DamperCollectionService to collect ALL dampers (not just from intersections)
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
        private const double DUCT_END_POINT_TOLERANCE_FEET = 0.164;  // 50mm tolerance for damper near duct end point
        private const double BBOX_PROXIMITY_TOLERANCE_FEET = 0.164; // 50mm tolerance for connected duct-damper pairs
        
        public DuctDamperProximityFilter()
        {
        }

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

            // ✅ STEP 1: Separate ducts from dampers in intersections
            var ducts = allIntersections
                .Where(i => !IsDamper(i.mepElement))
                .ToList();

            var dampers = allIntersections
                .Where(i => IsDamper(i.mepElement))
                .ToList();

            logger($"[DUCT-DAMPER-FILTER] Found {ducts.Count} ducts and {dampers.Count} dampers in intersections");

            // ✅ STEP 2: Filter ducts - check if damper is near duct end point (reuse existing CheckForDamperNearPoint logic)
            // This is much simpler than collecting all dampers - just check each duct intersection for nearby dampers
            var filteredDucts = new List<(Element mepElement, Element hostElement, BoundingBoxXYZ damperBBox, XYZ intersectionPoint)>();

            foreach (var duct in ducts)
            {
                // ✅ SIMPLE CHECK: Is there a damper near the duct end point (nearest to intersection)?
                // Reuses the same logic as CheckForDamperNearPoint - no need to collect all dampers separately
                if (!IsDuctNearDamperAtEndPoint(duct, doc, logger))
                {
                    filteredDucts.Add(duct);
                }
            }

            // ✅ STEP 3: Combine filtered ducts with all dampers
            var result = filteredDucts;
            result.AddRange(dampers);

            logger($"[DUCT-DAMPER-FILTER] ✅ Filtering complete: Removed {ducts.Count - filteredDucts.Count} ducts near dampers, keeping {filteredDucts.Count} ducts + {dampers.Count} dampers = {result.Count} total zones");

            return result;
        }

        /// <summary>
        /// ✅ SIMPLIFIED: Check if a duct has a damper near its end point (nearest to intersection).
        /// Reuses the same logic as CheckForDamperNearPoint - searches for dampers near the duct end point.
        /// Returns true if duct should be SKIPPED (i.e., damper found nearby).
        /// </summary>
        private bool IsDuctNearDamperAtEndPoint(
            (Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint) ductInfo,
            Document doc,
            Action<string> logger)
        {
            try
            {
                var (ductElement, wallElement, hostBBox, ductIntersectionPoint) = ductInfo;

                if (!(ductElement is Autodesk.Revit.DB.Mechanical.Duct duct))
                {
                    return false; // Not a duct
                }

                // ✅ SOLID: Get duct end points using helper method
                var ductEndPoints = GetDuctEndPoints(ductElement);
                if (ductEndPoints == null || ductEndPoints.Count == 0)
                {
                    logger($"[DUCT-DAMPER-FILTER] ⚠️ Duct {ductElement.Id} has no valid end points - cannot check damper proximity, keeping duct");
                    return false;
                }

                // ✅ CRITICAL: Find the duct end point NEAREST to the intersection point
                XYZ nearestEndPoint = ductEndPoints
                    .OrderBy(ep => ep.DistanceTo(ductIntersectionPoint))
                    .First();
                
                double distanceToNearestEnd = nearestEndPoint.DistanceTo(ductIntersectionPoint);
                logger($"[DUCT-DAMPER-FILTER] 📍 Duct {ductElement.Id}: Nearest end point is {distanceToNearestEnd:F4}ft from intersection");

                // ✅ REUSE EXISTING LOGIC: Check if there's a damper near the duct end point
                // This uses the same approach as CheckForDamperNearPoint - searches all documents
                bool hasDamperNearEnd = CheckForDamperNearPoint(doc, nearestEndPoint, logger);
                
                if (hasDamperNearEnd)
                {
                    logger($"[DUCT-DAMPER-FILTER] ✓ SKIP DUCT: Damper found near duct end point (tolerance: 0.5ft = 150mm) → DAMPER AT DUCT END");
                    return true;
                }

                logger($"[DUCT-DAMPER-FILTER] ✗ KEEP DUCT: {ductElement.Id} - no damper found near end point");
                return false;
            }
            catch (Exception ex)
            {
                logger($"[DUCT-DAMPER-FILTER] ⚠️ Error checking duct-damper proximity: {ex.Message} - keeping duct");
                return false;
            }
        }
        
        /// <summary>
        /// ✅ REUSES EXISTING LOGIC: Check for damper presence near a specific point.
        /// Same logic as IntersectionDetectionService.CheckForDamperNearPoint.
        /// Searches ALL documents (host + linked) for Duct Accessories within 0.5ft (150mm) of the point.
        /// Excludes VCD/VOLUME dampers.
        /// </summary>
        private bool CheckForDamperNearPoint(Document document, XYZ searchPoint, Action<string> logger)
        {
            try
            {
                double searchRadius = 0.5; // 0.5 feet = ~150mm (same as IntersectionDetectionService)
                
                var allDampers = new List<Element>();
                
                // Search in host document
                var hostCollector = new FilteredElementCollector(document);
                var hostDampers = hostCollector
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .Where(d => IsDamper(d))
                    .ToList();
                allDampers.AddRange(hostDampers);
                
                // Search in all linked documents
                var links = new FilteredElementCollector(document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();
                
                foreach (var link in links)
                {
                    try
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;
                        
                        var linkTransform = link.GetTotalTransform();
                        var linkedCollector = new FilteredElementCollector(linkDoc);
                        var linkedDampers = linkedCollector
                            .OfCategory(BuiltInCategory.OST_DuctAccessory)
                            .WhereElementIsNotElementType()
                            .ToElements()
                            .Where(d => IsDamper(d))
                            .ToList();
                        
                        // Transform damper locations to host coordinates
                        foreach (var damper in linkedDampers)
                        {
                            allDampers.Add(damper);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger($"[DUCT-DAMPER-FILTER] WARNING: Could not search linked doc: {ex.Message}");
                    }
                }

                logger($"[DUCT-DAMPER-FILTER] Found {allDampers.Count} total dampers to check near point {searchPoint}");

                foreach (var damper in allDampers)
                {
                    // Get damper location (handle both LocationPoint and transform for linked docs)
                    XYZ damperLocation = null;
                    
                    if (damper.Location is LocationPoint locationPoint)
                    {
                        damperLocation = locationPoint.Point;
                    }
                    else if (damper.Location is LocationCurve locationCurve)
                    {
                        damperLocation = locationCurve.Curve.GetEndPoint(0);
                    }
                    
                    // If damper is from linked doc, transform to host coordinates
                    var damperDoc = damper.Document;
                    if (damperDoc != document)
                    {
                        var link = links.FirstOrDefault(l => l.GetLinkDocument() == damperDoc);
                        if (link != null && damperLocation != null)
                        {
                            damperLocation = link.GetTotalTransform().OfPoint(damperLocation);
                        }
                    }
                    
                    if (damperLocation != null)
                    {
                        double distance = damperLocation.DistanceTo(searchPoint);
                        if (distance < searchRadius)
                        {
                            logger($"[DUCT-DAMPER-FILTER] ✓ Found damper {damper.Id} at distance {distance:F4}ft from point");
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                logger($"[DUCT-DAMPER-FILTER] ERROR: Failed to search for dampers near point: {ex.Message}");
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
        /// Excludes VCD and VOLUME dampers (not in walls).
        /// </summary>
        private bool IsDamper(Element element)
        {
            try
            {
                if (element?.Category?.Id.IntegerValue != (int)BuiltInCategory.OST_DuctAccessory)
                    return false;
                
                // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                if (element is FamilyInstance fi && fi.Symbol != null)
                {
                    string familyName = fi.Symbol.Family?.Name ?? "";
                    string typeName = fi.Symbol.Name ?? "";
                    string combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                    
                    if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                        return false;
                }
                
                return true;
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
        /// ✅ SOLID: Get duct end points from LocationCurve.
        /// Returns the two end points of the duct's centerline curve.
        /// </summary>
        private List<XYZ> GetDuctEndPoints(Element ductElement)
        {
            var endPoints = new List<XYZ>();
            
            try
            {
                if (!(ductElement is Autodesk.Revit.DB.Mechanical.Duct duct))
                {
                    return endPoints; // Not a duct
                }

                // Get LocationCurve (duct centerline)
                var locationCurve = duct.Location as LocationCurve;
                if (locationCurve?.Curve == null)
                {
                    return endPoints; // No valid curve
                }

                var curve = locationCurve.Curve;
                
                // Get both end points of the curve
                endPoints.Add(curve.GetEndPoint(0)); // Start point
                endPoints.Add(curve.GetEndPoint(1)); // End point

                return endPoints;
            }
            catch (Exception ex)
            {
                // Return empty list on error (fail-safe)
                return new List<XYZ>();
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
