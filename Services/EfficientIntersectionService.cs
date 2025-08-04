
#nullable enable
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using System.Collections.Generic;
using System.Linq;
using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized efficient intersection service with section box filtering and bounding box pre-checks
    /// </summary>
    public static class EfficientIntersectionService
    {
        /// <summary>
        /// Find wall intersections with MEP element using section box filtering and bounding box pre-checks
        /// </summary>
        public static List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)> FindWallIntersections(
            MEPCurve mepElement, 
            Line mepLine, 
            View3D view3D,
            List<XYZ> testPoints,
            XYZ rayDirection)
        {
            var allWallHits = new List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)>();
            
            // Get section box bounds for filtering
            var sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
            if (sectionBox == null)
            {
                DebugLogger.Log("[EfficientIntersectionService] No section box found, using full model");
            }
            
            // Create MEP bounding box for pre-filtering
            var mepBBox = GetMepElementBoundingBox(mepElement, mepLine);
            
            // Pre-filter walls by section box and bounding box intersection
            var filteredWalls = GetFilteredWallsInSectionBox(mepElement.Document, view3D, sectionBox ?? new BoundingBoxXYZ(), mepBBox);
            
            DebugLogger.Log($"[EfficientIntersectionService] Filtered walls: {filteredWalls.Count} (from section box and bbox pre-check)");
            
            // Create optimized ReferenceIntersector with filtered elements
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
            {
                FindReferencesInRevitLinks = true
            };
            
            // Cast rays from test points with early termination
            foreach (var testPoint in testPoints)
            {
                // Skip if test point is outside section box
                if (sectionBox != null && !IsPointInBoundingBox(testPoint, sectionBox))
                {
                    continue;
                }
                
                // Cast in primary directions
                var hitsFwd = refIntersector.Find(testPoint, rayDirection)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(3).ToList();
                var hitsBack = refIntersector.Find(testPoint, rayDirection.Negate())?.Where(h => h != null).OrderBy(h => h.Proximity).Take(3).ToList();
                
                // Cast in perpendicular directions for parallel walls
                var perpDir1 = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                var perpDir2 = perpDir1.Negate();
                var hitsPerp1 = refIntersector.Find(testPoint, perpDir1)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList();
                var hitsPerp2 = refIntersector.Find(testPoint, perpDir2)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList();
                
                // Add hits with their ray origin
                if (hitsFwd?.Any() == true) allWallHits.AddRange(hitsFwd.Select(h => (h, rayDirection, testPoint)));
                if (hitsBack?.Any() == true) allWallHits.AddRange(hitsBack.Select(h => (h, rayDirection.Negate(), testPoint)));
                if (hitsPerp1?.Any() == true) allWallHits.AddRange(hitsPerp1.Select(h => (h, perpDir1, testPoint)));
                if (hitsPerp2?.Any() == true) allWallHits.AddRange(hitsPerp2.Select(h => (h, perpDir2, testPoint)));
            }
            
            return allWallHits;
        }
        
        /// <summary>
        /// Find structural intersections with MEP element using efficient bounding box pre-filtering
        /// </summary>
        public static List<(Element structuralElement, BoundingBoxXYZ bbox, XYZ intersectionPoint)> FindStructuralIntersections(
            MEPCurve mepElement,
            Line mepLine,
            View3D view3D)
        {
            var intersections = new List<(Element, BoundingBoxXYZ, XYZ)>();
            
            // Get section box bounds for filtering
            var sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
            
            // Get MEP bounding box for pre-filtering
            var mepBBox = GetMepElementBoundingBox(mepElement, mepLine);
            
            // Get pre-filtered structural elements
            var structuralElements = GetFilteredStructuralElementsInSectionBox(mepElement.Document, view3D, sectionBox, mepBBox);
            
            DebugLogger.Log($"[EfficientIntersectionService] Processing {structuralElements.Count} pre-filtered structural elements");
            
            // Process each structural element with solid intersection
            foreach (var (structuralElement, linkTransform) in structuralElements)
            {
                try
                {
                    // Get element bounding box and check intersection with MEP bbox
                    var elementBBox = GetElementBoundingBox(structuralElement, linkTransform);
                    if (elementBBox == null || !BoundingBoxesIntersect(mepBBox, elementBBox))
                    {
                        continue; // Skip expensive solid intersection if bboxes don't intersect
                    }
                    
                    // Perform solid intersection check
                    var solidIntersections = PerformSolidIntersection(mepLine, structuralElement, linkTransform);
                    if (solidIntersections.Any())
                    {
                        var midpoint = CalculateIntersectionMidpoint(solidIntersections);
                        intersections.Add((structuralElement, elementBBox, midpoint));
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[EfficientIntersectionService] Error processing structural element {structuralElement.Id.Value}: {ex.Message}");
                }
            }
            
            return intersections;
        }
        
        /// <summary>
        /// Get MEP element bounding box for spatial filtering
        /// </summary>
        private static BoundingBoxXYZ GetMepElementBoundingBox(MEPCurve mepElement, Line mepLine)
        {
            try
            {
                var bbox = mepElement.get_BoundingBox(null);
                if (bbox != null)
                {
                    // Expand bbox slightly for clearance
                    var expansion = 2.0; // 2 feet expansion
                    bbox.Min = new XYZ(bbox.Min.X - expansion, bbox.Min.Y - expansion, bbox.Min.Z - expansion);
                    bbox.Max = new XYZ(bbox.Max.X + expansion, bbox.Max.Y + expansion, bbox.Max.Z + expansion);
                    return bbox;
                }
            }
            catch
            {
                // Fallback: create bbox from line endpoints
            }
            
            // Fallback: create bounding box from line with expansion
            var p1 = mepLine.GetEndPoint(0);
            var p2 = mepLine.GetEndPoint(1);
            var expansion2 = 2.0;
            
            return new BoundingBoxXYZ
            {
                Min = new XYZ(
                    Math.Min(p1.X, p2.X) - expansion2,
                    Math.Min(p1.Y, p2.Y) - expansion2,
                    Math.Min(p1.Z, p2.Z) - expansion2),
                Max = new XYZ(
                    Math.Max(p1.X, p2.X) + expansion2,
                    Math.Max(p1.Y, p2.Y) + expansion2,
                    Math.Max(p1.Z, p2.Z) + expansion2)
            };
        }
        
        /// <summary>
        /// Get element bounding box with optional transform
        /// </summary>
        private static BoundingBoxXYZ? GetElementBoundingBox(Element element, Transform? linkTransform)
        {
            var bbox = element.get_BoundingBox(null);
            if (bbox == null) return null;

            var transform = linkTransform ?? Transform.Identity;
            if (!transform.IsIdentity)
            {
                // Transform bounding box for linked elements
                var transformedMin = transform.OfPoint(bbox.Min);
                var transformedMax = transform.OfPoint(bbox.Max);

                return new BoundingBoxXYZ
                {
                    Min = new XYZ(
                        Math.Min(transformedMin.X, transformedMax.X),
                        Math.Min(transformedMin.Y, transformedMax.Y),
                        Math.Min(transformedMin.Z, transformedMax.Z)),
                    Max = new XYZ(
                        Math.Max(transformedMin.X, transformedMax.X),
                        Math.Max(transformedMin.Y, transformedMax.Y),
                        Math.Max(transformedMin.Z, transformedMax.Z))
                };
            }

            return bbox;
        }
        
        /// <summary>
        /// Check if two bounding boxes intersect
        /// </summary>
        private static bool BoundingBoxesIntersect(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            return !(bbox1.Max.X < bbox2.Min.X || bbox2.Max.X < bbox1.Min.X ||
                     bbox1.Max.Y < bbox2.Min.Y || bbox2.Max.Y < bbox1.Min.Y ||
                     bbox1.Max.Z < bbox2.Min.Z || bbox2.Max.Z < bbox1.Min.Z);
        }
        
        /// <summary>
        /// Check if point is within bounding box
        /// </summary>
        private static bool IsPointInBoundingBox(XYZ point, BoundingBoxXYZ bbox)
        {
            return point.X >= bbox.Min.X && point.X <= bbox.Max.X &&
                   point.Y >= bbox.Min.Y && point.Y <= bbox.Max.Y &&
                   point.Z >= bbox.Min.Z && point.Z <= bbox.Max.Z;
        }
        
        /// <summary>
        /// Get filtered walls within section box and intersecting MEP bounding box
        /// </summary>
        private static List<(Element, Transform)> GetFilteredWallsInSectionBox(
            Document doc,
            View3D view3D,
            BoundingBoxXYZ sectionBox,
            BoundingBoxXYZ mepBBox)
        {
            var filteredWalls = new List<(Element, Transform)>();

            // Get walls from host document
            var hostWalls = new FilteredElementCollector(doc, view3D.Id)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Where(w => IsElementInSectionBoxAndIntersectsMep(w, Transform.Identity, sectionBox, mepBBox))
                .Select(w => (w, Transform.Identity))
                .ToList();

            filteredWalls.AddRange(hostWalls);

            // Get walls from visible linked documents
            foreach (var linkInstance in new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => link.GetLinkDocument() != null))
            {
                var linkDoc = linkInstance.GetLinkDocument();
                var linkTransform = linkInstance.GetTotalTransform();

                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Where(w => IsElementInSectionBoxAndIntersectsMep(w, linkTransform, sectionBox, mepBBox))
                    .Select(w => (w as Element, linkTransform ?? Transform.Identity))
                    .ToList();

                filteredWalls.AddRange(linkedWalls);
            }

            return filteredWalls;
        }
        
        /// <summary>
        /// Get filtered structural elements within section box and intersecting MEP bounding box
        /// </summary>
        private static List<(Element, Transform)> GetFilteredStructuralElementsInSectionBox(
            Document doc,
            View3D view3D,
            BoundingBoxXYZ? sectionBox,
            BoundingBoxXYZ mepBBox)
        {
            var filteredElements = new List<(Element, Transform)>();
            if (doc == null || mepBBox == null)
                return filteredElements;
            // Use the document directly (like duct logic) to collect from both host and visible links
            var structuralElements = StructuralElementCollectorHelper.CollectStructuralElementsVisibleOnly(doc);
            var hostFiltered = structuralElements
                .Where(tuple => tuple.Item1 != null && IsElementInSectionBoxAndIntersectsMep(tuple.Item1, tuple.Item2 ?? Transform.Identity, sectionBox, mepBBox))
                .Select(tuple => (tuple.Item1, tuple.Item2 ?? Transform.Identity))
                .ToList();
            filteredElements.AddRange(hostFiltered);
            DebugLogger.Log($"[EfficientIntersectionService] Pre-filtered structural elements: {filteredElements.Count} (from {structuralElements.Count} total)");
            return filteredElements;
        }
        
        /// <summary>
        /// Check if element is in section box and intersects MEP bounding box
        /// </summary>
        private static bool IsElementInSectionBoxAndIntersectsMep(
            Element element,
            Transform? linkTransform,
            BoundingBoxXYZ? sectionBox,
            BoundingBoxXYZ mepBBox)
        {
            if (element == null || mepBBox == null)
                return false;
            var elementBBox = GetElementBoundingBox(element, linkTransform ?? Transform.Identity);
            if (elementBBox == null) return false;

            // Check section box intersection (if section box exists)
            if (sectionBox != null && !BoundingBoxesIntersect(elementBBox, sectionBox))
            {
                return false;
            }

            // Check MEP bounding box intersection
            return BoundingBoxesIntersect(elementBBox, mepBBox);
        }
        
        /// <summary>
        /// Perform solid intersection between line and element
        /// </summary>
        private static List<XYZ> PerformSolidIntersection(Line mepLine, Element structuralElement, Transform? linkTransform)
        {
            var intersectionPoints = new List<XYZ>();

            try
            {
                var structuralOptions = new Options();
                var structuralGeometry = structuralElement.get_Geometry(structuralOptions);
                Solid? structuralSolid = null;

                foreach (var geomObj in structuralGeometry)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        structuralSolid = solid;
                        break;
                    }
                    else if (geomObj is GeometryInstance instance)
                    {
                        foreach (var instObj in instance.GetInstanceGeometry())
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > 0)
                            {
                                structuralSolid = instSolid;
                                break;
                            }
                        }
                        if (structuralSolid != null) break;
                    }
                }

                if (structuralSolid == null) return intersectionPoints;

                // Apply transform if linked element
                if (linkTransform != null && !linkTransform.IsIdentity)
                {
                    structuralSolid = SolidUtils.CreateTransformed(structuralSolid, linkTransform);
                }

                // Check intersection using face.Intersect(line)
                foreach (Face face in structuralSolid.Faces)
                {
                    IntersectionResultArray? ira = null;
                    SetComparisonResult res = face.Intersect(mepLine, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        foreach (IntersectionResult ir in ira)
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[EfficientIntersectionService] Solid intersection error for element {structuralElement.Id.Value}: {ex.Message}");
            }

            return intersectionPoints;
        }
        
        /// <summary>
        /// Calculate midpoint from intersection points
        /// </summary>
        private static XYZ CalculateIntersectionMidpoint(List<XYZ> intersectionPoints)
        {
            if (intersectionPoints.Count == 1)
            {
                return intersectionPoints[0];
            }

            if (intersectionPoints.Count >= 2)
            {
                // Find the two points with maximum distance
                double maxDist = double.MinValue;
                XYZ? ptA = null, ptB = null;

                for (int i = 0; i < intersectionPoints.Count - 1; i++)
                {
                    for (int j = i + 1; j < intersectionPoints.Count; j++)
                    {
                        double dist = intersectionPoints[i].DistanceTo(intersectionPoints[j]);
                        if (dist > maxDist)
                        {
                            maxDist = dist;
                            ptA = intersectionPoints[i];
                            ptB = intersectionPoints[j];
                        }
                    }
                }

                if (ptA != null && ptB != null)
                {
                    return new XYZ((ptA.X + ptB.X) / 2, (ptA.Y + ptB.Y) / 2, (ptA.Z + ptB.Z) / 2);
                }
            }

            return intersectionPoints.FirstOrDefault() ?? XYZ.Zero;
        }
    }
}
