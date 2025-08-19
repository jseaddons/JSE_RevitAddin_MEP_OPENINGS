
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
        // ...existing code...
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

            // (diagnostics moved to after mepBBox is computed)
            
            // Create MEP bounding box for pre-filtering
            var mepBBox = GetMepElementBoundingBox(mepElement, mepLine);

            // Diagnostic: log section box and MEP bbox to help detect coordinate-frame mismatches
            try
            {
                var sb = sectionBox;
                DebugLogger.Log($"[EfficientIntersectionService] DIAG SectionBox: {(sb == null ? "<null>" : $"Min={FormatXYZ(sb.Min)}, Max={FormatXYZ(sb.Max)}")}");
                DebugLogger.Log($"[EfficientIntersectionService] DIAG MEP BBox: Min={FormatXYZ(mepBBox.Min)}, Max={FormatXYZ(mepBBox.Max)}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[EfficientIntersectionService] DIAG error while logging section/mep bbox: {ex.Message}");
            }
            
            // Pre-filter walls by section box and bounding box intersection
            var filteredWalls = GetFilteredWallsInSectionBox(mepElement.Document, view3D, sectionBox ?? new BoundingBoxXYZ(), mepBBox);
            
            DebugLogger.Log($"[EfficientIntersectionService] Filtered walls: {filteredWalls.Count} (from section box and bbox pre-check)");
            
            // Create optimized ReferenceIntersector with filtered elements
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
            {
                FindReferencesInRevitLinks = true
            };
            
            // If the MEP element lives in a linked document, resolve the link transform
            // from the host view so we can transform MEP test points into host coordinates.
            // ReferenceIntersector with FindReferencesInRevitLinks=true will handle linked walls automatically.
            Transform linkTransform = Transform.Identity;
            bool hasLinkTransform = false;
            try
            {
                if (mepElement.Document != view3D.Document)
                {
                    var linkInstanceInHost = new FilteredElementCollector(view3D.Document)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .FirstOrDefault(li => li.GetLinkDocument() != null && li.GetLinkDocument().Equals(mepElement.Document));

                    if (linkInstanceInHost != null)
                    {
                        linkTransform = linkInstanceInHost.GetTotalTransform();
                        hasLinkTransform = linkTransform != null && !linkTransform.IsIdentity;
                        var originStr = linkTransform != null ? FormatXYZ(linkTransform.Origin) : "<identity>";
                        var linkIdStr = linkInstanceInHost?.Id.IntegerValue.ToString() ?? "<no-id>";
                        DebugLogger.Log($"[EfficientIntersectionService] DIAG using MEP link transform from RevitLinkInstance Id={linkIdStr}, Origin={originStr}");
                    }
                    else
                    {
                        DebugLogger.Log($"[EfficientIntersectionService] DIAG WARNING: No RevitLinkInstance found for MEP document '{mepElement.Document.Title}'");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[EfficientIntersectionService] DIAG error resolving MEP link transform: {ex.Message}");
            }
            
            // Cast rays from test points with early termination
            foreach (var testPoint in testPoints)
            {
                // For linked MEP elements, we need to transform test points to host coordinates
                // since ReferenceIntersector operates in the view's coordinate system.
                // However, ReferenceIntersector with FindReferencesInRevitLinks=true will automatically
                // find both host and linked walls regardless of which document they're in.
                var testPointHost = testPoint;
                var rayDirHost = rayDirection;
                if (hasLinkTransform && linkTransform != null)
                {
                    try
                    {
                        testPointHost = linkTransform.OfPoint(testPoint);
                        rayDirHost = linkTransform.OfVector(rayDirection);
                        DebugLogger.Log($"[EfficientIntersectionService] DIAG transformed test point: {FormatXYZ(testPoint)} -> {FormatXYZ(testPointHost)}");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[EfficientIntersectionService] DIAG error transforming test point/ray to host: {ex.Message}");
                        testPointHost = testPoint;
                        rayDirHost = rayDirection;
                    }
                }

                // Skip section box check for linked MEP elements since the transform may move
                // test points far outside the section box even though the MEP element itself is within it.
                if (sectionBox != null && !hasLinkTransform && !IsPointInBoundingBox(testPoint, sectionBox))
                {
                    continue;
                }

                // TEMPORARY FIX: Use original test points since ReferenceIntersector handles linked discovery
                // TODO: Investigate proper coordinate space for ray casting with linked elements
                var rayOrigin = hasLinkTransform ? testPoint : testPointHost;
                var rayDir = hasLinkTransform ? rayDirection : rayDirHost;

                // Cast in primary directions
                var hitsFwd = refIntersector.Find(rayOrigin, rayDir)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(3).ToList();
                var hitsBack = refIntersector.Find(rayOrigin, rayDir.Negate())?.Where(h => h != null).OrderBy(h => h.Proximity).Take(3).ToList();
                
                // Cast in perpendicular directions for parallel walls (safe: skip when rayDirection has no horizontal component)
                var eps = 1e-9;
                bool hasHorizontal = Math.Abs(rayDir.X) > eps || Math.Abs(rayDir.Y) > eps;
                XYZ? perpDir1 = null;
                XYZ? perpDir2 = null;
                if (hasHorizontal)
                {
                    var tmp = new XYZ(-rayDir.Y, rayDir.X, 0);
                    var lenTmp = Math.Sqrt(tmp.X * tmp.X + tmp.Y * tmp.Y + tmp.Z * tmp.Z);
                    if (lenTmp > eps)
                    {
                        perpDir1 = new XYZ(tmp.X / lenTmp, tmp.Y / lenTmp, tmp.Z / lenTmp);
                        perpDir2 = perpDir1.Negate();
                    }
                }

                var hitsPerp1 = perpDir1 != null ? refIntersector.Find(rayOrigin, perpDir1)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList() : null;
                var hitsPerp2 = perpDir2 != null ? refIntersector.Find(rayOrigin, perpDir2)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList() : null;

                // Aggregate hits by element id and by direction seen
                var hitMap = new Dictionary<int, List<(ReferenceWithContext hit, XYZ dir)>>();
                void AddHitsToMap(IEnumerable<ReferenceWithContext>? hits, XYZ dir)
                {
                    if (hits == null) return;
                    foreach (var h in hits)
                    {
                        if (h == null) continue;
                        var eid = h.GetReference()?.ElementId.IntegerValue ?? -1;
                        if (eid == -1) continue;
                        if (!hitMap.TryGetValue(eid, out var list))
                        {
                            list = new List<(ReferenceWithContext, XYZ)>();
                            hitMap[eid] = list;
                        }
                        list.Add((h, dir));
                    }
                }
        
                AddHitsToMap(hitsFwd, rayDir);
                AddHitsToMap(hitsBack, rayDir.Negate());
                if (perpDir1 != null) AddHitsToMap(hitsPerp1, perpDir1);
                if (perpDir2 != null) AddHitsToMap(hitsPerp2, perpDir2);
        
                // Consider a wall a valid penetration if it is seen from two or more distinct ray directions
                foreach (var kv in hitMap)
                {
                    var dirs = kv.Value
                        .Select(v => v.dir)
                        .Select(d =>
                        {
                            var len = Math.Sqrt(d.X * d.X + d.Y * d.Y + d.Z * d.Z);
                            return len > eps ? new XYZ(d.X / len, d.Y / len, d.Z / len) : null;
                        })
                        .Where(d => d != null)
                        .Cast<XYZ>()
                        .ToList();

                    // count distinct directions using dot-product (angle test)
                    var distinct = new List<XYZ>();
                    double dirSameThreshold = 0.9995; // cosine threshold for considering two directions equivalent
                    foreach (var d in dirs)
                    {
                        bool isNew = true;
                        foreach (var ex in distinct)
                        {
                            if (Math.Abs(ex.DotProduct(d)) > dirSameThreshold)
                            {
                                isNew = false;
                                break;
                            }
                        }
                        if (isNew) distinct.Add(d);
                    }

                    if (distinct.Count < 2) continue;

                    // pick the closest hit among the collected hits for this element id
                    var representative = kv.Value.OrderBy(v => v.hit.Proximity).First();
                    // store the ray origin used for clarity downstream
                    allWallHits.Add((representative.hit, representative.dir, rayOrigin));
                }
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
                    DebugLogger.Log($"[EfficientIntersectionService] Error processing structural element {structuralElement.Id.IntegerValue}: {ex.Message}");
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
        /// Format XYZ for diagnostic logging
        /// </summary>
        private static string FormatXYZ(XYZ p)
        {
            if (p == null) return "<null>";
            return string.Format("({0:0.000},{1:0.000},{2:0.000})", p.X, p.Y, p.Z);
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

            // Transform-aware prefilter: convert the provided MEP bbox into the host view's coordinate space
            // so bounding-box comparisons are performed in the same frame as element bboxes.
            var mepBBoxInHost = mepBBox;
            try
            {
                if (doc != view3D.Document && mepBBox != null)
                {
                    var linkInstanceInHost = new FilteredElementCollector(view3D.Document)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .FirstOrDefault(li => li.GetLinkDocument() != null && li.GetLinkDocument().Equals(doc));

                    if (linkInstanceInHost != null)
                    {
                        var t = linkInstanceInHost.GetTotalTransform();
                        var pMin = t.OfPoint(mepBBox.Min);
                        var pMax = t.OfPoint(mepBBox.Max);
                        mepBBoxInHost = new BoundingBoxXYZ
                        {
                            Min = new XYZ(Math.Min(pMin.X, pMax.X), Math.Min(pMin.Y, pMax.Y), Math.Min(pMin.Z, pMax.Z)),
                            Max = new XYZ(Math.Max(pMin.X, pMax.X), Math.Max(pMin.Y, pMax.Y), Math.Max(pMin.Z, pMax.Z))
                        };

                        DebugLogger.Log($"[EfficientIntersectionService] DIAG converted MEP bbox into host coords: Min={FormatXYZ(mepBBoxInHost.Min)}, Max={FormatXYZ(mepBBoxInHost.Max)}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[EfficientIntersectionService] DIAG error converting MEP bbox to host coords: {ex.Message}");
            }

            // Get walls from host document (use the view's document when providing a view id)
            var hostWalls = new FilteredElementCollector(view3D.Document, view3D.Id)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .Where(w => IsElementInSectionBoxAndIntersectsMep(w, Transform.Identity, sectionBox, mepBBoxInHost))
                .Select(w => (w, Transform.Identity))
                .ToList();

            filteredWalls.AddRange(hostWalls);

            // Get walls from visible linked documents
            // NOTE: collect RevitLinkInstance objects from the host view's document so we enumerate
            // actual link instances placed in the host model (the previous code accidentally
            // collected from the linked document itself which yields no host link instances).
            foreach (var linkInstance in new FilteredElementCollector(view3D.Document)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null))
            {
                var linkDoc = linkInstance.GetLinkDocument();
                var linkTransform = linkInstance.GetTotalTransform();

                // Diagnostic: log link instance info and transform origin
                try
                {
                    var linkTitle = linkDoc != null ? linkDoc.Title : "<no-link-doc>";
                    var origin = linkTransform != null ? FormatXYZ(linkTransform.Origin) : "<identity>";
                    DebugLogger.Log($"[EfficientIntersectionService] DIAG LinkInstance Id={linkInstance.Id.IntegerValue}, LinkDoc='{linkTitle}', TransformOrigin={origin}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[EfficientIntersectionService] DIAG LinkInstance logging error: {ex.Message}");
                }

                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Where(w => IsElementInSectionBoxAndIntersectsMep(w, linkTransform, sectionBox, mepBBoxInHost))
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
                DebugLogger.Log($"[EfficientIntersectionService] Solid intersection error for element {structuralElement.Id.IntegerValue}: {ex.Message}");
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
