#nullable enable
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class MepIntersectionService
    {
        // Geometry cache to avoid re-processing same elements
        private static readonly Dictionary<string, Solid?> _geometryCache = new Dictionary<string, Solid?>();
        
        // Clear cache method for memory management
        public static void ClearGeometryCache()
        {
            _geometryCache.Clear();
        }
        
        // Main method to find intersections for a given MEP element - OPTIMIZED
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            var locationCurve = mepElement.Location as LocationCurve;
            if (locationCurve == null)
            {
                log($"ERROR: Could not get LocationCurve from element {mepElement.Id}.");
                return results;
            }

            var line = locationCurve.Curve as Line;
            if (line == null)
            {
                log($"ERROR: LocationCurve is not a Line for element {mepElement.Id}.");
                return results;
            }

            // Get MEP element bounding box for spatial pre-filtering
            var mepBBox = mepElement.get_BoundingBox(null);
            if (mepBBox == null)
            {
                log($"WARNING: Could not get bounding box for MEP element {mepElement.Id}.");
                return results;
            }

            // Expand MEP bounding box slightly for intersection tolerance
            const double tolerance = 1.0; // 1 foot tolerance
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);

            int processedCount = 0;
            int spatiallyFilteredCount = 0;

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;
                processedCount++;
                
                try
                {
                    // SPATIAL PRE-FILTERING: Check bounding box intersection first
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox != null)
                    {
                        // Transform structural bbox if it's from a linked doc
                        if (linkTransform != null)
                        {
                            var transformedMin = linkTransform.OfPoint(structBBox.Min);
                            var transformedMax = linkTransform.OfPoint(structBBox.Max);
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                                Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                            };
                        }
                        
                        // Quick bounding box intersection test
                        if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            continue; // Skip expensive geometry processing
                        }
                    }

                    // Get geometry from cache or compute it
                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    Solid? solid;
                    
                    if (!_geometryCache.TryGetValue(cacheKey, out solid))
                    {
                        var options = new Options();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;

                        solid = GetSolidFromGeometry(geometry);
                        if (solid != null && linkTransform != null)
                        {
                            solid = SolidUtils.CreateTransformed(solid, linkTransform);
                        }
                        
                        // Cache the result (even if null)
                        _geometryCache[cacheKey] = solid;
                    }
                    
                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, line, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        var center = GetBoundingBoxCenter(bbox);
                        results.Add((structuralElement, bbox, center));
                    }
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }
            
            log($"Spatial filtering: processed {processedCount}, skipped {spatiallyFilteredCount} elements via bounding box check");
            return results;
        }

        // Overload: accept a host-space Line (e.g. when the MEP element comes from a linked doc
        // and the caller has already transformed its curve into the active document coords).
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Line hostLine,
            BoundingBoxXYZ? mepBoundingBox,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            if (hostLine == null)
            {
                log("ERROR: hostLine is null in FindIntersections overload.");
                return results;
            }

            // Derive a MEP bounding box from provided bbox or from the line
            BoundingBoxXYZ mepBBox = mepBoundingBox ?? new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(hostLine.GetEndPoint(0).X, hostLine.GetEndPoint(1).X), Math.Min(hostLine.GetEndPoint(0).Y, hostLine.GetEndPoint(1).Y), Math.Min(hostLine.GetEndPoint(0).Z, hostLine.GetEndPoint(1).Z)),
                Max = new XYZ(Math.Max(hostLine.GetEndPoint(0).X, hostLine.GetEndPoint(1).X), Math.Max(hostLine.GetEndPoint(0).Y, hostLine.GetEndPoint(1).Y), Math.Max(hostLine.GetEndPoint(0).Z, hostLine.GetEndPoint(1).Z))
            };

            // Expand bbox slightly for tolerance (same approach as existing method)
            const double tolerance = 1.0;
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);

            int processedCount = 0;
            int spatiallyFilteredCount = 0;

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;
                processedCount++;
                try
                {
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox != null)
                    {
                        if (linkTransform != null)
                        {
                            var transformedMin = linkTransform.OfPoint(structBBox.Min);
                            var transformedMax = linkTransform.OfPoint(structBBox.Max);
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                                Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                            };
                        }
                        if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            continue;
                        }
                    }

                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    Solid? solid;
                    if (!_geometryCache.TryGetValue(cacheKey, out solid))
                    {
                        var options = new Options();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;
                        solid = GetSolidFromGeometry(geometry);
                        if (solid != null && linkTransform != null)
                            solid = SolidUtils.CreateTransformed(solid, linkTransform);
                        _geometryCache[cacheKey] = solid;
                    }
                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, hostLine);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        var center = GetBoundingBoxCenter(bbox);
                        results.Add((structuralElement, bbox, center));
                    }
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }

            log($"Spatial filtering: processed {processedCount}, skipped {spatiallyFilteredCount} elements via bounding box check");
            return results;
        }
        
        // Fast bounding box intersection test
        private static bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            return !(max1.X < min2.X || min1.X > max2.X ||
                     max1.Y < min2.Y || min1.Y > max2.Y ||
                     max1.Z < min2.Z || min1.Z > max2.Z);
        }

        // Extracts a solid from a geometry object
        private static Solid? GetSolidFromGeometry(GeometryElement geometry)
        {
            foreach (GeometryObject geomObj in geometry)
            {
                if (geomObj is Solid s && s.Volume > 0) return s;
                if (geomObj is GeometryInstance gi)
                {
                    foreach (GeometryObject instObj in gi.GetInstanceGeometry())
                    {
                        if (instObj is Solid s2 && s2.Volume > 0) return s2;
                    }
                }
            }
            return null;
        }

        // Intersects a solid with a line and returns the intersection points
        private static List<XYZ> GetIntersectionPoints(Solid solid, Line line, Action<string>? log = null)
        {
            var intersectionPoints = new List<XYZ>();
            try
            {
                int faceCount = solid.Faces.Size;
                log?.Invoke($"[Intersect] Solid face count = {faceCount}");
                foreach (Face face in solid.Faces)
                {
                    if (face == null) continue;
                    IntersectionResultArray? ira;
                    var res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        foreach (IntersectionResult ir in ira)
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                        if (intersectionPoints.Count > 0)
                        {
                            log?.Invoke($"[Intersect] Found {intersectionPoints.Count} intersection point(s). First: {intersectionPoints[0]}");
                            // early exit optional? keep collecting for bbox
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[Intersect] Exception while computing intersections: {ex.Message}");
            }
            return intersectionPoints;
        }

        // Creates a bounding box from a list of points
        private static BoundingBoxXYZ CreateBoundingBox(List<XYZ> points)
        {
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            foreach (var pt in points)
            {
                if (pt.X < minX) minX = pt.X;
                if (pt.Y < minY) minY = pt.Y;
                if (pt.Z < minZ) minZ = pt.Z;
                if (pt.X > maxX) maxX = pt.X;
                if (pt.Y > maxY) maxY = pt.Y;
                if (pt.Z > maxZ) maxZ = pt.Z;
            }
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        // Gets the center of a bounding box
        private static XYZ GetBoundingBoxCenter(BoundingBoxXYZ bbox)
        {
            return new XYZ((bbox.Min.X + bbox.Max.X) / 2, (bbox.Min.Y + bbox.Max.Y) / 2, (bbox.Min.Z + bbox.Max.Z) / 2);
        }

        // Collects structural elements within section box bounds only - MAJOR PERFORMANCE OPTIMIZATION
    public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc, Action<string> log)
        {
            var elements = new List<(Element, Transform?)>();
            log("Starting structural element collection.");

            // Get section box to drastically reduce search space
            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (doc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBox = view3D.GetSectionBox();
                    log($"Active view has a section box. Min: {sectionBox.Min}, Max: {sectionBox.Max}");
                }
                else
                {
                    log("No active section box found.");
                }
            }
            catch (Exception ex)
            {
                log($"Error getting section box: {ex.Message}");
            }

            // Use the same solid-based section-box filtering used by the MEP collector.
            // This mirrors the working logic and avoids AABB pitfalls across host/link transforms.
            var categories = new[] {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };

            try
            {
                var filtered = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectElementsVisibleOnly(doc, categories);
                foreach (var t in filtered) elements.Add((t.element, t.transform));
                log($"Finished structural element collection. Total elements found: {elements.Count}");
                
                // CRITICAL FIX: If section box filtering is too aggressive and filters out ALL elements,
                // fall back to collecting structural elements without section box filtering
                if (elements.Count == 0 && sectionBox != null)
                {
                    log("WARNING: Section box filtering removed ALL structural elements. Falling back to unfiltered collection.");
                    
                    // Collect structural elements without section box filtering as fallback
                    var fallbackElements = new List<(Element, Transform?)>();
                    
                    // Host model elements
                    var hostElements = new FilteredElementCollector(doc)
                        .WherePasses(new ElementMulticategoryFilter(categories))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    foreach (var e in hostElements) fallbackElements.Add((e, null));
                    
                    // Linked model elements
                    foreach (var link in new FilteredElementCollector(doc)
                                 .OfClass(typeof(RevitLinkInstance))
                                 .Cast<RevitLinkInstance>())
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null ||
                            doc.ActiveView.GetCategoryHidden(link.Category.Id) ||
                            link.IsHidden(doc.ActiveView))
                            continue;

                        var tr = link.GetTotalTransform();
                        var linked = new FilteredElementCollector(linkDoc)
                            .WherePasses(new ElementMulticategoryFilter(categories))
                            .WhereElementIsNotElementType()
                            .ToElements();
                        foreach (var e in linked) fallbackElements.Add((e, tr));
                    }
                    
                    log($"Fallback collection found {fallbackElements.Count} structural elements.");
                    return fallbackElements;
                }
                
                return elements;
            }
            catch (Exception ex)
            {
                log($"ERROR: Fallback structural collection failed: {ex.Message}");
                return elements;
            }
        }

        // Backwards-compatible overload: no-op logger
        public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc)
        {
            return CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, _ => { });
        }
    }
}