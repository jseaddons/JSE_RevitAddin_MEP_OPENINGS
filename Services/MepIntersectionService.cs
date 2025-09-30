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

            // REVERTED: Back to original 1.0 foot tolerance since coordinate transform issue is fixed
            // Both MEP and wall bounding boxes are now in host shared coordinates
            const double tolerance = 1.0; // 1 foot tolerance - ORIGINAL
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

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
                        // CRITICAL FIX: Transform structural bbox to host shared coordinates for distance check
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
                        
                        // Quick bounding box intersection test (both in host shared coordinates now)
                        if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            // PERFORMANCE OPTIMIZATION: Reduce logging frequency
                            if (spatiallyFilteredCount % 100 == 0) // Log every 100th skip
                            {
                                var wallType = structuralElement.GetType().Name;
                                var wallId = structuralElement.Id.IntegerValue;
                                var distance = GetDistanceToMepElement(mepBBox, structBBox, null);
                                log($"[MepIntersectionService] SPATIAL FILTER: Skipped {spatiallyFilteredCount} elements so far. Latest: {wallType} ID:{wallId} - Distance: {distance:F2}ft");
                            }
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

            // REVERTED: Back to original 1.0 foot tolerance since coordinate transform issue is fixed
            // Both MEP and wall bounding boxes are now in host shared coordinates
            const double tolerance = 1.0; // 1 foot tolerance - ORIGINAL
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering (overload method)");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

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
                        // CRITICAL FIX: Transform structural bbox to host shared coordinates for distance check
                        log($"[MepIntersectionService] Wall {structuralElement.Id} bbox before transform: Min=({structBBox.Min.X:F2}, {structBBox.Min.Y:F2}, {structBBox.Min.Z:F2}) Max=({structBBox.Max.X:F2}, {structBBox.Max.Y:F2}, {structBBox.Max.Z:F2})");
                        log($"[MepIntersectionService] Wall {structuralElement.Id} linkTransform is {(linkTransform != null ? "NOT NULL" : "NULL")}");
                        if (linkTransform != null)
                        {
                            log($"[MepIntersectionService] Transform matrix: Origin=({linkTransform.Origin.X:F2}, {linkTransform.Origin.Y:F2}, {linkTransform.Origin.Z:F2})");
                            log($"[MepIntersectionService] Applying 8-corner transform to wall {structuralElement.Id}");
                            
                            // Transform all 8 corners of the bounding box to shared coordinates
                            var pts = new[]
                            {
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Min.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Min.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Max.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Min.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Max.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Max.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Min.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Max.Y, structBBox.Min.Z))
                            };
                            
                            var newMin = new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Min(p => p.Z));
                            var newMax = new XYZ(pts.Max(p => p.X), pts.Max(p => p.Y), pts.Max(p => p.Z));
                            
                            log($"[MepIntersectionService] 8-corner transform result: Min=({newMin.X:F2}, {newMin.Y:F2}, {newMin.Z:F2}) Max=({newMax.X:F2}, {newMax.Y:F2}, {newMax.Z:F2})");
                            
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = newMin,
                                Max = newMax
                            };
                            log($"[MepIntersectionService] Wall {structuralElement.Id} bbox after transform: Min=({structBBox.Min.X:F2}, {structBBox.Min.Y:F2}, {structBBox.Min.Z:F2}) Max=({structBBox.Max.X:F2}, {structBBox.Max.Y:F2}, {structBBox.Max.Z:F2})");
                        }
                        else
                        {
                            log($"[MepIntersectionService] No transform needed for wall {structuralElement.Id} (host element)");
                        }
                        if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            // PERFORMANCE OPTIMIZATION: Reduce logging frequency (overload method)
                            if (spatiallyFilteredCount % 100 == 0) // Log every 100th skip
                            {
                                var wallType = structuralElement.GetType().Name;
                                var wallId = structuralElement.Id.IntegerValue;
                                var distance = GetDistanceToMepElement(mepBBox, structBBox, null);
                                log($"[MepIntersectionService] SPATIAL FILTER: Skipped {spatiallyFilteredCount} elements so far. Latest: {wallType} ID:{wallId} - Distance: {distance:F2}ft");
                            }
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
                // Use the existing CollectElements method from the command
                // This is already debugged and working
                var mepElements = new List<Element>();
                var wallElements = new List<Element>();
                
                // Get section box bounds
                if (sectionBox != null)
                {
                    Transform sectionTransform = sectionBox.Transform;
                    List<XYZ> corners = new List<XYZ>
                    {
                        sectionTransform.OfPoint(sectionBox.Min),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Max.Z)),
                        sectionBox.Max
                    };
                    
                    XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                    XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));
                    
                    // Use the existing CollectElements method
                    CollectElements(doc, modelMin, modelMax, ref mepElements, ref wallElements);
                    
                    // Convert wall elements to the expected format
                    var links = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .ToList();
                    
                    foreach (var wall in wallElements)
                    {
                        var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                        Transform? wallTransform = wallLink?.GetTotalTransform();
                        elements.Add((wall, wallTransform));
                    }
                }
                else
                {
                    // Fallback: collect all structural elements without section box filtering
                    var hostElements = new FilteredElementCollector(doc)
                        .WherePasses(new ElementMulticategoryFilter(categories))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    foreach (var e in hostElements) elements.Add((e, null));
                    
                    // Linked model elements
                    foreach (var link in new FilteredElementCollector(doc)
                                 .OfClass(typeof(RevitLinkInstance))
                                 .Cast<RevitLinkInstance>())
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;

                        var tr = link.GetTotalTransform();
                        var linked = new FilteredElementCollector(linkDoc)
                            .WherePasses(new ElementMulticategoryFilter(categories))
                            .WhereElementIsNotElementType()
                            .ToElements();
                        foreach (var e in linked) elements.Add((e, tr));
                    }
                }
                
                log($"Finished structural element collection. Total elements found: {elements.Count}");
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
        
        // Copy of the working CollectElements method from the command
        private static void CollectElements(Document doc, XYZ modelMin, XYZ modelMax, 
                                         ref List<Element> mepElements, ref List<Element> wallElements)
        {
            Outline hostOutline = new Outline(modelMin, modelMax);

            // MEP categories
            BuiltInCategory[] mepCats = {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            // Collect from host
            foreach (var cat in mepCats)
            {
                mepElements.AddRange(
                    new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements()
                );
            }

            // CRITICAL FIX: Also collect damper family instances that might not be properly categorized
            // as OST_DuctAccessory but are still dampers based on family name
            var damperFamilyInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                .Where(fi => {
                    var bbox = fi.get_BoundingBox(null);
                    if (bbox == null) return false;
                    return BoundingBoxesIntersect(modelMin, modelMax, bbox.Min, bbox.Max);
                })
                .Cast<Element>()
                .ToList();
            
            mepElements.AddRange(damperFamilyInstances);

            wallElements.AddRange(
                new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                    .ToElements()
            );

            // Collect from links
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                Transform linkTransform = link.GetTotalTransform();
                Transform invTransform = linkTransform.Inverse;

                XYZ linkMin = invTransform.OfPoint(modelMin);
                XYZ linkMax = invTransform.OfPoint(modelMax);

                XYZ actualMin = new XYZ(
                    Math.Min(linkMin.X, linkMax.X),
                    Math.Min(linkMin.Y, linkMax.Y),
                    Math.Min(linkMin.Z, linkMax.Z)
                );
                XYZ actualMax = new XYZ(
                    Math.Max(linkMin.X, linkMax.X),
                    Math.Max(linkMin.Y, linkMax.Y),
                    Math.Max(linkMin.Z, linkMax.Z)
                );

                Outline linkOutline = new Outline(actualMin, actualMax);

                foreach (var cat in mepCats)
                {
                    mepElements.AddRange(
                        new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements()
                    );
                }

                // CRITICAL FIX: Also collect damper family instances from linked documents
                var linkedDamperFamilyInstances = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                    .Where(fi => {
                        var bbox = fi.get_BoundingBox(null);
                        if (bbox == null) return false;
                        // Transform the damper bbox to link coordinates for intersection test
                        var transformedMin = invTransform.OfPoint(bbox.Min);
                        var transformedMax = invTransform.OfPoint(bbox.Max);
                        var transformedBboxMin = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z));
                        var transformedBboxMax = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z));
                        return BoundingBoxesIntersect(modelMin, modelMax, transformedBboxMin, transformedBboxMax);
                    })
                    .Cast<Element>()
                    .ToList();
                
                mepElements.AddRange(linkedDamperFamilyInstances);

                wallElements.AddRange(
                    new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                        .ToElements()
                );
            }
        }
        
        /// <summary>
        /// Collects damper family instances from host and linked documents for intersection detection
        /// </summary>
        public static List<(Element, Transform?)> CollectDamperElementsForIntersection(Document doc, Action<string> log)
        {
            var elements = new List<(Element, Transform?)>();
            log("Starting damper element collection for intersection detection.");

            // Get section box to filter by visible area
            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (doc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBox = view3D.GetSectionBox();
                    log($"Active view has a section box for damper filtering. Min: {sectionBox.Min}, Max: {sectionBox.Max}");
                }
                else
                {
                    log("No active section box found for damper collection.");
                }
            }
            catch (Exception ex)
            {
                log($"Error getting section box for dampers: {ex.Message}");
            }

            // Collect dampers from host document
            try
            {
                var hostDampers = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                    .ToList();

                foreach (var damper in hostDampers)
                {
                    // Apply section box filtering if available
                    if (sectionBox != null)
                    {
                        var bbox = damper.get_BoundingBox(null);
                        if (bbox != null && !BoundingBoxesIntersect(sectionBox.Min, sectionBox.Max, bbox.Min, bbox.Max))
                            continue;
                    }
                    elements.Add((damper, null));
                }

                log($"Found {hostDampers.Count} damper family instances in host document.");
            }
            catch (Exception ex)
            {
                log($"Error collecting host dampers: {ex.Message}");
            }

            // Collect dampers from linked documents
            try
            {
                var links = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var link in links)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc == null) continue;

                    var linkTransform = link.GetTotalTransform();
                    var linkedDampers = new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                        .ToList();

                    foreach (var damper in linkedDampers)
                    {
                        // Apply section box filtering if available (transform coordinates)
                        if (sectionBox != null)
                        {
                            var bbox = damper.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                var transformedMin = linkTransform.OfPoint(bbox.Min);
                                var transformedMax = linkTransform.OfPoint(bbox.Max);
                                var transformedBboxMin = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z));
                                var transformedBboxMax = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z));
                                
                                if (!BoundingBoxesIntersect(sectionBox.Min, sectionBox.Max, transformedBboxMin, transformedBboxMax))
                                    continue;
                            }
                        }
                        elements.Add((damper, linkTransform));
                    }

                    log($"Found {linkedDampers.Count} damper family instances in linked document: {linkDoc.Title}");
                }
            }
            catch (Exception ex)
            {
                log($"Error collecting linked dampers: {ex.Message}");
            }

            log($"Finished damper element collection. Total dampers found: {elements.Count}");
            return elements;
        }

        // Helper method to calculate distance between MEP element and structural element
        private static double GetDistanceToMepElement(BoundingBoxXYZ mepBBox, BoundingBoxXYZ structBBox, Transform? linkTransform)
        {
            try
            {
                // Transform structural bbox if it's from a linked doc
                BoundingBoxXYZ transformedStructBBox = structBBox;
                if (linkTransform != null)
                {
                    var transformedMin = linkTransform.OfPoint(structBBox.Min);
                    var transformedMax = linkTransform.OfPoint(structBBox.Max);
                    transformedStructBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }
                
                // Calculate center points
                var mepCenter = new XYZ(
                    (mepBBox.Min.X + mepBBox.Max.X) / 2,
                    (mepBBox.Min.Y + mepBBox.Max.Y) / 2,
                    (mepBBox.Min.Z + mepBBox.Max.Z) / 2
                );
                
                var structCenter = new XYZ(
                    (transformedStructBBox.Min.X + transformedStructBBox.Max.X) / 2,
                    (transformedStructBBox.Min.Y + transformedStructBBox.Max.Y) / 2,
                    (transformedStructBBox.Min.Z + transformedStructBBox.Max.Z) / 2
                );
                
                // Calculate distance in feet
                var distance = mepCenter.DistanceTo(structCenter);
                return UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Feet);
            }
            catch
            {
                return -1.0; // Return -1 if calculation fails
            }
        }
    }
}