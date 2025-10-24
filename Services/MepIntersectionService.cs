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
        
        // PHASE 1 OPTIMIZATION 2: Category Whitelist (2x speedup)
        private static readonly BuiltInCategory[] MEP_CATEGORY_WHITELIST = {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctAccessory,  // Includes dampers
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting
        };
        
        private static readonly BuiltInCategory[] STRUCTURAL_CATEGORY_WHITELIST = {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFoundation
        };
        
        // Clear cache method for memory management
        public static void ClearGeometryCache()
        {
            _geometryCache.Clear();
        }
        
        // PHASE 1 OPTIMIZATION 2: Category whitelist filtering methods
        public static bool IsMepCategoryWhitelisted(Element element)
        {
            if (element.Category?.Id?.IntegerValue == null) return false;
            
            var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
            return MEP_CATEGORY_WHITELIST.Contains(categoryId);
        }
        
        public static bool IsStructuralCategoryWhitelisted(Element element)
        {
            if (element.Category?.Id?.IntegerValue == null) return false;
            
            var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
            return STRUCTURAL_CATEGORY_WHITELIST.Contains(categoryId);
        }
        
        public static bool IsDamperElement(Element element)
        {
            if (element is FamilyInstance fi)
            {
                var familyName = fi.Symbol?.Family?.Name?.ToLower() ?? "";
                return familyName.Contains("damper");
            }
            return false;
        }
        
        // BATCH PROCESSING: Find intersections for multiple MEP elements efficiently
        public static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            log($"[BatchIntersection] Processing {mepElements.Count} MEP elements against {structuralElements.Count} structural elements");

            // Pre-compute structural element bounding boxes and geometry once
            var structuralData = new List<(Element element, Transform? transform, BoundingBoxXYZ bbox, Solid? solid)>();
            
            foreach (var (structElement, structTransform) in structuralElements)
            {
                var structBBox = structElement.get_BoundingBox(null);
                if (structBBox == null) continue;

                // Transform structural bbox to host shared coordinates
                if (structTransform != null)
                {
                    var transformedMin = structTransform.OfPoint(structBBox.Min);
                    var transformedMax = structTransform.OfPoint(structBBox.Max);
                    structBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }

                // Pre-compute and cache geometry
                string cacheKey = $"{structElement.Id.IntegerValue}_{structTransform?.GetHashCode() ?? 0}";
                Solid? solid = null;
                
                if (!_geometryCache.TryGetValue(cacheKey, out solid))
                {
                    var options = new Options();
                    var geometry = structElement.get_Geometry(options);
                    if (geometry != null)
                    {
                        solid = GetSolidFromGeometry(geometry);
                        if (solid != null && structTransform != null)
                        {
                            solid = SolidUtils.CreateTransformed(solid, structTransform);
                        }
                    }
                    _geometryCache[cacheKey] = solid;
                }

                structuralData.Add((structElement, structTransform, structBBox, solid));
            }

            log($"[BatchIntersection] Pre-computed {structuralData.Count} structural elements with geometry");

            // Process each MEP element against pre-computed structural data
            foreach (var (mepElement, mepTransform) in mepElements)
            {
                var mepBBox = mepElement.get_BoundingBox(null);
                if (mepBBox == null) continue;

                // Transform MEP bbox to host shared coordinates
                if (mepTransform != null)
                {
                    var transformedMin = mepTransform.OfPoint(mepBBox.Min);
                    var transformedMax = mepTransform.OfPoint(mepBBox.Max);
                    mepBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }

                var line = GetElementLine(mepElement, mepBBox, log);
                if (line == null)
                {
                    // Handle damper-style elements
                    var damperResults = FindDamperIntersectionsInternal(mepElement, mepBBox, 
                        structuralData.Select(sd => (sd.element, sd.transform)).ToList(), null, log);
                    results.AddRange(damperResults.Select(i => (mepElement, i.Item1, i.Item2, i.Item3)));
                    continue;
                }

                // Quick spatial pre-filtering with tolerance
                // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
                const double tolerance = 1.0; // 1.0ft tolerance - original working value
                var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
                var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);

                int spatiallyFiltered = 0;
                foreach (var (structElement, structTransform, structBBox, solid) in structuralData)
                {
                    // Quick bounding box intersection test
                    if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                    {
                        spatiallyFiltered++;
                        continue;
                    }

                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, line, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
                        var center = GetBoundingBoxCenter(bbox);
                        results.Add((mepElement, structElement, bbox, center));
                    }
                }

                if (spatiallyFiltered > 0)
                {
                    log($"[BatchIntersection] MEP {mepElement.Id}: spatially filtered {spatiallyFiltered}/{structuralData.Count} structural elements");
                }
            }

            log($"[BatchIntersection] Found {results.Count} total intersections");
            return results;
        }

        // Individual method for backwards compatibility - now delegates to batch processing
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            // Delegate to batch processing with single element
            var batchResults = FindIntersectionsBatch(
                new List<(Element, Transform?)> { (mepElement, null) },
                structuralElements,
                log);
            
            // Convert batch results to individual format
            return batchResults.Select(r => (r.Item2, r.Item3, r.Item4)).ToList();
        }

        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Element mepElement,
            Transform? mepTransform,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            // Delegate to batch processing with single element and transform
            var batchResults = FindIntersectionsBatch(
                new List<(Element, Transform?)> { (mepElement, mepTransform) },
                structuralElements,
                log);
            
            // Convert batch results to individual format
            return batchResults.Select(r => (r.Item2, r.Item3, r.Item4)).ToList();
        }

        // Legacy method - kept for compatibility
        private static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersectionsLegacy(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            var mepBBox = mepElement.get_BoundingBox(null);
            if (mepBBox == null)
            {
                log($"WARNING: Could not get bounding box for MEP element {mepElement.Id}.");
                return results;
            }

            var line = GetElementLine(mepElement, mepBBox, log);
            if (line == null)
            {
                log($"INFO: Falling back to damper-style processing for element {mepElement.Id}.");
                results.AddRange(FindDamperIntersectionsInternal(mepElement, mepBBox, structuralElements, null, log));
                return results;
            }

            // REVERTED: Back to original 1.0 foot tolerance since coordinate transform issue is fixed
            // Both MEP and wall bounding boxes are now in host shared coordinates
            // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
            const double tolerance = 1.0; // 1.0ft tolerance - original working value
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

            int processedCount = 0;
            int spatiallyFilteredCount = 0;
            int lastLoggedSkipCount = 0;

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
                            if (spatiallyFilteredCount - lastLoggedSkipCount >= 100)
                            {
                                lastLoggedSkipCount = spatiallyFilteredCount;
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
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
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
            // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
            const double tolerance = 1.0; // 1.0ft tolerance - original working value
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering (overload method)");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

            int processedCount = 0;
            int spatiallyFilteredCount = 0;
            int lastLoggedSkipCount = 0;

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
                        if (linkTransform != null)
                        {
                            // Apply 8-corner transform to wall bounding box
                            
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
                            
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Min(p => p.Z)),
                                Max = new XYZ(pts.Max(p => p.X), pts.Max(p => p.Y), pts.Max(p => p.Z))
                            };
                        }
                        
                        if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            if (spatiallyFilteredCount - lastLoggedSkipCount >= 100)
                            {
                                lastLoggedSkipCount = spatiallyFilteredCount;
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

                    var intersectionPoints = GetIntersectionPoints(solid, hostLine, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
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
    public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc, Action<string> log, List<string>? selectedHostTypes = null)
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

            // ✅ FIX: Filter structural categories based on UI host type selection
            var allCategories = new[] {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };

            var categories = allCategories;
            if (selectedHostTypes != null && selectedHostTypes.Count > 0)
            {
                var filteredCategories = new List<BuiltInCategory>();
                
                if (selectedHostTypes.Any(ht => ht.Equals("Walls", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_Walls);
                    log("UI Selection: Including Walls");
                }
                
                if (selectedHostTypes.Any(ht => ht.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_StructuralFraming);
                    log("UI Selection: Including Structural Framing");
                }
                
                if (selectedHostTypes.Any(ht => ht.Equals("Floors", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_Floors);
                    log("UI Selection: Including Floors");
                }
                
                if (filteredCategories.Count > 0)
                {
                    categories = filteredCategories.ToArray();
                    log($"✅ FILTERED: Only collecting {filteredCategories.Count} selected host types (was {allCategories.Length} total)");
                }
                else
                {
                    log("⚠️ WARNING: No valid host types selected, using all categories");
                }
            }
            else
            {
                log("No host type selection provided, collecting all structural categories");
            }

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
            return CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, _ => { }, null);
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

        public static List<(Element, BoundingBoxXYZ, XYZ)> FindDamperIntersections(
            Element damperElement,
            List<(Element, Transform?)> structuralElements,
            Transform? damperLinkTransform,
            Action<string> log)
        {
            var damperBBox = damperElement.get_BoundingBox(null);
            if (damperBBox == null)
            {
                log($"WARNING: Could not get bounding box for damper element {damperElement.Id}.");
                return new List<(Element, BoundingBoxXYZ, XYZ)>();
            }

            return FindDamperIntersectionsInternal(damperElement, damperBBox, structuralElements, damperLinkTransform, log);
        }

        private static List<(Element, BoundingBoxXYZ, XYZ)> FindDamperIntersectionsInternal(
            Element damperElement,
            BoundingBoxXYZ damperBBox,
            List<(Element, Transform?)> structuralElements,
            Transform? damperLinkTransform,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();

            BoundingBoxXYZ hostDamperBBox = damperBBox;
            if (damperLinkTransform != null)
            {
                var transformed = TransformBoundingBox(damperBBox, damperLinkTransform);
                if (transformed != null)
                {
                    hostDamperBBox = transformed;
                }
            }

            const double tolerance = 0.5; // 6 inches
            var expandedMin = new XYZ(
                hostDamperBBox.Min.X - tolerance,
                hostDamperBBox.Min.Y - tolerance,
                hostDamperBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(
                hostDamperBBox.Max.X + tolerance,
                hostDamperBBox.Max.Y + tolerance,
                hostDamperBBox.Max.Z + tolerance);

            log($"[DamperIntersection] Processing element {damperElement.Id} with bbox Min=({hostDamperBBox.Min.X:F2}, {hostDamperBBox.Min.Y:F2}, {hostDamperBBox.Min.Z:F2}) Max=({hostDamperBBox.Max.X:F2}, {hostDamperBBox.Max.Y:F2}, {hostDamperBBox.Max.Z:F2})");

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;

                try
                {
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox == null) continue;

                    if (linkTransform != null)
                    {
                        var transformedStruct = TransformBoundingBox(structBBox, linkTransform);
                        if (transformedStruct != null)
                        {
                            structBBox = transformedStruct;
                        }
                    }

                    if (!BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        continue;

                    log($"[DamperIntersection] Intersection candidate: damper {damperElement.Id} with structural {structuralElement.Id}");

                    var intersectionMin = new XYZ(
                        Math.Max(hostDamperBBox.Min.X, structBBox.Min.X),
                        Math.Max(hostDamperBBox.Min.Y, structBBox.Min.Y),
                        Math.Max(hostDamperBBox.Min.Z, structBBox.Min.Z));
                    var intersectionMax = new XYZ(
                        Math.Min(hostDamperBBox.Max.X, structBBox.Max.X),
                        Math.Min(hostDamperBBox.Max.Y, structBBox.Max.Y),
                        Math.Min(hostDamperBBox.Max.Z, structBBox.Max.Z));

                    if (intersectionMin.X > intersectionMax.X ||
                        intersectionMin.Y > intersectionMax.Y ||
                        intersectionMin.Z > intersectionMax.Z)
                    {
                        continue;
                    }

                    var intersectionBBox = new BoundingBoxXYZ
                    {
                        Min = intersectionMin,
                        Max = intersectionMax
                    };

                    var center = GetBoundingBoxCenter(intersectionBBox);
                    results.Add((structuralElement, intersectionBBox, center));
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process damper intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }

            return results;
        }

        private static BoundingBoxXYZ? TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            try
            {
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

                var newMin = new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Min(p => p.Z));
                var newMax = new XYZ(pts.Max(p => p.X), pts.Max(p => p.Y), pts.Max(p => p.Z));

                return new BoundingBoxXYZ
                {
                    Min = newMin,
                    Max = newMax
                };
            }
            catch
            {
                return null;
            }
        }

        private static Line? GetElementLine(Element element, BoundingBoxXYZ mepBBox, Action<string> log)
        {
            if (element is FamilyInstance fi && fi.Symbol?.Family?.Name?.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                log($"[MepIntersectionService] Element {element.Id} identified as damper; using bounding-box intersection approach.");
                return null;
            }

            if (element.Location is LocationCurve locCurve && locCurve.Curve is Line curveLine)
            {
                return curveLine;
            }

            if (element is MEPCurve mepCurve)
            {
                try
                {
                    var connectors = mepCurve.ConnectorManager?.Connectors?.Cast<Connector>().Where(c => c != null).ToList();
                    if (connectors != null && connectors.Count >= 2)
                    {
                        var endpoints = connectors
                            .SelectMany((c, idx) => connectors
                                .Skip(idx + 1)
                                .Select(other => new { First = c, Second = other, Distance = c.Origin.DistanceTo(other.Origin) }))
                            .OrderByDescending(x => x.Distance)
                            .FirstOrDefault();

                        if (endpoints != null && endpoints.Distance > 0)
                        {
                            return Line.CreateBound(endpoints.First.Origin, endpoints.Second.Origin);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log($"[MepIntersectionService] Failed deriving line from MEPCurve connectors for element {element.Id}: {ex.Message}");
                }
            }

            try
            {
                var min = mepBBox.Min;
                var max = mepBBox.Max;
                var centerX = (min.X + max.X) * 0.5;
                var centerY = (min.Y + max.Y) * 0.5;
                var p1 = new XYZ(centerX, centerY, min.Z);
                var p2 = new XYZ(centerX, centerY, max.Z);

                if (p1.DistanceTo(p2) < 1e-6)
                {
                    p1 = new XYZ(min.X, centerY, (min.Z + max.Z) * 0.5);
                    p2 = new XYZ(max.X, centerY, (min.Z + max.Z) * 0.5);
                }

                if (p1.DistanceTo(p2) < 1e-6)
                {
                    p2 = new XYZ(p1.X + 1.0, p1.Y, p1.Z);
                }

                log($"[MepIntersectionService] Fallback line derived from bounding box for element {element.Id}.");
                return Line.CreateBound(p1, p2);
            }
            catch (Exception ex)
            {
                log($"[MepIntersectionService] Failed to derive fallback line for element {element.Id}: {ex.Message}");
                return null;
            }
        }
    }
}