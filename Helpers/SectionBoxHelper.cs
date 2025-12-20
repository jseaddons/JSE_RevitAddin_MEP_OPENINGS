using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class SectionBoxHelper
    {
        /// <summary>
        /// Filters a combined list of host and linked elements against the active 3D view's section box.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument.</param>
        /// <param name="elementsWithTransforms">The list of elements to filter, containing both host (null transform) and linked elements.</param>
        /// <returns>A filtered list of tuples containing only the elements that intersect the section box.</returns>
        public static List<(Element element, Transform? transform)> FilterElementsBySectionBox(
            UIDocument uiDoc,
            List<(Element element, Transform? transform)> elementsWithTransforms)
        {
            if (!(uiDoc.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
            {
                // If there's no active section box, return the original unfiltered list.
                return elementsWithTransforms;
            }

            Solid? sectionBoxSolid = GetSectionBoxAsSolid(view3D);
            if (sectionBoxSolid == null || sectionBoxSolid.Volume <= 0)
            {
                return elementsWithTransforms; // Return original list if solid is invalid
            }

            var filteredList = new List<(Element element, Transform? transform)>();

            // Separate host and linked elements for efficient filtering
            var hostElements = elementsWithTransforms.Where(t => t.transform == null).Select(t => t.element).ToList();
            var linkedElementGroups = elementsWithTransforms.Where(t => t.transform != null).GroupBy(t => t.element.Document.Title);

            // Filter host elements
            if (hostElements.Any())
            {
                ElementFilter hostFilter;
                var hostIds = hostElements.Select(e => e.Id).ToList();
                
                // ✅ PRIORITY 1 OPTIMIZATION: Use BoundingBoxIntersectsFilter instead of ElementIntersectsSolidFilter
                // BoundingBoxIntersectsFilter is 20-30% faster (no geometry extraction required)
                if (OptimizationFlags.UseBoundingBoxSectionBoxFilter)
                {
                    // Fast path: Use bounding box filter (outline-based)
                    var sectionBoxBounds = GetSectionBoxBounds(view3D);
                    if (sectionBoxBounds != null)
                    {
                        var sectionBoxOutline = new Outline(sectionBoxBounds.Min, sectionBoxBounds.Max);
                        hostFilter = new BoundingBoxIntersectsFilter(sectionBoxOutline);
                    }
                    else
                    {
                        // Fallback: If bounds extraction fails, use solid filter
                        hostFilter = new ElementIntersectsSolidFilter(sectionBoxSolid);
                    }
                }
                else
                {
                    // Fallback: Use existing solid filter (slower but more precise)
                    hostFilter = new ElementIntersectsSolidFilter(sectionBoxSolid);
                }
                
                DebugLogger.Info($"[SectionBoxDiag] Host elements count={hostElements.Count}, sampleIds={string.Join(",", hostIds.Take(6).Select(id => id.IntegerValue.ToString()))}");
                var passingHostIds = new FilteredElementCollector(uiDoc.Document, hostIds)
                    .WherePasses(hostFilter)
                    .ToElementIds();
                
                DebugLogger.Info($"[SectionBoxDiag] ElementIntersectsSolidFilter returned {passingHostIds.Count} passing elements");
                
                // ✅ CRITICAL FIX: Manual fallback if filter fails
                // ElementIntersectsSolidFilter sometimes fails for family instances
                // Use manual bounding box intersection check as fallback
                if (passingHostIds.Count == 0 && hostElements.Count > 0)
                {
                    DebugLogger.Warning($"[SectionBoxDiag] ElementIntersectsSolidFilter returned 0 results, using manual BBox check fallback");
                    var sectionBoxBounds = GetSectionBoxBounds(view3D);
                    if (sectionBoxBounds != null)
                    {
                        var manualPassingIds = new List<ElementId>();
                        foreach (var elem in hostElements)
                        {
                            var elemBBox = elem.get_BoundingBox(null);
                            if (elemBBox != null && BoundingBoxesIntersect(
                                sectionBoxBounds.Min, sectionBoxBounds.Max,
                                elemBBox.Min, elemBBox.Max))
                            {
                                manualPassingIds.Add(elem.Id);
                                DebugLogger.Info($"[SectionBoxDiag] ✅ Manual check: Element {elem.Id.IntegerValue} PASSES (BBox intersects)");
                            }
                            else
                            {
                                DebugLogger.Info($"[SectionBoxDiag] ❌ Manual check: Element {elem.Id.IntegerValue} FAILS (BBox does not intersect)");
                            }
                        }
                        passingHostIds = manualPassingIds;
                        DebugLogger.Info($"[SectionBoxDiag] Manual BBox check found {passingHostIds.Count} passing elements");
                    }
                }
                
                filteredList.AddRange(hostElements.Where(e => passingHostIds.Contains(e.Id)).Select(e => (e, (Transform?)null)));
                try
                {
                    DebugLogger.Info($"[SectionBoxDiag] Host section-solid volume={sectionBoxSolid.Volume}, passingHostCount={passingHostIds.Count}");
                    foreach (var e in hostElements.Where(e => passingHostIds.Contains(e.Id)).Take(3))
                    {
                        var bbox = e.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            DebugLogger.Info($"[SectionBoxDiag] Host Element {e.Id} BBox Min({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}) Max({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})");
                        }
                    }
                }
                catch { }

                // No fallback for host elements - if no elements pass solid filter, return empty
                DebugLogger.Info($"[SectionBoxDiag] Host elements: manual fallback implemented");
            }

            // Filter linked elements
            var allLinkInstances = new FilteredElementCollector(uiDoc.Document)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var group in linkedElementGroups)
            {
                var linkInstance = allLinkInstances.FirstOrDefault(li => li.GetLinkDocument()?.Title == group.Key);
                if (linkInstance == null) continue;

                Transform inverseTransform = linkInstance.GetTotalTransform().Inverse;
                Solid transformedSolid = SolidUtils.CreateTransformed(sectionBoxSolid, inverseTransform);

                try
                {
                    // DebugLogger.Info($"[SectionBoxDiag] Processing link='{group.Key}', TransformOrigin={linkInstance.GetTotalTransform().Origin}, TransformedSolidVol={transformedSolid.Volume}");
                }
                catch { }

                var elementsInLink = group.Select(t => t.element).ToList();
                if (elementsInLink.Any())
                {
                    ElementFilter linkFilter;
                    
                    // ✅ PRIORITY 1 OPTIMIZATION: Use BoundingBoxIntersectsFilter instead of ElementIntersectsSolidFilter
                    // BoundingBoxIntersectsFilter is 20-30% faster (no geometry extraction required)
                    if (OptimizationFlags.UseBoundingBoxSectionBoxFilter)
                    {
                        // Fast path: Use bounding box filter (outline-based)
                        // Transform section box bounds to link coordinates
                        var sectionBoxBounds = GetSectionBoxBounds(view3D);
                        if (sectionBoxBounds != null)
                        {
                            // Transform bounds to link coordinates
                            var transformedMin = inverseTransform.OfPoint(sectionBoxBounds.Min);
                            var transformedMax = inverseTransform.OfPoint(sectionBoxBounds.Max);
                            
                            // Ensure min < max after transformation
                            var linkMin = new XYZ(
                                System.Math.Min(transformedMin.X, transformedMax.X),
                                System.Math.Min(transformedMin.Y, transformedMax.Y),
                                System.Math.Min(transformedMin.Z, transformedMax.Z));
                            var linkMax = new XYZ(
                                System.Math.Max(transformedMin.X, transformedMax.X),
                                System.Math.Max(transformedMin.Y, transformedMax.Y),
                                System.Math.Max(transformedMin.Z, transformedMax.Z));
                            
                            var linkOutline = new Outline(linkMin, linkMax);
                            linkFilter = new BoundingBoxIntersectsFilter(linkOutline);
                        }
                        else
                        {
                            // Fallback: If bounds extraction fails, use solid filter
                            linkFilter = new ElementIntersectsSolidFilter(transformedSolid);
                        }
                    }
                    else
                    {
                        // Fallback: Use existing solid filter (slower but more precise)
                        linkFilter = new ElementIntersectsSolidFilter(transformedSolid);
                    }
                    
                    var elementIds = new List<ElementId>(elementsInLink.Select(e => e.Id));
                    var passingLinkIds = new FilteredElementCollector(linkInstance.GetLinkDocument(), elementIds)
                        .WherePasses(linkFilter)
                        .ToElementIds();

                    try
                    {
                        // DebugLogger.Info($"[SectionBoxDiag] link='{group.Key}' passingCount={passingLinkIds.Count}");
                        foreach (var id in passingLinkIds.Take(3))
                        {
                            var e = elementsInLink.FirstOrDefault(el => el.Id.IntegerValue == id.IntegerValue);
                            if (e != null)
                            {
                                var bbox = e.get_BoundingBox(null);
                                if (bbox != null)
                                {
                                    // DebugLogger.Info($"[SectionBoxDiag] Linked Element {e.Id} BBox Min({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}) Max({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3})");
                                }
                            }
                        }
                    }
                    catch { }

                    // No fallback for linked elements - if no elements pass solid filter, skip
                    // DebugLogger.Info($"[SectionBoxDiag] Linked elements: no fallback implemented");
                    if (passingLinkIds.Count > 0)
                    {
                        filteredList.AddRange(elementsInLink
                            .Where(e => passingLinkIds.Contains(e.Id))
                            .Select(e => (e, (Transform?)linkInstance.GetTotalTransform())));
                    }
                }
            }

            return filteredList;
        }

        /// <summary>
        /// Get section box bounds from 3D view for efficient spatial filtering
        /// </summary>
        /// <param name="view3D">The 3D view</param>
        /// <returns>BoundingBoxXYZ in world coordinates, or null if no section box</returns>
        public static BoundingBoxXYZ? GetSectionBoxBounds(View3D view3D)
        {
            if (view3D == null || !view3D.IsSectionBoxActive)
            {
                return null;
            }
            
            try
            {
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                if (sectionBox == null) return null;
                
                Transform transform = sectionBox.Transform;
                
                // ✅ CRITICAL FIX: Transform ALL 8 corners to handle rotation correctly
                // Transforming just Min/Max works only if rotation is 0. For rotated boxes,
                // the "Min" corner locally might not be the "Min" corner in World logic.
                
                // 1. Get all 8 local corners
                XYZ min = sectionBox.Min;
                XYZ max = sectionBox.Max;
                
                XYZ[] localCorners = new XYZ[8]
                {
                    new XYZ(min.X, min.Y, min.Z), // 0: Min
                    new XYZ(max.X, min.Y, min.Z), // 1
                    new XYZ(min.X, max.Y, min.Z), // 2
                    new XYZ(max.X, max.Y, min.Z), // 3
                    
                    new XYZ(min.X, min.Y, max.Z), // 4
                    new XYZ(max.X, min.Y, max.Z), // 5
                    new XYZ(min.X, max.Y, max.Z), // 6
                    new XYZ(max.X, max.Y, max.Z)  // 7: Max
                };
                
                // 2. Transform all corners to world coordinates
                // 3. Find global Min/Max from the cloud of points
                double wMinX = double.MaxValue, wMinY = double.MaxValue, wMinZ = double.MaxValue;
                double wMaxX = double.MinValue, wMaxY = double.MinValue, wMaxZ = double.MinValue;
                
                foreach (XYZ corner in localCorners)
                {
                    XYZ worldCorner = transform.OfPoint(corner);
                    
                    if (worldCorner.X < wMinX) wMinX = worldCorner.X;
                    if (worldCorner.Y < wMinY) wMinY = worldCorner.Y;
                    if (worldCorner.Z < wMinZ) wMinZ = worldCorner.Z;
                    
                    if (worldCorner.X > wMaxX) wMaxX = worldCorner.X;
                    if (worldCorner.Y > wMaxY) wMaxY = worldCorner.Y;
                    if (worldCorner.Z > wMaxZ) wMaxZ = worldCorner.Z;
                }
                
                return new BoundingBoxXYZ
                {
                    Min = new XYZ(wMinX, wMinY, wMinZ),
                    Max = new XYZ(wMaxX, wMaxY, wMaxZ)
                };
            }
            catch
            {
                return null;
            }
        }

        private static Solid? GetSectionBoxAsSolid(View3D view3D)
        {
            BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
            Transform transform = sectionBox.Transform;

            XYZ pt0 = new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Min.Z);
            XYZ pt1 = new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z);
            XYZ pt2 = new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z);
            XYZ pt3 = new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z);

            var profile = new List<Curve> { Line.CreateBound(pt0, pt1), Line.CreateBound(pt1, pt2), Line.CreateBound(pt2, pt3), Line.CreateBound(pt3, pt0) };
            CurveLoop curveLoop = CurveLoop.Create(profile);
            double height = sectionBox.Max.Z - sectionBox.Min.Z;

            Solid axisAlignedSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { curveLoop }, XYZ.BasisZ, height);
            return SolidUtils.CreateTransformed(axisAlignedSolid, transform);
        }

        // Fast bounding box intersection test
        private static bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            return !(max1.X < min2.X || min1.X > max2.X ||
                     max1.Y < min2.Y || min1.Y > max2.Y ||
                     max1.Z < min2.Z || min1.Z > max2.Z);
        }
    }
}
