using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

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
                var hostFilter = new ElementIntersectsSolidFilter(sectionBoxSolid);
                var hostIds = hostElements.Select(e => e.Id).ToList();
                DebugLogger.Info($"[SectionBoxDiag] Host elements count={hostElements.Count}, sampleIds={string.Join(",", hostIds.Take(6).Select(id => id.IntegerValue.ToString()))}");
                var passingHostIds = new FilteredElementCollector(uiDoc.Document, hostIds)
                    .WherePasses(hostFilter)
                    .ToElementIds();
                filteredList.AddRange(hostElements.Where(e => passingHostIds.Contains(e.Id)).Select(e => (e, (Transform?)null)));
                try
                {
                    DebugLogger.Info($"[SectionBoxDiag] Host section-solid volume={sectionBoxSolid.Volume}, passingHostCount={passingHostIds.Count}");
                }
                catch { }

                // If no host elements passed the solid filter, attempt a conservative bounding-box overlap fallback
                if (passingHostIds.Count == 0)
                {
                    try
                    {
                        var secBbox = GetSectionBoxBounds(view3D);
                        if (secBbox != null)
                        {
                            DebugLogger.Info($"[SectionBoxDiag] No passingHostIds from solid filter; attempting bbox-overlap fallback. secBbox=Min={secBbox.Min},Max={secBbox.Max}");
                            int added = 0;
                            foreach (var e in hostElements)
                            {
                                try
                                {
                                    var ebbox = e.get_BoundingBox(null);
                                    if (ebbox == null)
                                    {
                                        DebugLogger.Info($"[SectionBoxDiag] Element {e.Id.IntegerValue} bbox=null");
                                        continue;
                                    }
                                    bool overlap = !(ebbox.Max.X < secBbox.Min.X || ebbox.Min.X > secBbox.Max.X ||
                                                     ebbox.Max.Y < secBbox.Min.Y || ebbox.Min.Y > secBbox.Max.Y ||
                                                     ebbox.Max.Z < secBbox.Min.Z || ebbox.Min.Z > secBbox.Max.Z);
                                    DebugLogger.Info($"[SectionBoxDiag] Element {e.Id.IntegerValue} bboxOverlap={overlap}, bboxMin={ebbox.Min}, bboxMax={ebbox.Max}");
                                    if (overlap)
                                    {
                                        filteredList.Add((e, (Transform?)null));
                                        added++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    DebugLogger.Info($"[SectionBoxDiag] Exception checking bbox for element {e.Id.IntegerValue}: {ex.Message}");
                                }
                            }
                            DebugLogger.Info($"[SectionBoxDiag] BBox fallback addedCount={added}");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Info($"[SectionBoxDiag] BBox fallback failed: {ex.Message}");
                    }
                }
            }

            // Filter linked elements
            var allLinkInstances = new FilteredElementCollector(uiDoc.Document).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

            foreach (var group in linkedElementGroups)
            {
                var linkInstance = allLinkInstances.FirstOrDefault(li => li.GetLinkDocument()?.Title == group.Key);
                if (linkInstance == null) continue;

                // Transform the section box solid into the coordinate system of the linked document
                Transform inverseTransform = linkInstance.GetTotalTransform().Inverse;
                Solid transformedSolid = SolidUtils.CreateTransformed(sectionBoxSolid, inverseTransform);

                try
                {
                    DebugLogger.Info($"[SectionBoxDiag] Processing link='{group.Key}', TransformOrigin={linkInstance.GetTotalTransform().Origin}, TransformedSolidVol={transformedSolid.Volume}");
                }
                catch { }

                var elementsInLink = group.Select(t => t.element).ToList();
                if (elementsInLink.Any())
                {
                    var linkFilter = new ElementIntersectsSolidFilter(transformedSolid);
                    var passingLinkIds = new FilteredElementCollector(linkInstance.GetLinkDocument(), elementsInLink.Select(e => e.Id).ToList())
                        .WherePasses(linkFilter)
                        .ToElementIds();
                    
                    try
                    {
                        DebugLogger.Info($"[SectionBoxDiag] link='{group.Key}' passingCount={passingLinkIds.Count}");
                        foreach (var id in passingLinkIds.Take(3)) DebugLogger.Info($"[SectionBoxDiag] link sample passing id={id.IntegerValue}");
                    }
                    catch { }

                    filteredList.AddRange(elementsInLink
                        .Where(e => passingLinkIds.Contains(e.Id))
                        .Select(e => (e, (Transform?)linkInstance.GetTotalTransform())));
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
                
                // Transform section box to world coordinates
                XYZ worldMin = transform.OfPoint(sectionBox.Min);
                XYZ worldMax = transform.OfPoint(sectionBox.Max);
                
                return new BoundingBoxXYZ
                {
                    Min = new XYZ(
                        System.Math.Min(worldMin.X, worldMax.X),
                        System.Math.Min(worldMin.Y, worldMax.Y),
                        System.Math.Min(worldMin.Z, worldMax.Z)),
                    Max = new XYZ(
                        System.Math.Max(worldMin.X, worldMax.X),
                        System.Math.Max(worldMin.Y, worldMax.Y),
                        System.Math.Max(worldMin.Z, worldMax.Z))
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
    }
}
