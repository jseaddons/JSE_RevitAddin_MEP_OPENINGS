using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Helper class to handle Ray Casting (ReferenceIntersector) logic for linear MEP elements.
    /// Separated from MepIntersectionService to maintain clean code and SOLID principles.
    /// </summary>
    public static class RayCastIntersectionHelper
    {
        public static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            View3D view3D,
            Action<string> log)
        {
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();

            if (view3D == null)
            {
                log?.Invoke("[RayCastHelper] View3D is null. Skipping Ray Casting.");
                return results;
            }

            try
            {
                // 1. Setup ReferenceIntersector
                // Filter for structural categories to optimize the intersector
                var structCats = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFoundation
                };

                var filter = new ElementMulticategoryFilter(structCats);

                // Use FindReferencesInRevitLinks to support linked models
                var intersector = new ReferenceIntersector(filter, FindReferenceTarget.All, view3D);
                intersector.FindReferencesInRevitLinks = true;

                // 2. Index Structural Elements for fast lookup
                // ReferenceIntersector finds *everything* visible. We only want intersections with our target 'structuralElements'.
                // We build a HashSet of UniqueIds for O(1) checking.
                var structuralIds = new HashSet<string>();
                foreach (var (elem, _) in structuralElements)
                {
                    if (elem != null) structuralIds.Add(elem.UniqueId);
                }

                log?.Invoke($"[RayCastHelper] Initialized Intersector. Target Structural Elements: {structuralIds.Count}");

                // 3. Process MEP Elements
                foreach (var (mepElement, mepTransform) in mepElements)
                {
                    // Get Centerline
                    var mepBBox = mepElement.get_BoundingBox(null);
                    if (mepBBox == null) continue;

                    // Use MepIntersectionService's logic to get the line (we assume access to public/internal static or duplicate duplicate logic?)
                    // Since MepIntersectionService.GetElementLineWithSource is private, we will duplicate the basic line extraction here 
                    // or assume the caller passes the line? 
                    // Better: The caller (MepIntersectionService) can just pass the element, and we'll access the LocationCurve or Connector logic.
                    // For Simplicity reusing existing public methods or simple LocationCurve check for now.
                    // IMPORTANT: MepIntersectionService has specific logic for "Fallback" lines from BBox.
                    // To avoid duplicating complex logic, we should probably make GetElementLineWithSource internal or public in MepIntersectionService?
                    // Or just use LocationCurve which covers 99% of Ducts/Pipes. 
                    // Let's use LocationCurve + Connector fallback locally implemented for robustness.
                    
                    Line line = GetElementCenterLine(mepElement);
                    if (line == null) continue;

                    // If MEP element is transformed (e.g. from link), we interpret the line in HOST coordinates
                    XYZ start = line.GetEndPoint(0);
                    XYZ end = line.GetEndPoint(1);

                    if (mepTransform != null)
                    {
                        start = mepTransform.OfPoint(start);
                        end = mepTransform.OfPoint(end);
                    }

                    var dir = (end - start).Normalize();
                    var length = start.DistanceTo(end);

                    // Shoot Ray
                    // Find all hits along the ray
                    var hits = intersector.Find(start, dir);

                    foreach (var context in hits)
                    {
                        var reference = context.GetReference();
                        var hitPoint = reference.GlobalPoint;

                        // Distance Check
                        double dist = start.DistanceTo(hitPoint);
                        // Accessing 'tolerance' - slight buffer
                        if (dist > length + 0.01) continue; 
                        if (dist < -0.01) continue; 

                        // Identify Hit Element
                        Element hitElement = null;

                        if (reference.LinkedElementId != ElementId.InvalidElementId)
                        {
                            // Linked Element
                            var linkInstance = view3D.Document.GetElement(reference.ElementId) as RevitLinkInstance;
                            if (linkInstance != null)
                            {
                                var linkDoc = linkInstance.GetLinkDocument();
                                if (linkDoc != null)
                                {
                                    hitElement = linkDoc.GetElement(reference.LinkedElementId);
                                }
                            }
                        }
                        else
                        {
                            // Host Element
                            hitElement = view3D.Document.GetElement(reference.ElementId);
                        }

                        // Validation: Is this element in our structural list?
                        if (hitElement != null && structuralIds.Contains(hitElement.UniqueId))
                        {
                            // Create Result
                            // Create a small bounding box around intersection point
                            double d = 0.1;
                            var bbox = new BoundingBoxXYZ();
                            bbox.Min = new XYZ(hitPoint.X - d, hitPoint.Y - d, hitPoint.Z - d);
                            bbox.Max = new XYZ(hitPoint.X + d, hitPoint.Y + d, hitPoint.Z + d);

                            results.Add((mepElement, hitElement, bbox, hitPoint));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[RayCastHelper] Error: {ex.Message}");
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"Stack: {ex.StackTrace}");
            }

            return results;
        }

        private static Line GetElementCenterLine(Element element)
        {
            if (element.Location is LocationCurve locCurve && locCurve.Curve is Line line)
            {
                return line;
            }
            
            // Fallback for some MEP elements (e.g. fitting based?)
            // If it's a MEPCurve (Duct, Pipe, etc.) it should have LocationCurve.
            // If not (e.g. Fittings), Ray Casting might not be appropriate or they are short.
            // For now, return null to skip non-linear elements in this helper.
            return null;
        }
    }
}
