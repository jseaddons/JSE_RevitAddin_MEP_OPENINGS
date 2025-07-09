

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Generalized service for detecting intersections between elements in any two documents (active or linked),
    /// with support for filtering linked files by visibility in the active view.
    /// </summary>
    public class IntersectionDetectionService
    {
        // ...existing code...

        /// <summary>
        /// Finds intersections using bounding box logic only (mimics FireDamperPlaceCommand).
        /// </summary>
        public List<IntersectionResult> FindIntersectionsBoundingBoxOnly(
            IEnumerable<Element> sourceElements,
            IEnumerable<Element> targetElements,
            Dictionary<Element, Transform> sourceTransforms,
            Dictionary<Element, Transform> targetTransforms,
            View3D view3D,
            IEnumerable<Element> openingElements = null)
        {
            var results = new List<IntersectionResult>();
            foreach (var sourceElem in sourceElements)
            {
                var srcBbox = sourceElem.get_BoundingBox(view3D);
                if (srcBbox == null) continue;
                var srcXf = sourceTransforms != null && sourceTransforms.ContainsKey(sourceElem) ? sourceTransforms[sourceElem] : Transform.Identity;
                var srcBboxTrans = TransformBoundingBox(srcBbox, srcXf);
                foreach (var targetElem in targetElements)
                {
                    var tgtBbox = targetElem.get_BoundingBox(view3D);
                    if (tgtBbox == null) continue;
                    var tgtXf = targetTransforms != null && targetTransforms.ContainsKey(targetElem) ? targetTransforms[targetElem] : Transform.Identity;
                    var tgtBboxTrans = TransformBoundingBox(tgtBbox, tgtXf);
                    if (BoundingBoxesIntersect(srcBboxTrans, tgtBboxTrans))
                    {
                        // Check for existing opening at intersection point (for Floor only)
                        bool hasOpening = false;
                        if (openingElements != null && targetElem.Category != null && targetElem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                        {
                            // Use center of intersection bbox as test point
                            var intersectionPt = (srcBboxTrans.Min + srcBboxTrans.Max) / 2.0;
                            foreach (var opening in openingElements)
                            {
                                var openingBbox = opening.get_BoundingBox(view3D);
                                if (openingBbox == null) continue;
                                // Simple bbox check for opening overlap
                                if (BoundingBoxesIntersect(openingBbox, tgtBboxTrans))
                                {
                                    // Optionally, check if intersectionPt is inside opening bbox
                                    if (intersectionPt.X >= openingBbox.Min.X && intersectionPt.X <= openingBbox.Max.X &&
                                        intersectionPt.Y >= openingBbox.Min.Y && intersectionPt.Y <= openingBbox.Max.Y &&
                                        intersectionPt.Z >= openingBbox.Min.Z && intersectionPt.Z <= openingBbox.Max.Z)
                                    {
                                        hasOpening = true;
                                        break;
                                    }
                                }
                            }
                        }
                        if (!hasOpening)
                        {
                            results.Add(new IntersectionResult
                            {
                                SourceElement = sourceElem,
                                TargetElement = targetElem,
                                IntersectionPoint = (srcBboxTrans.Min + srcBboxTrans.Max) / 2.0,
                                Proximity = 0
                            });
                        }
                    }
                }
            }
            return results;
        }

        // ...existing code...

        /// <summary>
        /// Finds intersections using ReferenceIntersector (robust, Revit-native, handles links and transforms).
        /// </summary>
        public List<IntersectionResult> FindIntersectionsWithReferenceIntersector(
            UIDocument uiDoc,
            Document doc,
            IEnumerable<Element> mepElements,
            View3D view3D,
            IEnumerable<Element> openingElements = null)
        {
            var results = new List<IntersectionResult>();
            // Filter for structural categories
            var structuralFilters = new List<ElementFilter>
            {
                new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming),
                new ElementCategoryFilter(BuiltInCategory.OST_Walls)
            };
            ElementFilter structuralFilter = new LogicalOrFilter(structuralFilters);

            var refIntersector = new ReferenceIntersector(structuralFilter, FindReferenceTarget.Face, view3D)
            {
                FindReferencesInRevitLinks = true
            };

            foreach (var mepElement in mepElements)
            {
                var locCurve = mepElement.Location as LocationCurve;
                if (locCurve?.Curve is Line line)
                {
                    var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };
                    foreach (double t in sampleFractions)
                    {
                        var samplePt = line.Evaluate(t, true);
                        var elementDirection = line.Direction.Normalize();
                        var directions = new[]
                        {
                            elementDirection,
                            elementDirection.Negate(),
                            XYZ.BasisZ,
                            XYZ.BasisZ.Negate(),
                            new XYZ(-elementDirection.Y, elementDirection.X, 0).Normalize(),
                            new XYZ(elementDirection.Y, -elementDirection.X, 0).Normalize()
                        };
                        foreach (var direction in directions)
                        {
                            var hits = refIntersector.Find(samplePt, direction);
                            if (hits != null && hits.Count > 0)
                            {
                                foreach (var hit in hits)
                                {
                                    if (hit.Proximity > 0.33) continue;
                                    var reference = hit.GetReference();
                                    Element structuralElement = null;
                                    var linkInst = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                                    if (linkInst != null)
                                    {
                                        var linkedDoc = linkInst.GetLinkDocument();
                                        if (linkedDoc != null)
                                            structuralElement = linkedDoc.GetElement(reference.LinkedElementId);
                                    }
                                    else
                                    {
                                        structuralElement = doc.GetElement(reference.ElementId);
                                    }
                                    if (structuralElement != null &&
                                        (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors ||
                                         structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming ||
                                         structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls))
                                    {
                                        // For floors, check for existing opening at intersection point
                                        bool hasOpening = false;
                                        if (openingElements != null && structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                                        {
                                            foreach (var opening in openingElements)
                                            {
                                                var openingBbox = opening.get_BoundingBox(view3D);
                                                if (openingBbox == null) continue;
                                                if (samplePt.X >= openingBbox.Min.X && samplePt.X <= openingBbox.Max.X &&
                                                    samplePt.Y >= openingBbox.Min.Y && samplePt.Y <= openingBbox.Max.Y &&
                                                    samplePt.Z >= openingBbox.Min.Z && samplePt.Z <= openingBbox.Max.Z)
                                                {
                                                    hasOpening = true;
                                                    break;
                                                }
                                            }
                                        }
                                        if (!hasOpening)
                                        {
                                            results.Add(new IntersectionResult
                                            {
                                                SourceElement = mepElement,
                                                TargetElement = structuralElement,
                                                IntersectionPoint = samplePt,
                                                Proximity = hit.Proximity
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return results;
        }

        /// <summary>
        /// Finds intersections between source and target elements, supporting active and linked documents.
        /// Only processes linked files that are visible in the given view.
        /// </summary>
        /// <param name="sourceTransforms">Dictionary mapping each source element to its transform (Identity for host, link transform for linked)</param>
        /// <param name="targetTransforms">Dictionary mapping each target element to its transform (Identity for host, link transform for linked)</param>
        public List<IntersectionResult> FindIntersections(
            UIDocument uiDoc,
            Document sourceDoc,
            IEnumerable<Element> sourceElements,
            Dictionary<Element, Transform> sourceTransforms,
            Document targetDoc,
            IEnumerable<Element> targetElements,
            Dictionary<Element, Transform> targetTransforms,
            View3D view3D)
        {
            var results = new List<IntersectionResult>();

            // If source or target is a linked file, check visibility in the active view
            if (sourceDoc.IsLinked && !IsLinkedFileVisible(uiDoc.ActiveView, sourceDoc))
                return results;
            if (targetDoc.IsLinked && !IsLinkedFileVisible(uiDoc.ActiveView, targetDoc))
                return results;

            // Get section box (host coordinates)
            BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
            foreach (var sourceElem in sourceElements)
            {
                BoundingBoxXYZ srcBbox = sourceElem.get_BoundingBox(view3D);
                if (srcBbox == null) continue;
                Transform srcXf = sourceTransforms != null && sourceTransforms.ContainsKey(sourceElem) ? sourceTransforms[sourceElem] : Transform.Identity;
                // Transform bbox to host coords
                BoundingBoxXYZ srcBboxTrans = TransformBoundingBox(srcBbox, srcXf);
                if (!BoundingBoxesIntersect(srcBboxTrans, sectionBox)) continue;

                foreach (var targetElem in targetElements)
                {
                    BoundingBoxXYZ tgtBbox = targetElem.get_BoundingBox(view3D);
                    if (tgtBbox == null) continue;
                    Transform tgtXf = targetTransforms != null && targetTransforms.ContainsKey(targetElem) ? targetTransforms[targetElem] : Transform.Identity;
                    BoundingBoxXYZ tgtBboxTrans = TransformBoundingBox(tgtBbox, tgtXf);
                    if (!BoundingBoxesIntersect(tgtBboxTrans, sectionBox)) continue;

                    // Check if bounding boxes intersect (in host coords)
                    if (!BoundingBoxesIntersect(srcBboxTrans, tgtBboxTrans))
                        continue;

                    // Try to get solids for more precise intersection
                    Solid srcSolid = GetElementSolid(sourceElem);
                    Solid tgtSolid = GetElementSolid(targetElem);
                    if (srcSolid != null && tgtSolid != null)
                    {
                        Solid srcSolidTrans = SolidUtils.CreateTransformed(srcSolid, srcXf);
                        Solid tgtSolidTrans = SolidUtils.CreateTransformed(tgtSolid, tgtXf);
                        Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(srcSolidTrans, tgtSolidTrans, BooleanOperationsType.Intersect);
                        if (intersection != null && intersection.Volume > 1e-6)
                        {
                            // Use centroid of intersection solid as intersection point
                            XYZ intersectionPoint = GetSolidCentroid(intersection);
                            results.Add(new IntersectionResult
                            {
                                SourceElement = sourceElem,
                                TargetElement = targetElem,
                                IntersectionPoint = intersectionPoint,
                                Proximity = 0 // Not used here
                            });
                            continue;
                        }
                    }

                    // Fallback: bounding box center as intersection point
                    XYZ approxPoint = (srcBboxTrans.Min + srcBboxTrans.Max) / 2.0;
                    results.Add(new IntersectionResult
                    {
                        SourceElement = sourceElem,
                        TargetElement = targetElem,
                        IntersectionPoint = approxPoint,
                        Proximity = 0
                    });
                }
            }
            return results;
        }

        /// <summary>
        /// Transforms a bounding box by a given transform (all 8 corners, then computes new min/max)
        /// </summary>
        public BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform xf)
        {
            if (xf == null || xf.IsIdentity) return bbox;
            var pts = new List<XYZ>
            {
                xf.OfPoint(bbox.Min),
                xf.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z)),
                xf.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z)),
                xf.OfPoint(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z)),
                xf.OfPoint(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z)),
                xf.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z)),
                xf.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z)),
                xf.OfPoint(bbox.Max)
            };
            double minX = pts.Min(p => p.X), minY = pts.Min(p => p.Y), minZ = pts.Min(p => p.Z);
            double maxX = pts.Max(p => p.X), maxY = pts.Max(p => p.Y), maxZ = pts.Max(p => p.Z);
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        /// <summary>
        /// Checks if two bounding boxes intersect.
        /// </summary>
        public bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            return (a.Min.X <= b.Max.X && a.Max.X >= b.Min.X)
                && (a.Min.Y <= b.Max.Y && a.Max.Y >= b.Min.Y)
                && (a.Min.Z <= b.Max.Z && a.Max.Z >= b.Min.Z);
        }

        /// <summary>
        /// Attempts to extract a solid from an element (first geometry solid found).
        /// </summary>
        public Solid GetElementSolid(Element elem)
        {
            Options options = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geomElem = elem.get_Geometry(options);
            if (geomElem == null) return null;
            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid solid && solid.Volume > 1e-6)
                    return solid;
                if (obj is GeometryInstance inst)
                {
                    foreach (GeometryObject instObj in inst.GetInstanceGeometry())
                    {
                        if (instObj is Solid instSolid && instSolid.Volume > 1e-6)
                            return instSolid;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Gets the centroid of a solid.
        /// </summary>
        private XYZ GetSolidCentroid(Solid solid)
        {
            if (solid == null || solid.Faces.Size == 0) return XYZ.Zero;
            double x = 0, y = 0, z = 0;
            int n = 0;
            foreach (Edge edge in solid.Edges)
            {
                IList<XYZ> pts = edge.Tessellate();
                foreach (var pt in pts)
                {
                    x += pt.X; y += pt.Y; z += pt.Z; n++;
                }
            }
            if (n == 0) return XYZ.Zero;
            return new XYZ(x / n, y / n, z / n);
        }
        /// <summary>
        /// Checks if a linked file is visible in the given view.
        /// </summary>
        public bool IsLinkedFileVisible(View view, Document linkedDoc)
        {
            // Find the RevitLinkInstance for the linked document
            var linkInstances = new FilteredElementCollector(view.Document)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() == linkedDoc);

            // Use reflection to support Revit versions before 2022
            var getHiddenElementIdsMethod = view.GetType().GetMethod("GetHiddenElementIds");
            HashSet<ElementId> hiddenIds = null;
            if (getHiddenElementIdsMethod != null)
            {
                hiddenIds = getHiddenElementIdsMethod.Invoke(view, null) as HashSet<ElementId>;
            }
            foreach (var linkInstance in linkInstances)
            {
                // If we can't get hidden IDs, assume visible
                if (hiddenIds == null || !hiddenIds.Contains(linkInstance.Id))
                {
                    return true;
                }
            }
            return false;
        }
    }

    /// <summary>
    /// Generalized intersection result for any two elements (source and target)
    /// </summary>
    public class IntersectionResult
    {
        public Element SourceElement { get; set; }
        public Element TargetElement { get; set; }
        public XYZ IntersectionPoint { get; set; }
        public double Proximity { get; set; }
    }
}