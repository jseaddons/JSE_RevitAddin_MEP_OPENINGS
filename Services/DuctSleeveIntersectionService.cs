using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class DuctSleeveIntersectionService
    {
        // Collect all structural elements (walls, floors, framing) from host and visible linked models only
        public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc)
        {
            var elements = new List<(Element, Transform?)>();
            // Host model
            var hostElements = new FilteredElementCollector(doc)
                .WherePasses(new ElementMulticategoryFilter(new[] {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_Floors
                }))
                .WhereElementIsNotElementType()
                .ToElements();
            foreach (Element e in hostElements) elements.Add((e, null));
            // Linked models (only visible links)
            var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null || doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(doc.ActiveView)) continue;
                var linkTransform = linkInstance.GetTotalTransform();
                var linkedElements = new FilteredElementCollector(linkDoc)
                    .WherePasses(new ElementMulticategoryFilter(new[] {
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_StructuralFraming,
                        BuiltInCategory.OST_Floors
                    }))
                    .WhereElementIsNotElementType()
                    .ToElements();
                foreach (Element e in linkedElements) elements.Add((e, linkTransform));
            }
            return elements;
        }

        // Returns intersection bounding box and center for sleeve placement
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(
            Duct duct, List<(Element, Transform?)> structuralElements)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            var locCurve = duct.Location as LocationCurve;
            var curve = locCurve?.Curve as Line;
            if (curve == null) return results;
            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;
                try
                {
                    var options = new Options();
                    var geometry = structuralElement.get_Geometry(options);
                    Solid? solid = null;
                    foreach (GeometryObject geomObj in geometry)
                    {
                        if (geomObj is Solid s && s.Volume > 0) { solid = s; break; }
                        else if (geomObj is GeometryInstance gi)
                        {
                            foreach (GeometryObject instObj in gi.GetInstanceGeometry())
                                if (instObj is Solid s2 && s2.Volume > 0) { solid = s2; break; }
                            if (solid != null) break;
                        }
                    }
                    if (solid == null) continue;
                    if (linkTransform != null) solid = SolidUtils.CreateTransformed(solid, linkTransform);
                    // Intersect solid with duct line, collect intersection points
                    var intersectionPoints = new List<XYZ>();
                    foreach (Face face in solid.Faces)
                    {
                        IntersectionResultArray ira;
                        if (face.Intersect(curve, out ira) == SetComparisonResult.Overlap && ira != null)
                            foreach (IntersectionResult ir in ira) intersectionPoints.Add(ir.XYZPoint);
                    }
                    if (intersectionPoints.Count > 0)
                    {
                        // Compute bounding box of intersection points
                        double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                        foreach (var pt in intersectionPoints)
                        {
                            if (pt.X < minX) minX = pt.X;
                            if (pt.Y < minY) minY = pt.Y;
                            if (pt.Z < minZ) minZ = pt.Z;
                            if (pt.X > maxX) maxX = pt.X;
                            if (pt.Y > maxY) maxY = pt.Y;
                            if (pt.Z > maxZ) maxZ = pt.Z;
                        }
                        var bbox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(minX, minY, minZ),
                            Max = new XYZ(maxX, maxY, maxZ)
                        };
                        var center = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
                        results.Add((structuralElement, bbox, center));
                    }
                }
                catch { }
            }
            return results;
        }
    }
}

