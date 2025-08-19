using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using Autodesk.Revit.ApplicationServices;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CableTraySleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            int deletedCount = SleeveZeroValueDeleter.DeleteZeroValueSleeves(doc);
            DebugLogger.Log($"Deleted {deletedCount} cable tray sleeves with zero width, height, or depth.");

            var allFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family != null).ToList();

            var ctWallSymbols = allFamilySymbols
                .Where(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnWall", System.StringComparison.OrdinalIgnoreCase)).ToList();
            var ctSlabSymbols = allFamilySymbols
                .Where(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnSlab", System.StringComparison.OrdinalIgnoreCase)).ToList();
            if (!ctWallSymbols.Any() && !ctSlabSymbols.Any())
            {
                TaskDialog.Show("Error", "Please load cable tray sleeve opening families.");
                return Result.Failed;
            }

            using (var txActivate = new Transaction(doc, "Activate CT Symbols"))
            {
                txActivate.Start();
                foreach (var sym in ctWallSymbols) if (!sym.IsActive) sym.Activate();
                foreach (var sym in ctSlabSymbols) if (!sym.IsActive) sym.Activate();
                txActivate.Commit();
            }

            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                TaskDialog.Show("Error", "No non-template 3D view.");
                return Result.Failed;
            }

            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
            {
                FindReferencesInRevitLinks = true
            };

            var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var trayTuples = mepElements
                .Where(tuple => tuple.Item1 is CableTray)
                .Select(tuple => (tuple.Item1 as CableTray, tuple.Item2))
                .Where(t => t.Item1 != null).ToList();

            if (trayTuples.Count == 0)
            {
                TaskDialog.Show("Info", "No cable trays found.");
                return Result.Succeeded;
            }

            var directStructuralElements = CollectStructuralElementsForDirectIntersection(doc);

            var placer = new CableTraySleevePlacer(doc);

            HashSet<ElementId> processedCableTrays = new HashSet<ElementId>();

            int placedCount = 0, skippedExistingCount = 0, missingCount = 0;

            using (var tx = new Transaction(doc, "Place Cable Tray Sleeves"))
            {
                tx.Start();
                foreach (var tuple in trayTuples)
                {
                    var tray = tuple.Item1;
                    var transform = tuple.Item2;
                    if (tray == null || processedCableTrays.Contains(tray.Id)) continue;

                    var curve = (tray.Location as LocationCurve)?.Curve as Line;
                    if (curve == null) continue;

                    Line hostLine = curve;
                    if (transform != null)
                        hostLine = Line.CreateBound(transform.OfPoint(curve.GetEndPoint(0)),
                                                     transform.OfPoint(curve.GetEndPoint(1)));

                    double width  = tray.LookupParameter("Width")?.AsDouble() ?? 0.0;
                    double height = tray.LookupParameter("Height")?.AsDouble() ?? 0.0;
                    XYZ rayDir    = hostLine.Direction;

                    // --- Ray-cast for walls ---
                    var testPoints = new List<XYZ>
                    {
                        hostLine.GetEndPoint(0),
                        hostLine.Evaluate(0.25, true),
                        hostLine.Evaluate(0.5, true),
                        hostLine.Evaluate(0.75, true),
                        hostLine.GetEndPoint(1)
                    };
                    var allWallHits = new List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)>();
                    foreach (var pt in testPoints)
                    {
                        var fwd  = refIntersector.Find(pt, rayDir)?.OrderBy(h => h.Proximity);
                        var back = refIntersector.Find(pt, rayDir.Negate())?.OrderBy(h => h.Proximity);
                        var perp1 = refIntersector.Find(pt, new XYZ(-rayDir.Y, rayDir.X, 0).Normalize())?.OrderBy(h => h.Proximity);
                        var perp2 = refIntersector.Find(pt, new XYZ(rayDir.Y, -rayDir.X, 0).Normalize())?.OrderBy(h => h.Proximity);

                        if (fwd  != null) allWallHits.AddRange(fwd.Select(h => (h, rayDir, pt)));
                        if (back != null) allWallHits.AddRange(back.Select(h => (h, rayDir.Negate(), pt)));
                        if (perp1 != null) allWallHits.AddRange(perp1.Select(h => (h, new XYZ(-rayDir.Y, rayDir.X, 0).Normalize(), pt)));
                        if (perp2 != null) allWallHits.AddRange(perp2.Select(h => (h, new XYZ(rayDir.Y, -rayDir.X, 0).Normalize(), pt)));
                    }

                    Element thickestWall = null;
                    double maxThickness = 0.0;
                    XYZ wallIntersectionPoint = XYZ.Zero;

                    foreach (var hit in allWallHits)
                    {
                var r = hit.hit.GetReference();
                var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                ElementId elemId = linkInst != null ? r.LinkedElementId : r.ElementId;
                Element hitElement = targetDoc?.GetElement(elemId);

                    Wall? w = hitElement as Wall;
                    if (w != null)
                {
                    if (w.Width > maxThickness)
                    {
                        maxThickness = w.Width;
                        thickestWall = w;
                        wallIntersectionPoint = hit.rayOrigin + hit.direction * hit.hit.Proximity;
                    }
                }
                    }

                    if (thickestWall == null) continue;

                    // Project tray point onto wall centre-line
                    XYZ wallCentre = ((thickestWall.Location as LocationCurve)?.Curve as Line)
                                     ?.Project(wallIntersectionPoint)?.XYZPoint ?? wallIntersectionPoint;
                    XYZ wallNormal = ((thickestWall.Location as LocationCurve)?.Curve as Line)?.Direction.CrossProduct(XYZ.BasisZ).Normalize() ?? XYZ.BasisY;
                    double wallWidth = (thickestWall is Wall wallObj) ? wallObj.Width : 0.0;
                    XYZ placePoint = wallCentre + wallNormal * (-wallWidth * 0.5);

                    var wallFamilySymbol = ctWallSymbols.FirstOrDefault();
                    if (wallFamilySymbol == null)
                    {
                        TaskDialog.Show("Error", "Cable tray wall sleeve family missing.");
                        continue;
                    }

                    if (PlaceCableTraySleeveAtLocation_Wall(doc, wallFamilySymbol, thickestWall, placePoint,
                                                             wallNormal, width, height, tray, Math.Abs(rayDir.X) > Math.Abs(rayDir.Y), wallWidth))
                    {
                        placedCount++;
                        processedCableTrays.Add(tray.Id);
                    }
                    else
                    {
                        skippedExistingCount++;
                        processedCableTrays.Add(tray.Id);
                    }
                }
                tx.Commit();
            }

            string summary = $"CableTraySleeveCommand summary: Placed={placedCount}, Skipped={skippedExistingCount}, Missing={missingCount}";
            DebugLogger.Log(summary);
            return Result.Succeeded;
        }

        private bool PlaceCableTraySleeveAtLocation_Wall(Document doc, FamilySymbol sym, Element host, XYZ pt, XYZ dir,
                                                         double w, double h, CableTray tray, bool isX, double? depth = null)
        {
            return new CableTraySleevePlacer(doc).PlaceCableTraySleeve(tray, pt, w, h, dir, sym, host) != null;
        }

        private bool PlaceCableTraySleeveAtLocation_Structural(Document doc, FamilySymbol sym, Element host, XYZ pt,
                                                               XYZ dir, double w, double h, CableTray tray)
        {
            return new CableTraySleevePlacer(doc).PlaceCableTraySleeve(tray, pt, w, h, dir, sym, host) != null;
        }

        private List<(Element element, Transform linkTransform)> CollectStructuralElementsForDirectIntersection(Document doc)
        {
            var elements = new List<(Element, Transform)>();
            var links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                            .Where(l => l.GetLinkDocument() != null);
            foreach (var li in links)
            {
                var linked = li.GetLinkDocument();
                var t = li.GetTotalTransform();
                elements.AddRange(new FilteredElementCollector(linked)
                    .WherePasses(new ElementMulticategoryFilter(new[] { BuiltInCategory.OST_Walls,
                                                                         BuiltInCategory.OST_StructuralFraming,
                                                                         BuiltInCategory.OST_Floors }))
                    .WhereElementIsNotElementType()
                    .Select(e => (e, t)));
            }
            return elements;
        }

        private List<(Element structuralElement, XYZ intersectionPoint)> FindDirectStructuralIntersections(
            CableTray cableTray, List<(Element element, Transform linkTransform)> structuralElements)
        {
            var intersections = new List<(Element, XYZ)>();
            var curve = (cableTray.Location as LocationCurve)?.Curve as Line;
            if (curve == null) return intersections;

            foreach (var (elem, xf) in structuralElements)
            {
                try
                {
                    Options opt = new Options();
                    GeometryElement geom = elem.get_Geometry(opt);
                    Solid? solid = null;
                    foreach (var g in geom)
                    {
                        if (g is Solid s && s.Volume > 0) { solid = s; break; }
                        if (g is GeometryInstance gi)
                            foreach (var ig in gi.GetInstanceGeometry())
                                if (ig is Solid s2 && s2.Volume > 0) { solid = s2; break; }
                    }
                    if (solid == null) continue;
                    if (xf != null) solid = SolidUtils.CreateTransformed(solid, xf);

                    var pts = new List<XYZ>();
                    foreach (Face f in solid.Faces)
                    {
                        IntersectionResultArray ira;
                        if (f.Intersect(curve, out ira) == SetComparisonResult.Overlap && ira != null)
                            foreach (IntersectionResult ir in ira) pts.Add(ir.XYZPoint);
                    }
                    if (pts.Count > 0)
                    {
                        if (pts.Count >= 2)
                        {
                            double max = pts.Max(p => p.DistanceTo(pts.First()));
                            XYZ a = pts.First(), b = pts.Last();
                            foreach (var p1 in pts)
                                foreach (var p2 in pts)
                                    if (p1.DistanceTo(p2) > max) { max = p1.DistanceTo(p2); a = p1; b = p2; }
                            intersections.Add((elem, (a + b) * 0.5));
                        }
                        else
                        {
                            intersections.Add((elem, pts[0]));
                        }
                    }
                }
                catch { /* ignore */ }
            }
            return intersections;
        }

        private XYZ ProjectPointOntoLine(XYZ point, Line line)
        {
            XYZ start = line.GetEndPoint(0);
            XYZ dir = line.Direction;
            double t = Math.Max(0, Math.Min((point - start).DotProduct(dir), line.Length));
            return start + dir * t;
        }

        private bool IsPointOnLineSegment(XYZ point, Line line)
        {
            double d1 = point.DistanceTo(line.GetEndPoint(0));
            double d2 = point.DistanceTo(line.GetEndPoint(1));
            return Math.Abs(d1 + d2 - line.Length) < 0.001;
        }
    }
}
