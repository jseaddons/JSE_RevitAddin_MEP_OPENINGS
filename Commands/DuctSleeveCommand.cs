using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DuctSleeveCommand : IExternalCommand
    {

        // Helper: collect all structural elements (walls, floors, framing) from host and linked models

        // Helper: collect all structural elements (walls, floors, framing) from host and linked models
        private static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersection(Document doc)
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

        // Helper: direct solid intersection for ducts (mimics cable tray logic)
        private static List<(Element, XYZ)> FindDirectStructuralIntersections(
            Duct duct, List<(Element, Transform?)> structuralElements)
        {
            var intersections = new List<(Element, XYZ)>();
            var locCurve = duct.Location as LocationCurve;
            var curve = locCurve?.Curve as Line;
            if (curve == null) return intersections;
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
                    var intersectionPoints = new List<XYZ>();
                    foreach (Face face in solid.Faces)
                    {
                        IntersectionResultArray ira;
                        if (face.Intersect(curve, out ira) == SetComparisonResult.Overlap && ira != null)
                            foreach (IntersectionResult ir in ira) intersectionPoints.Add(ir.XYZPoint);
                    }
                    if (intersectionPoints.Count > 0)
                    {
                        if (intersectionPoints.Count >= 2)
                        {
                            double maxDist = double.MinValue;
                            XYZ ptA = intersectionPoints[0], ptB = intersectionPoints[1];
                            for (int i = 0; i < intersectionPoints.Count - 1; i++)
                                for (int j = i + 1; j < intersectionPoints.Count; j++)
                                {
                                    double dist = intersectionPoints[i].DistanceTo(intersectionPoints[j]);
                                    if (dist > maxDist) { maxDist = dist; ptA = intersectionPoints[i]; ptB = intersectionPoints[j]; }
                                }
                            intersections.Add((structuralElement, new XYZ((ptA.X + ptB.X) / 2, (ptA.Y + ptB.Y) / 2, (ptA.Z + ptB.Z) / 2)));
                        }
                        else
                        {
                            intersections.Add((structuralElement, intersectionPoints[0]));
                        }
                    }
                }
                catch { }
            }
            return intersections;

        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get both wall and slab duct sleeve family symbols
            var ductWallSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("OpeningOnWall")
                    && sym.Name.Replace(" ", "").StartsWith("DS#", StringComparison.OrdinalIgnoreCase));
            var ductSlabSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("OpeningOnSlab")
                    && sym.Name.Replace(" ", "").StartsWith("DS#", StringComparison.OrdinalIgnoreCase));
            if (ductWallSymbol == null && ductSlabSymbol == null)
            {
                TaskDialog.Show("Error", "Please load both wall and slab duct sleeve opening families.");
                return Result.Failed;
            }
            using (var txActivate = new Transaction(doc, "Activate Duct Symbols"))
            {
                txActivate.Start();
                if (ductWallSymbol != null && !ductWallSymbol.IsActive)
                    ductWallSymbol.Activate();
                if (ductSlabSymbol != null && !ductSlabSymbol.IsActive)
                    ductSlabSymbol.Activate();
                txActivate.Commit();
            }

            // Find a non-template 3D view
            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                TaskDialog.Show("Error", "No non-template 3D view found.");
                return Result.Failed;
            }

            // Collect ducts from both host and visible linked models
            var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var ducts = mepElements
                .Where(tuple => tuple.Item1 is Duct)
                .Select(tuple => tuple.Item1 as Duct)
                .Where(d => d != null)
                .ToList();
            if (ducts.Count == 0)
            {
                TaskDialog.Show("Info", "No ducts found in host or linked models.");
                return Result.Succeeded;
            }

            // Collect existing sleeves for suppression (both wall and floor sleeves)
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                    (fi.Symbol.Family.Name.Contains("OpeningOnWall") ||
                     fi.Symbol.Family.Name.Contains("OpeningOnSlab") ||
                     fi.Symbol.Family.Name.Contains("OpeningOnSlabRect")) &&
                    fi.Symbol.Name.StartsWith("DS#"))
                .ToList();

            DebugLogger.Log($"Found {existingSleeves.Count} existing duct sleeves in the model (wall and floor)");

            // Main intersection and placement loop (direct solid intersection, cable tray style)
            int placedCount = 0, skippedCount = 0, errorCount = 0;
            string logDir = @"c:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Logs";
            string logFile = System.IO.Path.Combine(logDir, $"DuctSleeveLog_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            var structuralElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.StructuralElementCollectorHelper.CollectStructuralElementsVisibleOnly(doc);
            // Use robust bounding box intersection for floors
            var structuralElementsVisible = JSE_RevitAddin_MEP_OPENINGS.Helpers.StructuralElementCollectorHelper.CollectStructuralElementsVisibleOnly(doc);
            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                foreach (var duct in ducts)
                {
                    if (duct == null) { skippedCount++; continue; }
                    var locCurve = duct.Location as LocationCurve;
                    var ductLine = locCurve?.Curve as Line;
                    if (ductLine == null)
                    {
                        string msg = $"SKIP: Duct {duct?.Id} is not a line";
                        DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                        System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                        continue;
                    }
                    string ductMsg = $"PROCESSING: Duct {duct.Id} Line Start={ductLine.GetEndPoint(0)}, End={ductLine.GetEndPoint(1)}";
                    DebugLogger.Log($"[DuctSleeveCommand] {ductMsg}");
                    System.IO.File.AppendAllText(logFile, ductMsg + System.Environment.NewLine);

                    // Use robust bounding box intersection for floors
                    var bboxIntersections = JSE_RevitAddin_MEP_OPENINGS.Services.DuctSleeveIntersectionService.FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(duct, structuralElementsVisible);
                    if (bboxIntersections.Count > 0)
                    {
                        foreach (var tuple in bboxIntersections)
                        {
                            Element hostElem = tuple.Item1;
                            BoundingBoxXYZ bbox = tuple.Item2;
                            XYZ placePt = tuple.Item3;
                            string hostType = hostElem?.GetType().Name ?? "null";
                            string hostId = hostElem?.Id.Value.ToString() ?? "null";
                            string hostMsg = $"HOST: Duct {duct.Id} intersects {hostType} {hostId} BBox=({bbox.Min},{bbox.Max})";
                            DebugLogger.Log($"[DuctSleeveCommand] {hostMsg}");
                            System.IO.File.AppendAllText(logFile, hostMsg + System.Environment.NewLine);
                            // Duplicate suppression: log what sleeves are found at this location
                            double tol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                            string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placePt, tol);
                            DebugLogger.Log($"[DuctSleeveCommand] DUPLICATION CHECK at {placePt}:\n{sleeveSummary}");
                            System.IO.File.AppendAllText(logFile, $"DUPLICATION CHECK at {placePt}:\n{sleeveSummary}\n");
                            var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placePt, tol);
                            var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placePt, tol);
                            if (indivDup.Any() || clusterDup.Any())
                            {
                                string msg = $"SKIP: Duct {duct.Id} host {hostType} {hostId} duplicate sleeve (individual or cluster) exists near {placePt}";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                skippedCount++;
                                continue;
                            }
                            double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                            double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                            try
                            {
                                FamilySymbol? symbolToUse = null;
                                if (hostElem is Floor)
                                    symbolToUse = ductSlabSymbol;
                                else if (hostElem is Wall)
                                    symbolToUse = ductWallSymbol;
                                else if (hostElem is FamilyInstance fi && fi.StructuralType == StructuralType.Beam)
                                    symbolToUse = ductWallSymbol;
                                else
                                    symbolToUse = ductWallSymbol; // fallback
                                if (symbolToUse == null)
                                {
                                    string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} no suitable family symbol found.";
                                    DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                    System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                    errorCount++;
                                    continue;
                                }
                                var placer = new DuctSleevePlacer(doc);
                                placer.PlaceDuctSleeve(duct, placePt, w, h2, ductLine.Direction, symbolToUse, hostElem);
                                string msg2 = $"PLACED: Duct {duct.Id} host {hostType} {hostId} at {placePt} size=({w},{h2})";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg2}");
                                System.IO.File.AppendAllText(logFile, msg2 + System.Environment.NewLine);
                                placedCount++;
                            }
                            catch (Exception ex)
                            {
                                string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} failed to place sleeve: {ex.Message}";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                errorCount++;
                            }
                        }
                        continue;
                    }
                    // Fallback: use old intersection logic for non-floor hosts
                    var intersections = FindDirectStructuralIntersections(duct, structuralElements);
                    foreach (var tuple in intersections)
                    {
                        Element hostElem = tuple.Item1;
                        XYZ placePt = tuple.Item2;
                        string hostType = hostElem?.GetType().Name ?? "null";
                        string hostId = hostElem?.Id.Value.ToString() ?? "null";
                        string hostMsg = $"HOST: Duct {duct.Id} intersects {hostType} {hostId}";
                        DebugLogger.Log($"[DuctSleeveCommand] {hostMsg}");
                        System.IO.File.AppendAllText(logFile, hostMsg + System.Environment.NewLine);
                        double tol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                        string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placePt, tol);
                        DebugLogger.Log($"[DuctSleeveCommand] DUPLICATION CHECK at {placePt}:\n{sleeveSummary}");
                        System.IO.File.AppendAllText(logFile, $"DUPLICATION CHECK at {placePt}:\n{sleeveSummary}\n");
                        var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placePt, tol);
                        var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placePt, tol);
                        if (indivDup.Any() || clusterDup.Any())
                        {
                            string msg = $"SKIP: Duct {duct.Id} host {hostType} {hostId} duplicate sleeve (individual or cluster) exists near {placePt}";
                            DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                            skippedCount++;
                            continue;
                        }
                        double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                        double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                        try
                        {
                            FamilySymbol? symbolToUse = null;
                            if (hostElem is Wall)
                                symbolToUse = ductWallSymbol;
                            else if (hostElem is Floor)
                                symbolToUse = ductSlabSymbol;
                            else if (hostElem is FamilyInstance fi && fi.StructuralType == StructuralType.Beam)
                                symbolToUse = ductWallSymbol;
                            else
                                symbolToUse = ductWallSymbol; // fallback
                            if (symbolToUse == null)
                            {
                                string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} no suitable family symbol found.";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                errorCount++;
                                continue;
                            }
                            var placer = new DuctSleevePlacer(doc);
                            placer.PlaceDuctSleeve(duct, placePt, w, h2, ductLine.Direction, symbolToUse, hostElem);
                            string msg2 = $"PLACED: Duct {duct.Id} host {hostType} {hostId} at {placePt} size=({w},{h2})";
                            DebugLogger.Log($"[DuctSleeveCommand] {msg2}");
                            System.IO.File.AppendAllText(logFile, msg2 + System.Environment.NewLine);
                            placedCount++;
                        }
                        catch (Exception ex)
                        {
                            string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} failed to place sleeve: {ex.Message}";
                            DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                            errorCount++;
                        }
                    }
                    if (intersections.Count == 0)
                    {
                        string msg = $"SKIP: Duct {duct.Id} no intersection with any structural element";
                        DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                        System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                    }
                }
                tx.Commit();
            }
            // Write summary to dedicated duct sleeve log file
            string summary = $"DUCT SLEEVE SUMMARY: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}";
            System.IO.File.AppendAllText(logFile, summary + System.Environment.NewLine);
            DebugLogger.Log($"[DuctSleeveCommand] Wrote duct summary log to: {logFile}");
            DebugLogger.Log(summary);
            return Result.Succeeded;
        }
    }
}