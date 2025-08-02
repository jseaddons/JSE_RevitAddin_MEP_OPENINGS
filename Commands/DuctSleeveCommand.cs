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

        /// <summary>
        /// Main entry point. Optionally accepts a pre-filtered list of ducts (with transforms) for OOP section box filtering.
        /// If filteredDucts is null, falls back to legacy collection logic.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
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

            // Use provided filtered ducts if available, else collect as before
            List<(Duct, Transform?)> ductTuples;
            if (filteredDucts != null)
            {
                ductTuples = filteredDucts;
            }
            else
            {
                var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
                ductTuples = mepElements
                    .Where(tuple => tuple.Item1 is Duct && tuple.Item1 != null)
                    .Select(tuple => ((Duct)tuple.Item1, tuple.Item2))
                    .ToList();
            }
            if (ductTuples.Count == 0)
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
            // Section box filter for structural elements
            var allStructuralElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.StructuralElementCollectorHelper.CollectStructuralElementsVisibleOnly(doc);
            var filteredStructuralElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.FilterElementsBySectionBox(
                commandData.Application.ActiveUIDocument,
                allStructuralElements
            );
            // Use robust bounding box intersection for floors
            var structuralElementsVisible = filteredStructuralElements;
            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                foreach (var tuple in ductTuples)
                {
                    var duct = tuple.Item1;
                    var transform = tuple.Item2;
                    if (duct == null) { skippedCount++; continue; }
                    var locCurve = duct.Location as LocationCurve;
                    var ductLine = locCurve?.Curve as Line;
                    if (ductLine == null)
                    {
                        string msg = $"SKIP: Duct {duct?.Id} is not a line";
                        DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                        if (DebugLogger.IsEnabled)
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                        continue;
                    }
                    // Transform geometry if from a linked model
                    Line hostLine = ductLine;
                    if (transform != null)
                    {
                        hostLine = Line.CreateBound(
                            transform.OfPoint(ductLine.GetEndPoint(0)),
                            transform.OfPoint(ductLine.GetEndPoint(1))
                        );
                    }
                    string ductMsg = $"PROCESSING: Duct {duct.Id} Line Start={hostLine.GetEndPoint(0)}, End={hostLine.GetEndPoint(1)}";
                    DebugLogger.Log($"[DuctSleeveCommand] {ductMsg}");
                    if (DebugLogger.IsEnabled)
                        System.IO.File.AppendAllText(logFile, ductMsg + System.Environment.NewLine);

                    // Use robust bounding box intersection for floors (pass hostLine as needed)
                    var bboxIntersections = JSE_RevitAddin_MEP_OPENINGS.Services.DuctSleeveIntersectionService.FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(duct, structuralElementsVisible, hostLine);
                    if (bboxIntersections.Count > 0)
                    {
                        foreach (var bboxTuple in bboxIntersections)
                        {
                            Element hostElem = bboxTuple.Item1;
                            BoundingBoxXYZ bbox = bboxTuple.Item2;
                            XYZ placePt = bboxTuple.Item3;
                            string hostType = hostElem?.GetType().Name ?? "null";
                            string hostId = hostElem?.Id.Value.ToString() ?? "null";
                            string hostMsg = $"HOST: Duct {duct.Id} intersects {hostType} {hostId} BBox=({bbox.Min},{bbox.Max})";
                            DebugLogger.Log($"[DuctSleeveCommand] {hostMsg}");
                            if (DebugLogger.IsEnabled)
                                System.IO.File.AppendAllText(logFile, hostMsg + System.Environment.NewLine);

                            // FLOOR-SPECIFIC DEBUG: Add detailed floor detection logging and enforce structural filter
                            if (hostElem is Floor floor)
                            {
                                var isStructuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                                bool isStructural = isStructuralParam != null && isStructuralParam.AsInteger() == 1;
                                string floorStructuralStatus = isStructural ? "STRUCTURAL" : "NON-STRUCTURAL";
                                string floorDebugMsg = $"FLOOR DEBUG: Duct {duct.Id} intersects Floor {floor.Id.Value} - Status: {floorStructuralStatus}";
                                DebugLogger.Log($"[DuctSleeveCommand] {floorDebugMsg}");
                                if (DebugLogger.IsEnabled)
                                    System.IO.File.AppendAllText(logFile, floorDebugMsg + System.Environment.NewLine);

                                if (!isStructural)
                                {
                                    string skipMsg = $"SKIP: Duct {duct.Id} host Floor {floor.Id.Value} is NON-STRUCTURAL. Sleeve will NOT be placed.";
                                    DebugLogger.Log($"[DuctSleeveCommand] {skipMsg}");
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, skipMsg + System.Environment.NewLine);
                                    skippedCount++;
                                    continue;
                                }

                                // Check floor family symbol availability
                                if (ductSlabSymbol != null)
                                {
                                    string symbolMsg = $"FLOOR SYMBOL: Found floor symbol: {ductSlabSymbol.Family.Name} - {ductSlabSymbol.Name}";
                                    DebugLogger.Log($"[DuctSleeveCommand] {symbolMsg}");
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, symbolMsg + System.Environment.NewLine);
                                }
                                else
                                {
                                    string noSymbolMsg = $"FLOOR SYMBOL ERROR: No floor sleeve symbol available for Duct {duct.Id}";
                                    DebugLogger.Log($"[DuctSleeveCommand] {noSymbolMsg}");
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, noSymbolMsg + System.Environment.NewLine);
                                }
                            }
                            // Duplicate suppression: log what sleeves are found at this location
                            double indivTol = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters); // 5mm for individual sleeves
                            double clusterTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm for clusters
                            string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placePt, clusterTol);
                            DebugLogger.Log($"[DuctSleeveCommand] DUPLICATION CHECK at {placePt}:\n{sleeveSummary}");
                            if (DebugLogger.IsEnabled)
                                System.IO.File.AppendAllText(logFile, $"DUPLICATION CHECK at {placePt}:\n{sleeveSummary}\n");
                            var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placePt, indivTol);
                            var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placePt, clusterTol);
                            if (indivDup.Any() || clusterDup.Any())
                            {
                                string msg = $"SKIP: Duct {duct.Id} host {hostType} {hostId} duplicate sleeve (individual or cluster) exists near {placePt}";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                if (DebugLogger.IsEnabled)
                                    System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                skippedCount++;
                                continue;
                            }
                            double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                            double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                            // Add clearance (insulation+25mm or 50mm) on all sides
                            double clearance = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveClearanceHelper.GetClearance(duct);
                            w = w + 2 * clearance;
                            h2 = h2 + 2 * clearance;
                            if (hostElem == null) { errorCount++; continue; }
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
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                    errorCount++;
                                    continue;
                                }
                                var placer = new DuctSleevePlacer(doc);
                                // Only place sleeve if hostElem is not null
                                if (hostElem != null)
                                {
                                    // === DEBUG LOGGING: Data passed to placer ===
                                    DebugLogger.Log($"[DuctSleeveCommand] DEBUG BEFORE PLACER: DuctId={duct.Id}, HostType={hostElem.GetType().Name}, HostId={hostElem.Id}");
                                    DebugLogger.Log($"[DuctSleeveCommand] DEBUG INTERSECTION: ({placePt.X:F6}, {placePt.Y:F6}, {placePt.Z:F6})");
                                    DebugLogger.Log($"[DuctSleeveCommand] DEBUG DIRECTION: ({hostLine.Direction.X:F6}, {hostLine.Direction.Y:F6}, {hostLine.Direction.Z:F6})");
                                    DebugLogger.Log($"[DuctSleeveCommand] DEBUG SIZE: Width={w:F6}, Height={h2:F6}");
                                    DebugLogger.Log($"[DuctSleeveCommand] DEBUG FAMILY: {symbolToUse.FamilyName} - {symbolToUse.Name}");
                                    
                                    // DETAILED DUCT ORIENTATION ANALYSIS AND PRE-CALCULATION
                                    DebugLogger.Log($"[DuctSleeveCommand] === DUCT ORIENTATION ANALYSIS ===");
                                    DebugLogger.Log($"[DuctSleeveCommand] Duct Width Parameter: {UnitUtils.ConvertFromInternalUnits(duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0, UnitTypeId.Millimeters):F1}mm");
                                    DebugLogger.Log($"[DuctSleeveCommand] Duct Height Parameter: {UnitUtils.ConvertFromInternalUnits(duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0, UnitTypeId.Millimeters):F1}mm");
                                    DebugLogger.Log($"[DuctSleeveCommand] Original ductLine.Direction: ({ductLine.Direction.X:F6}, {ductLine.Direction.Y:F6}, {ductLine.Direction.Z:F6})");
                                    DebugLogger.Log($"[DuctSleeveCommand] Transformed hostLine.Direction: ({hostLine.Direction.X:F6}, {hostLine.Direction.Y:F6}, {hostLine.Direction.Z:F6})");
                                    
                                    // Get the duct's actual width orientation from its connectors
                                    LocationCurve? ductLocation = duct.Location as LocationCurve;
                                    XYZ ductWidthDirection = XYZ.BasisY; // Default
                                    
                                    // Try to get the actual width direction from duct connectors
                                    try
                                    {
                                        ConnectorManager connectorManager = duct.ConnectorManager;
                                        if (connectorManager != null)
                                        {
                                            foreach (Connector connector in connectorManager.Connectors)
                                            {
                                                if (connector.ConnectorType == ConnectorType.End)
                                                {
                                                    // Get the connector's coordinate system
                                                    Transform connectorTransform = connector.CoordinateSystem;
                                                    if (connectorTransform != null)
                                                    {
                                                        // BasisX of connector represents width direction
                                                        ductWidthDirection = connectorTransform.BasisX;
                                                        DebugLogger.Log($"[DuctSleeveCommand] Got width direction from connector: ({ductWidthDirection.X:F3}, {ductWidthDirection.Y:F3}, {ductWidthDirection.Z:F3})");
                                                        break; // Use first valid connector
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        DebugLogger.Log($"[DuctSleeveCommand] Error getting connector orientation: {ex.Message}");
                                        // Fall back to calculation method for non-vertical ducts
                                        if (ductLocation?.Curve is Line ductLocationLine)
                                        {
                                            XYZ ductFlowDirection = ductLocationLine.Direction;
                                            if (Math.Abs(ductFlowDirection.Z) < 0.9) // Not vertical
                                            {
                                                ductWidthDirection = new XYZ(-ductFlowDirection.Y, ductFlowDirection.X, 0);
                                                if (ductWidthDirection.GetLength() > 0.001)
                                                    ductWidthDirection = ductWidthDirection.Normalize();
                                                else
                                                    ductWidthDirection = XYZ.BasisY;
                                            }
                                        }
                                    }
                                    
                                    // Check if duct width is oriented along X or Y world axis
                                    double dotY = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisY));
                                    double dotX = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisX));
                                    string orientationStatus = dotY > dotX ? "Y-ORIENTED" : "X-ORIENTED";
                                    
                                    DebugLogger.Log($"[DuctSleeveCommand] Duct width direction (world): ({ductWidthDirection.X:F3}, {ductWidthDirection.Y:F3}, {ductWidthDirection.Z:F3})");
                                    DebugLogger.Log($"[DuctSleeveCommand] ORIENTATION STATUS: {orientationStatus}");

                                    // Only rotate for Y-oriented ducts (width along Y)
                                    if (orientationStatus == "Y-ORIENTED")
                                    {
                                        DebugLogger.Log($"[DuctSleeveCommand] Passing Y-oriented duct to placer for rotation.");
                                        placer.PlaceDuctSleeveWithOrientation(duct, placePt, w, h2, hostLine.Direction, ductWidthDirection, symbolToUse, hostElem);
                                    }
                                    else
                                    {
                                        DebugLogger.Log($"[DuctSleeveCommand] Passing X-oriented duct to placer (no rotation).");
                                        placer.PlaceDuctSleeve(duct, placePt, w, h2, hostLine.Direction, symbolToUse, hostElem);
                                    }
                                    string msg2 = $"PLACED: Duct {duct.Id} host {hostType} {hostId} at {placePt} size=({w},{h2})";
                                    DebugLogger.Log($"[DuctSleeveCommand] {msg2}");
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, msg2 + System.Environment.NewLine);
                                    placedCount++;
                                }
                                else
                                {
                                    string nullHostMsg = $"SKIP: Duct {duct.Id} intersection host element is null (not placing sleeve)";
                                    DebugLogger.Log($"[DuctSleeveCommand] {nullHostMsg}");
                                    if (DebugLogger.IsEnabled)
                                        System.IO.File.AppendAllText(logFile, nullHostMsg + System.Environment.NewLine);
                                    skippedCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} failed to place sleeve: {ex.Message}";
                                DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                                if (DebugLogger.IsEnabled)
                                    System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                errorCount++;
                            }
                        }
                        continue;
                    }
                    if (bboxIntersections.Count == 0)
                    {
                        string msg = $"SKIP: Duct {duct.Id} no intersection with any structural element";
                        DebugLogger.Log($"[DuctSleeveCommand] {msg}");
                        if (DebugLogger.IsEnabled)
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                    }
                }
                tx.Commit();
            }
            // Write summary to dedicated duct sleeve log file
            string summary = $"DUCT SLEEVE SUMMARY: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}";
            if (DebugLogger.IsEnabled)
                System.IO.File.AppendAllText(logFile, summary + System.Environment.NewLine);
            DebugLogger.Log($"[DuctSleeveCommand] Wrote duct summary log to: {logFile}");
            DebugLogger.Log(summary);
            return Result.Succeeded;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData, ref message, elements, null);
        }
    }
}