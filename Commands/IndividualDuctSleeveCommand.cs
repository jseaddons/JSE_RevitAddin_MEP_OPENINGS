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
    public class IndividualDuctSleeveCommand : IExternalCommand
    {
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

            // Collect all ducts
            var ducts = new FilteredElementCollector(doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>()
                .ToList();

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
            var structuralElements = DuctSleeveIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc);
            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                foreach (var duct in ducts)
                {
                    var locCurve = duct.Location as LocationCurve;
                    var ductLine = locCurve?.Curve as Line;
                    if (ductLine == null)
                    {
                        string msg = $"SKIP: Duct {duct.Id} is not a line";
                        DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg}");
                        System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                        continue;
                    }
                    string ductMsg = $"PROCESSING: Duct {duct.Id} Line Start={ductLine.GetEndPoint(0)}, End={ductLine.GetEndPoint(1)}";
                    DebugLogger.Log($"[IndividualDuctSleeveCommand] {ductMsg}");
                    System.IO.File.AppendAllText(logFile, ductMsg + System.Environment.NewLine);

                    // Use bounding box intersection for sleeve placement
                    var intersections = DuctSleeveIntersectionService.FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(duct, structuralElements);
                    foreach (var tuple in intersections)
                    {
                        Element hostElem = tuple.Item1;
                        BoundingBoxXYZ bbox = tuple.Item2;
                        XYZ placePt = tuple.Item3;
                        string hostType = hostElem?.GetType().Name ?? "null";
                        string hostId = hostElem?.Id.Value.ToString() ?? "null";
                        string hostMsg = $"HOST: Duct {duct.Id} intersects {hostType} {hostId} BBox=({bbox.Min},{bbox.Max})";
                        DebugLogger.Log($"[IndividualDuctSleeveCommand] {hostMsg}");
                        System.IO.File.AppendAllText(logFile, hostMsg + System.Environment.NewLine);
                        // Suppression: skip if sleeve already exists nearby
                        double tol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                        if (existingSleeves.Any(slv => (slv.Location as LocationPoint)?.Point.DistanceTo(placePt) < tol))
                        {
                            string msg = $"SKIP: Duct {duct.Id} host {hostType} {hostId} sleeve already exists near {placePt}";
                            DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg}");
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                            skippedCount++;
                            continue;
                        }
                        // Get duct dimensions
                        double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                        double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                        double sw = w;
                        double sh = h2;
                        try
                        {
                            // Select correct family symbol based on host type
                            FamilySymbol? symbolToUse = null;
                            bool isFloor = false;
                            if (hostElem is Wall || hostElem is FamilyInstance fi && fi.StructuralType == StructuralType.Beam)
                            symbolToUse = ductWallSymbol;
                            else if (hostElem is Floor)
                            {
                                symbolToUse = ductSlabSymbol;
                                isFloor = true;
                            }
                            else
                                symbolToUse = ductWallSymbol; // fallback for other types
                            if (symbolToUse == null)
                            {
                                string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} no suitable family symbol found.";
                                DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg}");
                                System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                                errorCount++;
                                continue;
                            }
                            var placer = new DuctSleevePlacer(doc);
                            // Place the sleeve, passing raw dimensions. The placer will add clearance.
                            placer.PlaceDuctSleeve(duct, placePt, w, h2, ductLine.Direction, symbolToUse, hostElem);
                            string msg2 = $"PLACED: Duct {duct.Id} host {hostType} {hostId} at {placePt} size=({sw},{sh})";
                            DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg2}");
                            System.IO.File.AppendAllText(logFile, msg2 + System.Environment.NewLine);
                            placedCount++;
                        }
                        catch (Exception ex)
                        {
                            string msg = $"ERROR: Duct {duct.Id} host {hostType} {hostId} failed to place sleeve: {ex.Message}";
                            DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg}");
                            System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                            errorCount++;
                        }
                    }
                    if (intersections.Count == 0)
                    {
                        string msg = $"SKIP: Duct {duct.Id} no intersection with any structural element";
                        DebugLogger.Log($"[IndividualDuctSleeveCommand] {msg}");
                        System.IO.File.AppendAllText(logFile, msg + System.Environment.NewLine);
                        skippedCount++;
                    }
                }
                tx.Commit();
            }
            // Write summary to dedicated duct sleeve log file
            string summary = $"DUCT SLEEVE SUMMARY: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}";
            System.IO.File.AppendAllText(logFile, summary + System.Environment.NewLine);
            DebugLogger.Log($"[IndividualDuctSleeveCommand] Wrote duct summary log to: {logFile}");
            DebugLogger.Log(summary);
            return Result.Succeeded;
        }
    }
}
