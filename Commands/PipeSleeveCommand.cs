using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class PipeSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var pipeWallSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Equals("PipeOpeningOnWall", StringComparison.OrdinalIgnoreCase));
            var pipeSlabSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Equals("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase));
            if (pipeWallSymbol == null && pipeSlabSymbol == null)
            {
                TaskDialog.Show("Error", "Please load both wall and slab pipe sleeve opening families (PS#).");
                return Result.Failed;
            }
            using (var txActivate = new Transaction(doc, "Activate Pipe Symbols"))
            {
                txActivate.Start();
                if (pipeWallSymbol != null && !pipeWallSymbol.IsActive)
                    pipeWallSymbol.Activate();
                if (pipeSlabSymbol != null && !pipeSlabSymbol.IsActive)
                    pipeSlabSymbol.Activate();
                txActivate.Commit();
            }

            // Collect pipes from both host and visible linked models
            var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var pipes = mepElements
                .Where(tuple => tuple.Item1 is Pipe)
                .Select(tuple => tuple.Item1 as Pipe)
                .Where(p => p != null)
                .ToList();
            if (pipes.Count == 0)
            {
                TaskDialog.Show("Info", "No pipes found in host or linked models.");
                return Result.Succeeded;
            }

            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                    fi.Symbol.Family.Name.Equals("PipeOpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                    fi.Symbol.Family.Name.Equals("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase))
                .ToList();

            int placedCount = 0, skippedCount = 0, errorCount = 0;
            var structuralElements = PipeSleeveIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc);
            var placer = new PipeSleevePlacer(doc);

            using (var tx = new Transaction(doc, "Place Pipe Sleeves"))
            {
                tx.Start();
                foreach (var pipe in pipes)
                {
                    if (pipe == null) { skippedCount++; continue; }
                    var locCurve = pipe.Location as LocationCurve;
                    var pipeLine = locCurve?.Curve as Line;
                    if (pipeLine == null)
                    {
                        skippedCount++;
                        continue;
                    }

                    var intersections = PipeSleeveIntersectionService.FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(pipe, structuralElements);
                    if (intersections != null && intersections.Count > 0)
                    {
                        foreach (var tuple in intersections)
                        {
                            Element hostElem = tuple.Item1;
                            XYZ intersectionPoint = tuple.Item3; // Always use intersection point for placement

                            double indivTol = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters); // 5mm for individual sleeves
                            double clusterTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm for clusters (unchanged)
                            // --- Robust duplicate suppression and logging (like Duct/CableTray) ---
                            string hostId = hostElem?.Id.Value.ToString() ?? "null";
                            string hostType = hostElem?.GetType().Name ?? "null";
                            string hostMsg = $"HOST: Pipe {pipe.Id} intersects {hostType} {hostId}";
                            DebugLogger.Log($"[PipeSleeveCommand] {hostMsg}");
                            string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, intersectionPoint, clusterTol);
                            DebugLogger.Log($"[PipeSleeveCommand] DUPLICATION CHECK at {intersectionPoint}:\n{sleeveSummary}");
                            var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, intersectionPoint, indivTol);
                            var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, intersectionPoint, clusterTol);
                            if (clusterDup.Any())
                            {
                                string msg = $"SKIP: Pipe {pipe.Id} host {hostType} {hostId} duplicate cluster sleeve exists near {intersectionPoint}";
                                DebugLogger.Log($"[PipeSleeveCommand] {msg}");
                                skippedCount++;
                                continue;
                            }

                            try
                            {
                                FamilySymbol? symbolToUse = null;
                                // Robust framing detection: match duct/cable tray logic
                                if (hostElem is Floor)
                                {
                                    symbolToUse = pipeSlabSymbol;
                                    DebugLogger.Log($"[PipeSleeveCommand] Host is Floor. Using pipeSlabSymbol.");
                                }
                                else if (hostElem is Wall)
                                {
                                    symbolToUse = pipeWallSymbol;
                                    DebugLogger.Log($"[PipeSleeveCommand] Host is Wall. Using pipeWallSymbol.");
                                }
                                else if (hostElem is FamilyInstance fi && fi.Category != null && fi.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                                {
                                    symbolToUse = pipeWallSymbol;
                                    DebugLogger.Log($"[PipeSleeveCommand] Host is Structural Framing. Using pipeWallSymbol.");
                                }
                                else
                                {
                                    symbolToUse = pipeWallSymbol; // fallback
                                    DebugLogger.Log($"[PipeSleeveCommand] Host is other type ({hostType}). Using pipeWallSymbol as fallback.");
                                }

                                if (symbolToUse == null)
                                {
                                    DebugLogger.Log($"[PipeSleeveCommand] ERROR: symbolToUse is null for host {hostType} {hostId}.");
                                    errorCount++;
                                    continue;
                                }

                                placer.PlaceSleeve(pipe, intersectionPoint, pipeLine.Direction, symbolToUse, hostElem);
                                placedCount++;
                                DebugLogger.Log($"[PipeSleeveCommand] Placed sleeve for pipe {pipe.Id} in host {hostType} {hostId}.");
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Log($"[PipeSleeveCommand] ERROR placing sleeve for pipe {pipe.Id} in host {hostType} {hostId}: {ex.Message}");
                                errorCount++;
                            }
                        }
                    }
                    else
                    {
                        skippedCount++;
                    }
                }
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}
