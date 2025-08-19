using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DuctSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData, ref message, elements, null);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

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

            var structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, m => { });

            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                // minimal/no-op logger: focus on core logic (no file logging)
                void Log(string m) { }

                var placerService = new DuctSleevePlacerService(
                    doc,
                    ductTuples,
                    structuralElements,
                    ductWallSymbol!,
                    ductSlabSymbol!,
                    Log
                );
                placerService.PlaceAllDuctSleeves();
                tx.Commit();

                string summary = $"DUCT SLEEVE SUMMARY: Placed={placerService.PlacedCount}, Skipped={placerService.SkippedCount}, Errors={placerService.ErrorCount}";
                TaskDialog.Show("Duct Sleeve Placement", summary);
            }

            return Result.Succeeded;
        }
    }
}