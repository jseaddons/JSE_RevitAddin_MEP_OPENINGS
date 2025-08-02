using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System.IO;

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
            var pipeTuples = mepElements
                .Where(tuple => tuple.Item1 is Pipe)
                .Select(tuple => ((Pipe)tuple.Item1, tuple.Item2))
                .ToList();
            if (pipeTuples.Count == 0)
            {
                TaskDialog.Show("Info", "No pipes found in host or linked models.");
                return Result.Succeeded;
            }

            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                    fi.Symbol.Family.Name.Equals("PipeOpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                    fi.Symbol.Family.Name.Equals("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase) ||
                    fi.Symbol.Family.Name.Equals("ClusterOpeningOnWallX", StringComparison.OrdinalIgnoreCase) ||
                    fi.Symbol.Family.Name.Equals("ClusterOpeningOnSlab", StringComparison.OrdinalIgnoreCase))
                .ToList();

            int placedCount = 0, skippedCount = 0, errorCount = 0;
              
            
            // --- Insert here ---
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "JSE_RevitAddin_Logs",
                $"PipeSleeve_{DateTime.Now:yyyyMMdd_HHmmss}.log"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
            
            void Log(string msg)
            {
                File.AppendAllText(logPath, $"{DateTime.Now:HH:mm:ss} {msg}\n");
            }
            // --- End insert ---
            
            
            var structuralElements = PipeSleeveIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc);
            // Ensure structuralElements is List<(Element, Transform?)>
            // If the method returns List<Element>, convert it:
            // var structuralElements = PipeSleeveIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc)
            //     .Select(e => (e, (Transform?)null)).ToList();


            using (var tx = new Transaction(doc, "Place Pipe Sleeves"))
            {
                tx.Start();
                var placerService = new PipeSleevePlacerService(
                    doc,
                    pipeTuples,
                    structuralElements,
                    pipeWallSymbol!,
                    pipeSlabSymbol!,
                    existingSleeves,
                    Log
                );
                placerService.PlaceAllPipeSleeves();
                placedCount = placerService.PlacedCount;
                skippedCount = placerService.SkippedCount;
                errorCount = placerService.ErrorCount;
                tx.Commit();
                Log($"Placement complete. Placed: {placedCount}, Skipped: {skippedCount}, Errors: {errorCount}");

            }
            return Result.Succeeded;
        }
    }
}