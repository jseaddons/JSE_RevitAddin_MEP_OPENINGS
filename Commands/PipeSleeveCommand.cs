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


            // --- Initialize diagnostics logger early so collector diagnostics are recorded ---
            string absoluteLogDir = Path.Combine("C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23", "Log");
            string absoluteLogPath = Path.Combine(absoluteLogDir, $"PipeSleeve_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            Services.DebugLogger.InitAbsoluteLogFile(absoluteLogPath);
            try { Services.DebugLogger.Info("PipeSleeveCommand initialized log at: " + absoluteLogPath); } catch { }
            if (!System.IO.File.Exists(absoluteLogPath))
            {
                try
                {
                    Services.DebugLogger.InitCustomLogFileOverwrite($"PipeSleeve_{DateTime.Now:yyyyMMdd_HHmmss}");
                    Services.DebugLogger.Info("PipeSleeveCommand - absolute log path not writable, fell back to MyDocuments.");
                }
                catch { }
            }
            void Log(string msg) => Services.DebugLogger.Info(msg);

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
            
            
            var structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, Log);
            // Ensure structuralElements is List<(Element, Transform?)>
            // If the method returns List<Element>, convert it:
            // var structuralElements = PipeSleeveIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc)
            //     .Select(e => (e, (Transform?)null)).ToList();


            using (var tx = new Transaction(doc, "Place Pipe Sleeves"))
            {
                tx.Start();
                try
                {
                    // Register failure preprocessor to suppress duplicate-instance warnings
                    var fho = tx.GetFailureHandlingOptions();
                    fho.SetFailuresPreprocessor(new JSE_RevitAddin_MEP_OPENINGS.Services.DuplicateInstanceSuppressor());
                    tx.SetFailureHandlingOptions(fho);
                }
                catch { /* If API not available or fails, continue without suppressor */ }

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
