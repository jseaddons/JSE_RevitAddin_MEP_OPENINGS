using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DuctSleeveCommand : IExternalCommand
    {
        private ClearanceValues? _uiClearances; // Cached clearance values

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData, ref message, elements, null);
        }

        /// <summary>
        /// Execute command logic directly without ExternalCommandData (for orchestrator use)
        /// </summary>
        public Result ExecuteImpl(UIApplication uiApp)
        {
            try
            {
                UIDocument uidoc = uiApp.ActiveUIDocument;
                Document doc = uidoc.Document;
                string message = "";
                ElementSet elements = new ElementSet();
                
                return Execute(uidoc, doc, ref message, elements, null);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleeveCommand] ExecuteImpl failed: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result Execute(UIDocument uidoc, Document doc, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
        {
            return ExecuteCore(uidoc, doc, ref message, elements, filteredDucts);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            return ExecuteCore(uidoc, doc, ref message, elements, filteredDucts);
        }

    private Result ExecuteCore(UIDocument uidoc, Document doc, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
    {
        // Create separate timestamped log file for DuctSleeveCommand activities
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        string ductSleeveLogPath = $@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\DuctSleeve_{timestamp}.log";

        try
        {
            // Ensure directory exists
            string logDir = Path.GetDirectoryName(ductSleeveLogPath) ?? @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log";
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // Write initial log entry to separate file
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] === DUCT SLEEVE COMMAND STARTED ===\n");
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] Log file: DuctSleeve_{timestamp}.log\n");
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] Document: {doc.Title}\n");
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] Active View: {doc.ActiveView?.Name ?? "Unknown"}\n");
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] Filtered ducts provided: {filteredDucts != null}\n");
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[DuctSleeveCommand] Failed to create separate log file: {ex.Message}");
            // Continue with command even if logging fails
        }

        // Helper method to log to both separate file and DebugLogger
        void LogToFile(string message)
        {
            try
            {
                File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] {message}\n");
            }
            catch { /* Ignore file logging errors */ }
            DebugLogger.Info($"[DuctSleeveCommand] {message}");
        }

        // Enable comprehensive logging for DuctSleeveCommand activities
        DebugLogger.SetServiceContext("DuctSleeveCommand");
        LogToFile("=== DUCT SLEEVE COMMAND STARTED ===");
        LogToFile($"Document: {doc.Title}");
        LogToFile($"Active View: {doc.ActiveView?.Name ?? "Unknown"}");
        LogToFile($"Filtered ducts provided: {filteredDucts != null}");

            LogToFile("Loading duct sleeve family symbols...");

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

            LogToFile($"Wall symbol found: {ductWallSymbol != null} (ID: {ductWallSymbol?.Id.IntegerValue ?? 0})");
            LogToFile($"Slab symbol found: {ductSlabSymbol != null} (ID: {ductSlabSymbol?.Id.IntegerValue ?? 0})");

            if (ductWallSymbol == null && ductSlabSymbol == null)
            {
                LogToFile("ERROR: No duct sleeve family symbols found");
                TaskDialog.Show("Error", "Please load both wall and slab duct sleeve opening families.");
                return Result.Failed;
            }

            LogToFile("Activating duct sleeve family symbols...");
            using (var txActivate = new Transaction(doc, "Activate Duct Symbols"))
            {
                txActivate.Start();
                if (ductWallSymbol != null && !ductWallSymbol.IsActive)
                {
                    ductWallSymbol.Activate();
                    LogToFile($"Activated wall symbol: {ductWallSymbol.Name}");
                }
                if (ductSlabSymbol != null && !ductSlabSymbol.IsActive)
                {
                    ductSlabSymbol.Activate();
                    LogToFile($"Activated slab symbol: {ductSlabSymbol.Name}");
                }
                txActivate.Commit();
            }
            LogToFile("Family symbol activation completed");

            List<(Duct, Transform?)> ductTuples;
            if (filteredDucts != null)
            {
                ductTuples = filteredDucts;
                LogToFile($"Using filtered ducts: {ductTuples.Count}");
            }
            else
            {
                var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
                ductTuples = mepElements
                    .Where(tuple => tuple.Item1 is Duct && tuple.Item1 != null)
                    .Select(tuple => ((Duct)tuple.Item1, tuple.Item2))
                    .ToList();
                LogToFile($"Collected ducts from MEP elements: {ductTuples.Count}");
            }

            if (ductTuples.Count == 0)
            {
                LogToFile("No ducts found - command will exit");
                // Removed TaskDialog.Show to avoid interrupting user workflow
                return Result.Succeeded;
            }

            var structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, m => { });
            LogToFile($"Collected {structuralElements.Count} structural elements");

            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                // Enable logging for DuctSleeveCommand debugging
                void Log(string m) { DebugLogger.Info($"[DuctSleeveCommand] {m}"); }

                // 1. Read UI clearance values once at command start
                _uiClearances = GetUIClearanceValues();
                Log($"UI clearances: {_uiClearances}");

                // 2. Run the placer with the cached values
                var placerService = new DuctSleevePlacerService(
                    doc,
                    ductTuples,
                    structuralElements,
                    ductWallSymbol!,
                    ductSlabSymbol!,
                    Log,
                    _uiClearances
                );
                placerService.PlaceAllDuctSleeves();
                tx.Commit();

                string summary = $"DUCT SLEEVE SUMMARY: Placed={placerService.PlacedCount}, Skipped={placerService.SkippedCount}, Errors={placerService.ErrorCount}";
                LogToFile($"{summary}");
                // Removed TaskDialog.Show to avoid interrupting user workflow
            }

            LogToFile("=== DUCT SLEEVE COMMAND COMPLETED SUCCESSFULLY ===");
            return Result.Succeeded;
        }
        
        /// <summary>
        /// Get UI clearance values once at command start - optimized pattern
        /// </summary>
        private ClearanceValues GetUIClearanceValues()
        {
            try
            {
                // Get clearance values from ClearanceManager (one-time read)
                var uiClearances = ClearanceManager.Instance.GetUIClearances();

                // Convert to immutable value object
                var clearanceValues = ClearanceValues.FromDictionary(uiClearances);

                // Note: We can't use LogToFile here because it's defined in ExecuteCore
                // So we use DebugLogger directly for this method
                DebugLogger.Info($"[DuctSleeveCommand] Retrieved UI clearances: {clearanceValues}");
                return clearanceValues;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleeveCommand] Error getting UI clearance values: {ex.Message}");
                // Return defaults on error
                return new ClearanceValues();
            }
        }
    }
}
