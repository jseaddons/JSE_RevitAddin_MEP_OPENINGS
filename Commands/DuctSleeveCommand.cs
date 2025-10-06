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
    [Transaction(TransactionMode.ReadOnly)]
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
            return ExecuteCore(uidoc, doc, ref message, elements, filteredDucts, null);
        }

        /// <summary>
        /// Execute with clash zones for optimized sleeve placement
        /// </summary>
        public Result Execute(UIDocument uidoc, Document doc, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts, List<ClashZone>? clashZones)
        {
            return ExecuteCore(uidoc, doc, ref message, elements, filteredDucts, clashZones);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            return ExecuteCore(uidoc, doc, ref message, elements, filteredDucts, null);
        }

    private Result ExecuteCore(UIDocument uidoc, Document doc, ref string message, ElementSet elements, List<(Duct, Transform?)>? filteredDucts, List<ClashZone>? clashZones)
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
            File.AppendAllText(ductSleeveLogPath, $"[{DateTime.Now}] Clash zones provided: {clashZones != null} (Count: {clashZones?.Count ?? 0})\n");
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
        LogToFile($"Clash zones provided: {clashZones != null} (Count: {clashZones?.Count ?? 0})");

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

            LogToFile("Checking duct sleeve family symbol status...");
            try
            {
                // Check if symbols are already active (no transaction needed)
                bool wallActive = ductWallSymbol?.IsActive ?? false;
                bool slabActive = ductSlabSymbol?.IsActive ?? false;
                
                LogToFile($"Wall symbol status: IsActive={wallActive}, Name={ductWallSymbol?.Name ?? "null"}");
                LogToFile($"Slab symbol status: IsActive={slabActive}, Name={ductSlabSymbol?.Name ?? "null"}");
                
                // Activate symbols directly (no separate transaction needed)
                if ((!wallActive && ductWallSymbol != null) || (!slabActive && ductSlabSymbol != null))
                {
                    if (!wallActive && ductWallSymbol != null)
                    {
                        ductWallSymbol.Activate();
                        LogToFile($"Activated wall symbol: {ductWallSymbol.Name}");
                    }
                    
                    if (!slabActive && ductSlabSymbol != null)
                    {
                        ductSlabSymbol.Activate();
                        LogToFile($"Activated slab symbol: {ductSlabSymbol.Name}");
                    }
                    
                    LogToFile("Family symbol activation completed successfully");
                }
                
                LogToFile("Family symbol activation check completed");
            }
            catch (Exception ex)
            {
                LogToFile($"ERROR: Family symbol status check failed: {ex.Message}");
                LogToFile("Continuing with existing symbol status...");
            }

            List<(Duct, Transform?)> ductTuples;
            if (filteredDucts != null)
            {
                ductTuples = filteredDucts;
                LogToFile($"Using filtered ducts: {ductTuples.Count}");
            }
            else if (clashZones != null && clashZones.Count > 0)
            {
                // 🎯 OPTIMIZATION: Only process ducts that have clash zones (not all ducts in document)
                var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
                var allDuctTuples = mepElements
                    .Where(tuple => tuple.Item1 is Duct && tuple.Item1 != null)
                    .Select(tuple => ((Duct)tuple.Item1, tuple.Item2))
                    .ToList();
                
                // Filter to only ducts that have clash zones
                var clashZoneDuctIds = clashZones.Select(cz => cz.MepElementId).ToHashSet();
                ductTuples = allDuctTuples
                    .Where(tuple => clashZoneDuctIds.Contains(tuple.Item1.Id))
                    .ToList();
                
                LogToFile($"Collected ALL ducts from MEP elements: {allDuctTuples.Count}");
                LogToFile($"Filtered to ducts WITH clash zones: {ductTuples.Count}");
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

            // OPTIMIZATION: Use clash zones if provided, otherwise collect structural elements
            List<(Element, Transform?)> structuralElements;
            if (clashZones != null && clashZones.Count > 0)
            {
                LogToFile($"OPTIMIZATION: Using {clashZones.Count} clash zones instead of re-finding intersections");
                // Extract structural elements from clash zones (handle linked documents)
                structuralElements = new List<(Element, Transform?)>();
                
                foreach (var cz in clashZones)
                {
                    var structuralElement = doc.GetElement(cz.StructuralElementId);
                    
                    // If element not found in host document, it might be from a linked document
                    if (structuralElement == null && cz.StructuralElementId.IntegerValue > 0)
                    {
                        LogToFile($"Structural element {cz.StructuralElementId.IntegerValue} not found in host document - checking linked documents");
                        
                        // Try to find the element in linked documents
                        var linkedFileService = new Services.LinkedFileService();
                        var linkedFiles = linkedFileService.GetLinkedFiles(doc);
                        
                        foreach (var linkedFile in linkedFiles)
                        {
                            if (linkedFile.LinkInstance?.GetLinkDocument() != null)
                            {
                                var linkedDoc = linkedFile.LinkInstance.GetLinkDocument();
                                var linkedElement = linkedDoc.GetElement(cz.StructuralElementId);
                                
                                if (linkedElement != null)
                                {
                                    LogToFile($"Found structural element {cz.StructuralElementId.IntegerValue} in linked document: {linkedFile.FileName}");
                                    structuralElement = linkedElement;
                                    break;
                                }
                            }
                        }
                    }
                    
                    if (structuralElement != null)
                    {
                        structuralElements.Add((structuralElement, null));
                        LogToFile($"Added structural element: {structuralElement.Name} (ID: {structuralElement.Id.IntegerValue})");
                    }
                    else
                    {
                        LogToFile($"WARNING: Could not find structural element {cz.StructuralElementId.IntegerValue} in any document");
                    }
                }
                
                // Remove duplicates by ID
                structuralElements = structuralElements
                    .GroupBy(tuple => tuple.Item1.Id)
                    .Select(g => g.First())
                    .ToList();
                    
                LogToFile($"Extracted {structuralElements.Count} unique structural elements from clash zones");
            }
            else
            {
                LogToFile("No clash zones provided - collecting structural elements (slower method)");
                structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, m => { });
                LogToFile($"Collected {structuralElements.Count} structural elements");
            }

            // Enable logging for DuctSleeveCommand debugging
            void Log(string m) { DebugLogger.Info($"[DuctSleeveCommand] {m}"); }

            // 1. Read UI clearance values once at command start
            _uiClearances = GetUIClearanceValues();
            Log($"UI clearances: {_uiClearances}");

            // 2. Run the placer (using SleeveClearanceHelper for UI clearance values)
            DuctSleevePlacerService placerService;
            try
            {
                LogToFile("Creating DuctSleevePlacerService...");
                placerService = new DuctSleevePlacerService(
                    doc,
                    ductTuples,
                    structuralElements,
                    ductWallSymbol!,
                    ductSlabSymbol!,
                    Log
                );
                LogToFile("DuctSleevePlacerService created successfully");
            }
            catch (Exception placerEx)
            {
                LogToFile($"ERROR: Failed to create DuctSleevePlacerService: {placerEx.Message}");
                LogToFile($"Stack trace: {placerEx.StackTrace}");
                return Result.Failed;
            }
            
            // CRITICAL FIX: Add transaction for sleeve placement (required for Revit API)
            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                try
                {
                    tx.Start();
                    LogToFile("Transaction started for duct sleeve placement");
                    
                    // OPTIMIZATION: Pass clash zones to placer for direct intersection point usage
                    if (clashZones != null && clashZones.Count > 0)
                    {
                        LogToFile($"OPTIMIZATION: Passing {clashZones.Count} clash zones to placer for direct intersection usage");
                        placerService.PlaceAllDuctSleevesWithClashZones(clashZones);
                        LogToFile("PlaceAllDuctSleevesWithClashZones completed");
                    }
                    else
                    {
                        LogToFile("Using traditional intersection detection method");
                        placerService.PlaceAllDuctSleeves();
                        LogToFile("PlaceAllDuctSleeves completed");
                    }
                    
                    tx.Commit();
                    LogToFile("Transaction committed successfully - sleeves placed in model");
                }
                catch (Exception placementEx)
                {
                    tx.RollBack();
                    LogToFile($"ERROR: Failed to place duct sleeves: {placementEx.Message}");
                    LogToFile($"Stack trace: {placementEx.StackTrace}");
                    LogToFile("Transaction rolled back due to error");
                    return Result.Failed;
                }
            }

            string summary = $"DUCT SLEEVE SUMMARY: Placed={placerService.PlacedCount}, Skipped={placerService.SkippedCount}, Errors={placerService.ErrorCount}";
            LogToFile($"{summary}");
            // Removed TaskDialog.Show to avoid interrupting user workflow

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

                // DEBUG: Log the raw UI clearances dictionary
                DebugLogger.Info($"[DuctSleeveCommand] Raw UI clearances from ClearanceManager:");
                foreach (var kvp in uiClearances)
                {
                    DebugLogger.Info($"  {kvp.Key} = {kvp.Value}mm");
                }

                // Convert to immutable value object
                var clearanceValues = ClearanceValues.FromDictionary(uiClearances);

                // Note: We can't use LogToFile here because it's defined in ExecuteCore
                // So we use DebugLogger directly for this method
                DebugLogger.Info($"[DuctSleeveCommand] Converted clearance values: {clearanceValues}");
                DebugLogger.Info($"[DuctSleeveCommand] DuctsNormalClearance: {clearanceValues.DuctsNormalClearance}mm");
                DebugLogger.Info($"[DuctSleeveCommand] DuctsInsulatedClearance: {clearanceValues.DuctsInsulatedClearance}mm");
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
