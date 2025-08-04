using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using Autodesk.Revit.DB.Electrical;  // for CableTray
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpeningsPLaceCommand : IExternalCommand
    {
        private void PlaceRectangularSleeveClusterV2(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing RectangularSleeveClusterCommandV2...");
                var rectClusterCommandV2 = new RectangularSleeveClusterCommandV2();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = rectClusterCommandV2.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Rectangular sleeve clustering V2 failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Rectangular sleeve clustering V2 completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error with rectangular sleeve clustering V2: {ex.Message}");
            }
        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Initialize logging for the main command
            DebugLogger.InitLogFile("OpeningsPLaceCommand");
            DebugLogger.Log("OpeningsPLaceCommand: Execute started - running all sleeve placement commands");
            
            // Initialize structural elements logger for comprehensive logging
            StructuralElementLogger.InitializeLogger();
            DebugLogger.Log($"StructuralElementLogger initialized. Log file: {StructuralElementLogger.GetLogFilePath()}");
            DebugLogger.Log($"StructuralElementLogger status: {(StructuralElementLogger.IsLoggerInitialized() ? "READY" : "FAILED")}");
            
            // TEST: Create a test log entry to verify the logger is working
            StructuralElementLogger.LogStructuralElement("SYSTEM", new Autodesk.Revit.DB.ElementId(0L), "OPENINGS PLACE COMMAND STARTED", "Master command executing all sleeve placement commands with structural support");
            StructuralElementLogger.LogStructuralElement("DIAGNOSTIC", new Autodesk.Revit.DB.ElementId(999L), "LOGGER_TEST", "This is a test entry to verify logging is working");
            
            // Show detailed diagnostic information
            string diagnosticInfo = $"DIAGNOSTIC INFO:\n" +
                                   $"- Logger Status: {(StructuralElementLogger.IsLoggerInitialized() ? "INITIALIZED" : "FAILED")}\n" +
                                   $"- Log File Path: {StructuralElementLogger.GetLogFilePath()}\n" +
                                   $"- Log File Exists: {(System.IO.File.Exists(StructuralElementLogger.GetLogFilePath()) ? "YES" : "NO")}\n" +
                                   $"- Project Directory Check: {System.IO.Directory.Exists(@"c:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS")}\n" +
                                   $"- Logs Directory Check: {System.IO.Directory.Exists(@"c:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Logs")}";
            
            DebugLogger.Log(diagnosticInfo);
            DebugLogger.Log($"Structural Logger Diagnostic: {diagnosticInfo}");
            
            // Initialization code
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Cleanup: Delete all zero-size sleeves before placement
            int deletedSleeves = SleeveCleanupHelper.DeleteZeroSizeSleeves(doc);
            if (deletedSleeves > 0)
            {
                DebugLogger.Log($"Deleted {deletedSleeves} zero-size sleeves before placement.");
            }

            // OOP: Collect and filter ducts by section box, then pass to PlaceDuctSleeves
            DebugLogger.Log("Starting duct sleeve placement...");
            var allMepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var allDucts = allMepElements
                .Where(tuple => tuple.Item1 is Autodesk.Revit.DB.Mechanical.Duct)
                .Select(tuple => ((Autodesk.Revit.DB.Mechanical.Duct)tuple.Item1, tuple.Item2))
                .ToList();
            var filteredDucts = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.FilterElementsBySectionBox(
                commandData.Application.ActiveUIDocument,
                allDucts.Select(d => (d.Item1 as Element, d.Item2)).ToList()
            )
            .Select(t => ((Autodesk.Revit.DB.Mechanical.Duct)t.element, t.transform)).ToList();
            PlaceDuctSleeves(commandData, doc, filteredDucts);
            
            DebugLogger.Log("Starting damper sleeve placement...");
            PlaceDamperSleeves(commandData, doc);
            
            DebugLogger.Log("Starting cable tray sleeve placement...");
            PlaceCableTraySleeves(commandData, doc);
            
            DebugLogger.Log("Starting pipe sleeve placement...");
            PlacePipeSleeves(commandData, doc);
            
            DebugLogger.Log("Starting rectangular pipe clustering...");
            PlaceRectangularPipeOpenings(commandData, doc); // Enabled to allow PipeOpeningsRectCommand to run
            
            DebugLogger.Log("Starting rectangular sleeve clustering...");
            PlaceRectangularSleeveClusterV2(commandData, doc);

            DebugLogger.Log("OpeningsPLaceCommand: All sleeve commands completed");
            
            // Log completion of structural elements processing
            StructuralElementLogger.LogStructuralElement("SYSTEM", new Autodesk.Revit.DB.ElementId(0L), "OPENINGS PLACE COMMAND COMPLETED", $"All sleeve placement commands finished. Structural log file: {StructuralElementLogger.GetLogFilePath()}");
            DebugLogger.Log($"OpeningsPLaceCommand: All sleeve placement commands completed. Structural elements log available at: {StructuralElementLogger.GetLogFilePath()}");
            
            TaskDialog.Show("All Sleeve Commands Completed", 
                "All sleeve placement commands have been executed successfully.\n\n" +
                "Check the individual log files for detailed results:\n" +
                "- DuctSleeveCommand.log\n" +
                "- FireDamperPlaceCommand.log (or DamperLogger.log)\n" +
                "- CableTraySleeveCommand.log (or RevitAddin_Debug.log)\n" +
                "- PipeSleeveCommand.log (or RevitAddin_Debug.log)\n" +
                "- PipeOpeningsRectCommand.log (rectangular pipe clustering)\n" +
                "- RectangularSleeveClusterCommandV2.log (rectangular sleeve clustering)\n" +
                "- OpeningsPLaceCommand.log\n\n" +
                "Features:\n" +
                "• Accurate reporting: Each command shows actual placed vs. skipped counts\n" +
                "• Automatic suppression: Skips ALL elements that already have sleeves (pipes, ducts, cable trays, dampers)\n" +
                "• Pipe exclusions: Steep inclines (>15°) and fitting clusters (3+ fittings) are excluded\n" +
                "• Rectangular clustering: Groups nearby circular sleeves into efficient rectangular openings\n" +
                "• Safe re-execution: Can run multiple times without creating duplicates\n" +
                "• Comprehensive logging: All decisions and exclusions are logged");
            return Result.Succeeded;
        }

        private void PlaceDuctSleeves(ExternalCommandData commandData, Document doc, List<(Autodesk.Revit.DB.Mechanical.Duct, Autodesk.Revit.DB.Transform?)>? filteredDucts = null)
        {
            try
            {
                var ductCommand = new DuctSleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = ductCommand.Execute(commandData, ref message, elements, filteredDucts);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Duct sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Duct sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing duct sleeves: {ex.Message}");
            }
        }

        private void PlaceDamperSleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing FireDamperPlaceCommand...");
                var damperCommand = new FireDamperPlaceCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = damperCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Damper sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Damper sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing damper sleeves: {ex.Message}");
            }
        }

        private void PlaceCableTraySleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing CableTraySleeveCommand...");
                var cableTrayCommand = new CableTraySleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = cableTrayCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Cable tray sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Cable tray sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing cable tray sleeves: {ex.Message}");
            }
        }

        private void PlacePipeSleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing PipeSleeveCommand...");
                var pipeCommand = new PipeSleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = pipeCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Pipe sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Pipe sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing pipe sleeves: {ex.Message}");
            }
        }

        private void PlaceRectangularPipeOpenings(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing PipeOpeningsRectCommand...");
                var pipeRectCommand = new PipeOpeningsRectCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = pipeRectCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Rectangular pipe opening placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Rectangular pipe opening placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing rectangular pipe openings: {ex.Message}");
            }
        }

        // Remove legacy RectangularSleeveClusterCommand usage
        // private void PlaceRectangularSleeveCluster(ExternalCommandData commandData, Document doc)
        // {
        //     try
        //     {
        //         DebugLogger.Log("Executing RectangularSleeveClusterCommand...");
        //         var rectClusterCommand = new RectangularSleeveClusterCommand();
        //         string message = "";
        //         ElementSet elements = new ElementSet();
        //         var result = rectClusterCommand.Execute(commandData, ref message, elements);
        //         if (result != Result.Succeeded)
        //         {
        //             DebugLogger.Log($"Rectangular sleeve clustering failed: {message}");
        //         }
        //         else
        //         {
        //             DebugLogger.Log("Rectangular sleeve clustering completed successfully");
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         DebugLogger.Log($"Error with rectangular sleeve clustering: {ex.Message}");
        //     }
        // }
    }
}