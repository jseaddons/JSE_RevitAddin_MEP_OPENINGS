using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CombinedSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // DEBUG: Alert start
            TaskDialog.Show("Debug", "CombinedSleeveCommand Started");

            // Feature 20: Crash Safe Execution
                var crashSafeExecutor = new CrashSafeExecutor();
                string resultMsg = "";
                var result = crashSafeExecutor.ExecuteWithTimeout(() =>
                {
                    try
                    {
                        // Initialize Document and Services
                        var uidoc = commandData.Application.ActiveUIDocument;
                        var doc = uidoc.Document;

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[CombinedSleeveCommand] Starting Combined Sleeve UI...");

                        TaskDialog.Show("Debug", "Creating External Event...");
                        
                        // Create Event Handler
                        var handler = new CombinedSleeveRequestHandler();
                        // IMPORTANT: ExternalEvent can fail if not in valid context, but we are in IExternalCommand here.
                        var externalEvent = Autodesk.Revit.UI.ExternalEvent.Create(handler);

                        if (externalEvent == null) TaskDialog.Show("Error", "ExternalEvent creation failed!");

                        TaskDialog.Show("Debug", "Creating ViewModel...");

                        // Create ViewModel logic
                        var viewModel = new CombinedSleeveViewModel(uidoc, externalEvent, handler);

                        TaskDialog.Show("Debug", "Creating Window...");

                        // Create and Show Window
                        var window = new CombinedSleeveWindow(viewModel);
                        
                        // FIX: Set Owner logic to prevent window hiding behind Revit
                        try 
                        {
                            var handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                            new System.Windows.Interop.WindowInteropHelper(window).Owner = handle;
                            TaskDialog.Show("Debug", $"Window Owner Set to {handle}");
                        }
                        catch (Exception winEx)
                        {
                             TaskDialog.Show("Warning", $"Could not set window owner: {winEx.Message}");
                        }

                        window.Show();

                        TaskDialog.Show("Debug", "Window Shown!");

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
                        // VITAL: Show error to user since command fails silently otherwise
                        TaskDialog.Show("Error", $"Command Failed: {ex.Message}\n{ex.StackTrace}");
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[CombinedSleeveCommand] Failed to launch UI: {ex.Message}");
                        resultMsg = ex.Message;
                        return Result.Failed;
                    }
                }, "CombinedSleeveCommand");
                
                if (result == Result.Failed && !string.IsNullOrEmpty(resultMsg))
                {
                    message = resultMsg;
                }
                return result;
        }
    }
}
