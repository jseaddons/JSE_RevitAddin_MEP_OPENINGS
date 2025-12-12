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

                        // Create ViewModel logic
                        var viewModel = new CombinedSleeveViewModel(uidoc);

                        // Create and Show Window
                        var window = new CombinedSleeveWindow(viewModel);
                        window.Show();

                        return Result.Succeeded;
                    }
                    catch (Exception ex)
                    {
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
