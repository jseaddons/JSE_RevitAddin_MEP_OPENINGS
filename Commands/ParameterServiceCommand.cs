using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to launch the Parameter Service dialog
    /// This provides a separate way to access parameter extraction without affecting the main UI
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ParameterServiceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                DebugLogger.Info("=== PARAMETER SERVICE COMMAND STARTED ===");

                // Show parameter service dialog
                using (var parameterServiceDialog = new ParameterServiceDialog(doc, uiDoc))
                {
                    var result = parameterServiceDialog.ShowDialog();
                    
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        var extractedParameters = parameterServiceDialog.GetExtractedParameters();
                        DebugLogger.Info($"Parameter service completed successfully. Extracted {extractedParameters.Count} parameters.");
                        
                        // Show success message
                        TaskDialog.Show("Parameter Service", 
                            $"Parameter extraction completed successfully!\n\n" +
                            $"Found {extractedParameters.Count} opening parameters:\n" +
                            string.Join(", ", extractedParameters.Take(10)) + 
                            (extractedParameters.Count > 10 ? "..." : ""));
                        
                        return Result.Succeeded;
                    }
                    else
                    {
                        DebugLogger.Info("Parameter service cancelled by user");
                        return Result.Cancelled;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error in ParameterServiceCommand: {ex.Message}");
                message = $"Error launching parameter service: {ex.Message}";
                
                TaskDialog.Show("Parameter Service Error", 
                    $"An error occurred while launching the parameter service:\n\n{ex.Message}");
                
                return Result.Failed;
            }
        }
    }
}
