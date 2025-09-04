using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// MINIMAL TEST - NO UI - ZERO CRASH RISK
    /// Just tests if basic services work without any UI
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class MinimalTestCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // MINIMAL TEST - NO UI AT ALL
                TaskDialog.Show("Minimal Test", "Command started successfully - NO UI!");
                
                // Test basic service access
                var appProfileService = ApplicationProfileService.Instance;
                
                // Update profile service for current document
                var doc = commandData.Application.ActiveUIDocument.Document;
                var docPath = doc.PathName;
                appProfileService.UpdateForCurrentDocument(docPath);
                
                TaskDialog.Show("Service Test", $"ApplicationProfileService created: {appProfileService != null}\nDocument: {docPath}");
                
                // Test profile status
                var isSetupRequired = appProfileService?.IsProfileSetupRequired ?? false;
                TaskDialog.Show("Profile Test", $"IsProfileSetupRequired: {isSetupRequired}");
                
                if (appProfileService.CurrentProfile != null)
                {
                    TaskDialog.Show("Current Profile", $"Current Profile: {appProfileService.CurrentProfile.Name}");
                }
                else
                {
                    TaskDialog.Show("Current Profile", "No current profile");
                }
                
                TaskDialog.Show("Success", "All tests passed - NO UI CRASHES!");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Exception: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
