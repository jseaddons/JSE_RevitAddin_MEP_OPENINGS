using System;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TestFiltersLayoutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the active document
                Document doc = commandData.Application.ActiveUIDocument.Document;
                
                // Show the main dialog with the new 3-column layout
                var appProfileService = ApplicationProfileService.Instance;
                using (var dialog = new EmergencyMainDialog(appProfileService, doc))
                {
                    // Set the dialog title to indicate this is a test
                    dialog.Text = "Test: 3-Column Layout with Filters";
                    
                    // Show the dialog
                    DialogResult result = dialog.ShowDialog();
                    
                    if (result == DialogResult.OK)
                    {
                        TaskDialog.Show("Test Result", "3-Column layout test completed successfully!");
                        return Result.Succeeded;
                    }
                    else
                    {
                        TaskDialog.Show("Test Cancelled", "3-Column layout test was cancelled.");
                        return Result.Cancelled;
                    }
                }
            }
            catch (Exception ex)
            {
                message = $"Error testing 3-column layout: {ex.Message}";
                TaskDialog.Show("Test Error", $"Error testing 3-column layout:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
