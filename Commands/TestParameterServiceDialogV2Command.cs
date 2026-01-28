using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Test command to open the new Parameter Service Dialog V2
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TestParameterServiceDialogV2Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var document = commandData.Application.ActiveUIDocument.Document;
                var uiDocument = commandData.Application.ActiveUIDocument;

                // DEPRECATED: Parameter Service UI is disabled.
                TaskDialog.Show("Deprecated", "The Parameter Service V2 is currently deprecated and disabled.");
                return Result.Succeeded;
                /*
                // Open the new UI
                using (var dialog = new ParameterServiceDialogV2(document, uiDocument))
                {
                    dialog.ShowDialog();
                }

                return Result.Succeeded;
                */
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Error", $"Failed to open Parameter Service V2:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}

