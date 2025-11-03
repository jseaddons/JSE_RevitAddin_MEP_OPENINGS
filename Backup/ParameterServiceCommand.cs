using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to launch the standalone Parameter Service dialog
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ParameterServiceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc?.Document;

                // Launch the Parameter Service dialog
                using (var parameterServiceDialog = new ParameterServiceDialog(doc, uiDoc))
                {
                    parameterServiceDialog.ShowDialog();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error launching Parameter Service: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}





