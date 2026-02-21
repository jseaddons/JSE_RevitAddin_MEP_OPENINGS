using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to show the Parameter Service Dialog V2
    /// Replaces the external JSE_Parameter_Service.dll command
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ShowParameterServiceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var doc = commandData.Application.ActiveUIDocument.Document;
                var uiDoc = commandData.Application.ActiveUIDocument;

                // Show the Parameter Service Dialog V2
                using (var dialog = new ParameterServiceDialogV2(doc, uiDoc))
                {
                    dialog.ShowDialog();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error opening Parameter Service: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
