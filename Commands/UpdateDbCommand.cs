using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// updates DB from Model (Sleeves)
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class UpdateDbCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var doc = commandData.Application.ActiveUIDocument.Document;

            // Confirm
            var result = WinForms.MessageBox.Show(
                "This will synchronize the Database with the current Model state.\n\n" +
                "• Updates dimensions/locations of existing sleeves.\n" +
                "• Adds manually placed sleeves to the database.\n\n" + 
                "Continue?", 
                "Update DB from Model",
                WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Question);

            if (result != WinForms.DialogResult.Yes) return Result.Cancelled;

            try
            {
                var service = new UpdateDbService(msg => System.Diagnostics.Debug.WriteLine(msg));
                
                // We need a Transaction because we might write GUIDs to new manual sleeves
                using (var t = new Transaction(doc, "Update DB from Model"))
                {
                    t.Start();
                    int count = service.UpdateDbFromRevit(doc);
                    t.Commit();
                    
                    TaskDialog.Show("Success", $"Database Updated.\nProcessed {count} sleeves.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
