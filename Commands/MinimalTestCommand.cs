using System;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Minimal test command to isolate the crash issue
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class MinimalTestCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Create a simple WinForms dialog
                using (var testDialog = new System.Windows.Forms.Form())
                {
                    testDialog.Text = "Minimal Test Dialog";
                    testDialog.Size = new System.Drawing.Size(400, 300);
                    testDialog.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    
                    var label = new System.Windows.Forms.Label
                    {
                        Text = "This is a minimal test dialog.\nClose this and run the command again to test for crashes.",
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                    };
                    
                    testDialog.Controls.Add(label);
                    
                    // Show as modal dialog
                    testDialog.ShowDialog();
                }
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}