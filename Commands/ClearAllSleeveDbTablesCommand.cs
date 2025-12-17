using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ClearAllSleeveDbTablesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                using (var ctx = new SleeveDbContext(commandData.Application.ActiveUIDocument?.Document))
                {
                    ctx.ClearAllTables();
                }
                TaskDialog.Show("DB Cleared", "All sleeve database tables have been cleared.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"Failed to clear DB: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
