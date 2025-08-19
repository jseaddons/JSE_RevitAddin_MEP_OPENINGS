using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DeletePipeSleevesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            int deletedCount = SleeveCleanupHelper.DeletePipeSleeves(doc);

            TaskDialog.Show("Pipe Sleeve Cleanup", $"{deletedCount} pipe sleeves have been deleted.");

            return Result.Succeeded;
        }
    }
}
