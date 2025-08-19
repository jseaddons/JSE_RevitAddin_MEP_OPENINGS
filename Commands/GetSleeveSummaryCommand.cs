using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class GetSleeveSummaryCommand : IExternalCommand
    {        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            XYZ location = new XYZ(62.332725748, -5.980708094, 62.906824147);
            double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);

            string summary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, location, tolerance);

            TaskDialog.Show("Sleeve Summary", summary);

            return Result.Succeeded;
        }
    }
}
