using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class LogCableTraySleevePositionsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("CableTrayOpeningOnWall") || fi.Symbol.Family.Name.Contains("CableTrayOpeningOnSlab"))
                .ToList();

            DebugLogger.Log($"--- CableTray Sleeve Position Diagnostic ---");
            DebugLogger.Log($"Found {sleeves.Count} cable tray sleeves in the active document.");
            foreach (var sleeve in sleeves)
            {
                var loc = sleeve.Location as LocationPoint;
                XYZ? pt = loc?.Point;
                double width = sleeve.LookupParameter("Width")?.AsDouble() ?? 0.0;
                double height = sleeve.LookupParameter("Height")?.AsDouble() ?? 0.0;
                double depth = sleeve.LookupParameter("Depth")?.AsDouble() ?? 0.0;
                DebugLogger.Log($"Sleeve ID: {sleeve.Id.IntegerValue}, Family: {sleeve.Symbol.Family.Name}, Symbol: {sleeve.Symbol.Name}, Position: {pt?.ToString() ?? "N/A"}, Width: {UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height: {UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, Depth: {UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm");
            }
            DebugLogger.Log($"--- End CableTray Sleeve Position Diagnostic ---");
            TaskDialog.Show("CableTray Sleeve Diagnostic", $"Logged {sleeves.Count} cable tray sleeves to DebugLogger.");
            return Result.Succeeded;
        }
    }
}
