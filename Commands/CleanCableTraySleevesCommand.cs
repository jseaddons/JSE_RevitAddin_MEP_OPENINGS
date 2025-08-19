using System;
using System.Linq;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CleanCableTraySleevesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }
            Document doc = uidoc.Document;

            try
            {
                // Find all cable tray sleeves in the active document
                var sleeves = OpeningDuplicationChecker.FindCableTraySleeves(doc);
                if (sleeves == null || !sleeves.Any())
                {
                    TaskDialog.Show("Clean Cable Tray Sleeves", "No cable tray sleeves found in the active document.");
                    return Result.Succeeded;
                }

                // Build a pre-delete summary using existing summary helper for each sleeve location
                double summaryTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                var sb = new StringBuilder();
                sb.AppendLine($"Found {sleeves.Count} cable tray sleeve(s). Pre-delete summary:");

                foreach (var fi in sleeves)
                {
                    var lp = fi.Location as LocationPoint;
                    if (lp == null)
                    {
                        sb.AppendLine($" - Sleeve ID:{fi.Id.IntegerValue} has no LocationPoint; Family: {fi.Symbol?.Name}");
                        continue;
                    }
                    var loc = lp.Point;
                    string summary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, loc, summaryTol);
                    sb.AppendLine($" - Sleeve ID:{fi.Id.IntegerValue} at {loc}:\n    {summary.Replace("\n", "\n    ")}");
                }

                // Confirm deletion quickly by showing count (lightweight confirmation)
                var confirm = TaskDialog.Show("Clean Cable Tray Sleeves", sb.ToString(), TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);
                if (confirm != TaskDialogResult.Ok)
                    return Result.Cancelled;

                // Delete all found sleeves in one transaction
                using (var tx = new Transaction(doc, "Clean Cable Tray Sleeves"))
                {
                    if (tx.Start() != TransactionStatus.Started)
                    {
                        TaskDialog.Show("Clean Cable Tray Sleeves", "Failed to start transaction.");
                        return Result.Failed;
                    }
                    var ids = sleeves.Select(s => s.Id).ToList();
                    doc.Delete(ids);
                    tx.Commit();
                }

                // Post-delete verification
                var remaining = OpeningDuplicationChecker.FindCableTraySleeves(doc);
                int remainingCount = remaining?.Count ?? 0;
                TaskDialog.Show("Clean Cable Tray Sleeves", $"Deleted {sleeves.Count} sleeve(s). Remaining cable tray sleeves: {remainingCount}.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Clean Cable Tray Sleeves - Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}
