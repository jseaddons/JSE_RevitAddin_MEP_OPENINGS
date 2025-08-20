using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CleanCableTraySleevesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Collect candidate family instances that look like cable-tray sleeves
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>();

            var candidates = collector
                .Where(fi => fi?.Symbol?.Family != null)
                .Where(fi =>
                    string.Equals(fi.Symbol.Family.Name, "CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(fi.Symbol.Family.Name, "CableTrayOpeningOnSlab", StringComparison.OrdinalIgnoreCase)
                    || fi.Symbol.Family.Name.IndexOf("CableTray", StringComparison.OrdinalIgnoreCase) >= 0
                    || fi.Symbol.Family.Name.IndexOf("CT", StringComparison.OrdinalIgnoreCase) >= 0
                )
                .ToList();

            int total = candidates.Count;
            if (total == 0)
            {
                TaskDialog.Show("Clean Cable Tray Sleeves", "No cable-tray sleeve family instances found.");
                return Result.Succeeded;
            }

            // Log a small sample of the matches and their clearance (if host available)
            int sample = Math.Min(10, total);
            for (int i = 0; i < sample; ++i)
            {
                var fi = candidates[i];
                double clearance = 0.0;
                try
                {
                    if (fi.Host != null)
                        clearance = SleeveClearanceHelper.GetClearance(fi.Host);
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"CleanCableTraySleeves: failed to get clearance for instance {fi.Id.IntegerValue}: {ex.Message}");
                }
                StructuralElementLogger.LogStructuralElement("INFO", fi.Id, "CLEANUP_CANDIDATE", $"Candidate family={fi.Symbol.Family.Name}, symbol={fi.Symbol.Name}, clearance_mm={UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F2}");
            }

            var td = new TaskDialog("Clean Cable Tray Sleeves")
            {
                MainInstruction = $"Found {total} cable-tray sleeve family instances.",
                MainContent = "This will permanently delete the found family instances from the model. Do you want to proceed?",
                AllowCancellation = true
            };
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Delete all found cable-tray sleeves");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancel (do nothing)");

            var result = td.Show();
            if (result != TaskDialogResult.CommandLink1)
            {
                StructuralElementLogger.LogStructuralElement("INFO", new ElementId(0), "CLEANUP_ABORT", "User cancelled cable-tray cleanup.");
                return Result.Cancelled;
            }

            var idsToDelete = candidates.Select(fi => fi.Id).ToList();
            using (var tx = new Transaction(doc, "Delete Cable Tray Sleeves"))
            {
                tx.Start();
                doc.Delete(idsToDelete);
                tx.Commit();
            }

            StructuralElementLogger.LogStructuralElement("SUCCESS", new ElementId(0), "CLEANUP_DONE", $"Deleted {idsToDelete.Count} cable-tray sleeve instances.");
            TaskDialog.Show("Clean Cable Tray Sleeves", $"Deleted {idsToDelete.Count} cable-tray sleeve instances.");

            return Result.Succeeded;
        }
    }
}
