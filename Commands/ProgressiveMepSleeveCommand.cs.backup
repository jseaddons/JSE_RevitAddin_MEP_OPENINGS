using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ProgressiveMepSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Initialize logging
                string absoluteLogDir = Path.Combine("C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23", "Log");
                string absoluteLogPath = Path.Combine(absoluteLogDir, $"ProgressiveMepSleeve_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                Services.DebugLogger.InitAbsoluteLogFile(absoluteLogPath);
                
                void Log(string msg) => Services.DebugLogger.Info(msg);
                
                Log("=== PROGRESSIVE MEP SLEEVE PLACEMENT STARTED ===");
                Log($"Active view: {doc.ActiveView?.Name ?? "Unknown"}");
                
                // Check if we're in a 3D view with section box
                if (!(doc.ActiveView is View3D view3D))
                {
                    TaskDialog.Show("Error", "Please switch to a 3D view before running this command.");
                    return Result.Failed;
                }
                
                if (!view3D.IsSectionBoxActive)
                {
                    var result = TaskDialog.Show("Warning", 
                        "No active section box detected. This command works best with a section box to limit the search area.\n\n" +
                        "Do you want to continue anyway?", 
                        TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);
                    
                    if (result != TaskDialogResult.Yes)
                    {
                        return Result.Cancelled;
                    }
                }
                
                Log($"3D view confirmed. Section box active: {view3D.IsSectionBoxActive}");

                TaskDialog.Show("Info", "ProgressiveMepSleeveCommand is deprecated. Use the Parameter Service (universal placer) workflow.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Failed to initialize progressive MEP sleeve placement: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
