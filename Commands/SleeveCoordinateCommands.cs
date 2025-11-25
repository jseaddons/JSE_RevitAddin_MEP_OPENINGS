using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Command to log all sleeve coordinates from the model
    /// </summary>
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class LogSleeveCoordinatesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                
                var coordinateService = new SleeveCoordinateService(doc);
                coordinateService.LogAllSleeveCoordinates();
                
                TaskDialog.Show("Success", "Sleeve coordinates logged successfully!\n\nCheck: Log\\all_sleeve_coordinates.log");
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Exception: {ex.Message}");
                return Result.Failed;
            }
        }
    }
    
    /// <summary>
    /// ✅ OBSOLETE: Command to update sleeve coordinates - now database-only
    /// This command is deprecated as XML is obsolete. Coordinate updates happen automatically during placement.
    /// </summary>
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Obsolete("XML is obsolete - coordinate updates happen automatically during placement (database-only mode)")]
    public class UpdateSleeveCoordinatesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                
                TaskDialog.Show("Obsolete Command", 
                    "This command is obsolete.\n\n" +
                    "Coordinate updates now happen automatically during sleeve placement.\n" +
                    "All data is stored in the database (XML is obsolete).\n\n" +
                    "If you need to update coordinates for a specific category, use the placement command instead.");
                
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Exception: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}












