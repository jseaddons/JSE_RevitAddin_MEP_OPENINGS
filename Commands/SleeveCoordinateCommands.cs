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
    /// Command to update sleeve coordinates in XML with correct coordinates from model
    /// </summary>
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class UpdateSleeveCoordinatesCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                
                var coordinateService = new SleeveCoordinateService(doc);
                coordinateService.UpdateSleeveCoordinatesInXml();
                
                TaskDialog.Show("Success", "Sleeve coordinates updated successfully!\n\nCheck: Log\\coordinate_update.log");
                
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












