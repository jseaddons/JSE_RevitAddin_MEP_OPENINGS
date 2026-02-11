using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS
{
    [Transaction(TransactionMode.Manual)]
    public class GetSelectedElementId : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Try to pick an element (including linked elements with Tab)
                Reference reference = null;
                try
                {
                    // Use PickObject with Element type - this allows Tab to cycle through linked elements
                    reference = uiDoc.Selection.PickObject(
                        Autodesk.Revit.UI.Selection.ObjectType.Element, 
                        "Select an element (Tab to cycle through linked elements)");
                }
                catch
                {
                    TaskDialog.Show("Info", "No element selected or selection cancelled.");
                    return Result.Cancelled;
                }
                
                if (reference == null)
                {
                    TaskDialog.Show("Info", "No element selected.");
                    return Result.Cancelled;
                }

                string result = "";
                
                try
                {
                    // Check if this is a linked element reference
                    if (reference.LinkedElementId != ElementId.InvalidElementId)
                    {
                        // This is a linked element - get the link instance and the element
                        Element linkInstance = doc.GetElement(reference);
                        if (linkInstance is RevitLinkInstance revitLink)
                        {
                            var linkDoc = revitLink.GetLinkDocument();
                            if (linkDoc != null)
                            {
                                var linkElement = linkDoc.GetElement(reference.LinkedElementId);
                                if (linkElement != null)
                                {
                                    result += $"Element ID: {linkElement.Id.GetIntegerValue()}\n";
                                    result += $"Element Name: {linkElement.Name}\n";
                                    result += $"Document: {linkDoc.Title}\n";
                                    result += $"Is Linked: True\n";
                                    
                                    // If it's a wall, show wall direction info
                                    if (linkElement is Wall wall)
                                    {
                                        var locationCurve = wall.Location as LocationCurve;
                                        if (locationCurve != null)
                                        {
                                            var curve = locationCurve.Curve as Line;
                                            if (curve != null)
                                            {
                                                var wallDirection = curve.Direction;
                                                var wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                                                
                                                result += $"Wall Direction: ({wallDirection.X:F3}, {wallDirection.Y:F3}, {wallDirection.Z:F3})\n";
                                                result += $"Wall Normal: ({wallNormal.X:F3}, {wallNormal.Y:F3}, {wallNormal.Z:F3})\n";
                                                
                                                // Determine wall orientation
                                                double absX = Math.Abs(wallDirection.X);
                                                double absY = Math.Abs(wallDirection.Y);
                                                string wallType = absX > absY ? "X-WALL (Horizontal)" : "Y-WALL (Vertical)";
                                                result += $"Wall Type: {wallType}\n";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    result += $"Element ID: {reference.LinkedElementId.GetIntegerValue()} (not found in linked document)\n";
                                }
                            }
                        }
                    }
                    else
                    {
                        // This is an element from the active document
                        Element selectedElement = doc.GetElement(reference);
                        if (selectedElement != null)
                        {
                            result += $"Element ID: {selectedElement.Id.GetIntegerValue()}\n";
                            result += $"Element Name: {selectedElement.Name}\n";
                            result += $"Document: {selectedElement.Document.Title}\n";
                            result += $"Is Linked: False\n";
                        }
                        else
                        {
                            result += $"Element ID: {reference.ElementId.GetIntegerValue()} (not found)\n";
                        }
                    }
                }
                catch (Exception ex)
                {
                    result += $"Error getting element: {ex.Message}\n";
                }
                
                TaskDialog.Show("Element Information", result);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
