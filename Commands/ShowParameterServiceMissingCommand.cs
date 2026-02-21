using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Placeholder command shown when Parameter Service DLL is not found.
    /// Informs user to build the JSE_Parameter_Service project.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ShowParameterServiceMissingCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Parameter Service Not Found", 
                "The Parameter Service module (JSE_Parameter_Service.dll) was not found.\n\n" +
                "To enable Parameter Service functionality:\n" +
                "1. Build the JSE_Parameter_Service project\n" +
                "2. Copy JSE_Parameter_Service.dll to the same folder as this add-in\n" +
                "3. Restart Revit\n\n" +
                "Alternatively, use the integrated Parameter Transfer features in the main add-in.");
            
            return Result.Succeeded;
        }
    }

    /// <summary>
    /// Availability class that disables the button when Parameter Service is not found.
    /// </summary>
    public class ParameterServiceMissingAvailability : IExternalCommandAvailability
    {
        public bool IsCommandAvailable(UIApplication applicationData, CategorySet selectedCategories)
        {
            // Always return false to disable the button - it's just a placeholder
            // The button serves as a visual indicator that Parameter Service is not installed
            return true; // Keep it enabled so user can click and see the message
        }
    }
}
