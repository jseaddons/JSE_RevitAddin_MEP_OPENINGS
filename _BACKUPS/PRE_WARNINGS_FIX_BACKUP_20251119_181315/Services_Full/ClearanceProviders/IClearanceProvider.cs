using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Interface for providing clearance calculations for MEP elements
    /// Uses Strategy pattern to support different clearance approaches
    /// </summary>
    public interface IClearanceProvider
    {
        /// <summary>
        /// Get clearance for a MEP element, optionally overriding with UI settings
        /// </summary>
        /// <param name="mepElement">The MEP element to calculate clearance for</param>
        /// <param name="uiClearances">Optional UI clearance settings to override defaults</param>
        /// <returns>Clearance value in internal units (feet)</returns>
        double GetClearance(Element mepElement, Dictionary<string, double>? uiClearances = null);
        
        /// <summary>
        /// Get the category this provider handles
        /// </summary>
        /// <returns>Category name</returns>
        string GetCategory();
        
        /// <summary>
        /// Whether this provider supports insulation detection
        /// </summary>
        /// <returns>True if supports insulation detection</returns>
        bool SupportsInsulationDetection();
        
        /// <summary>
        /// Get clearance key for UI settings lookup
        /// </summary>
        /// <param name="mepElement">The MEP element</param>
        /// <returns>Key for UI clearance dictionary</returns>
        string GetClearanceKey(Element mepElement);
    }
}
