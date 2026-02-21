using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Strategy interface for category-specific sleeve placement logic
    /// Isolates the 10% differences between MEP categories while keeping 90% common code
    /// </summary>
    public interface ISleevePlacementStrategy
    {
        /// <summary>
        /// Extract size information from MEP element
        /// </summary>
        MepElementSize GetMepElementSize(Element mepElement, System.Collections.Generic.Dictionary<string, string>? parameters = null);
        
        /// <summary>
        /// Calculate clearance based on MEP size and category-specific rules
        /// Different for: Ducts (insulation-aware), Pipes (simple), Cable Trays (no addition), Dampers (no addition)
        /// </summary>
        double GetClearance(MepElementSize mepSize, OpeningConditions conditions);
        
        /// <summary>
        /// Get system abbreviation from MEP element
        /// </summary>
        string GetSystemAbbreviation(Element mepElement, System.Collections.Generic.Dictionary<string, string>? parameters = null);
        
        /// <summary>
        /// Get MEP category name
        /// </summary>
        string GetCategoryName();
    }
}

