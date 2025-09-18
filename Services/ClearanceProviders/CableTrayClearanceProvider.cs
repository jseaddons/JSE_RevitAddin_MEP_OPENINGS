using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Specialized clearance provider for Cable Trays
    /// Handles different clearances for top side vs other sides
    /// </summary>
    public class CableTrayClearanceProvider : IClearanceProvider
    {
        public double GetClearance(Element mepElement, Dictionary<string, double> uiClearances = null)
        {
            if (mepElement is CableTray cableTray)
            {
                // Check UI overrides first
                if (uiClearances != null)
                {
                    // Try to get top clearance first
                    if (uiClearances.TryGetValue("cabletray_top_normal", out double topClearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(topClearance, UnitTypeId.Millimeters);
                    }
                    if (uiClearances.TryGetValue("cabletray_top_insulated", out double topInsulatedClearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(topInsulatedClearance, UnitTypeId.Millimeters);
                    }
                    
                    // Fallback to other sides clearance
                    if (uiClearances.TryGetValue("cabletray_other_normal", out double otherClearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(otherClearance, UnitTypeId.Millimeters);
                    }
                    if (uiClearances.TryGetValue("cabletray_other_insulated", out double otherInsulatedClearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(otherInsulatedClearance, UnitTypeId.Millimeters);
                    }
                }
                
                // UI should always provide default values, but if somehow missing, use reasonable defaults
                return UnitUtils.ConvertToInternalUnits(75.0, UnitTypeId.Millimeters); // Default top clearance
            }
            
            return UnitUtils.ConvertToInternalUnits(75.0, UnitTypeId.Millimeters); // Default top clearance
        }

        /// <summary>
        /// Get clearance for a specific side of the cable tray
        /// </summary>
        /// <param name="mepElement">The cable tray element</param>
        /// <param name="side">The side: "top", "other", or "all"</param>
        /// <param name="uiClearances">UI clearance settings</param>
        /// <returns>Clearance value in internal units</returns>
        public double GetClearanceForSide(Element mepElement, string side, Dictionary<string, double> uiClearances = null)
        {
            if (mepElement is CableTray cableTray)
            {
                // Check UI overrides first
                if (uiClearances != null)
                {
                    string key = side.ToLower() switch
                    {
                        "top" => "cabletray_top_normal",
                        "other" => "cabletray_other_normal",
                        _ => "cabletray_other_normal" // Default to other sides
                    };
                    
                    if (uiClearances.TryGetValue(key, out double clearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(clearance, UnitTypeId.Millimeters);
                    }
                }
                
                // UI should always provide default values, but if somehow missing, use reasonable defaults
                return side.ToLower() switch
                {
                    "top" => UnitUtils.ConvertToInternalUnits(75.0, UnitTypeId.Millimeters), // Default top clearance
                    "other" => UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters), // Default other sides clearance
                    _ => UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters)
                };
            }
            
            return UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters); // Default other sides clearance
        }

        public string GetCategory()
        {
            return "Cable Trays";
        }

        public bool SupportsInsulationDetection()
        {
            return false; // Cable trays don't typically have insulation in standard Revit
        }

        public string GetClearanceKey(Element mepElement)
        {
            // Return the primary key for cable tray clearance
            return "cabletray_top_normal";
        }
    }
}
