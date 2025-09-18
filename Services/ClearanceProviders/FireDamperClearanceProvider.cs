using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Clearance provider for fire dampers with special MSFD logic
    /// Supports UI overrides while maintaining existing FireDamperSleevePlacerService logic
    /// </summary>
    public class FireDamperClearanceProvider : IClearanceProvider
    {
        public double GetClearance(Element mepElement, Dictionary<string, double> uiClearances = null)
        {
            if (mepElement is FamilyInstance damper)
            {
                // Use existing FireDamperSleevePlacerService logic
                string familyTypeName = damper.Symbol.Name;
                bool isMSFD = familyTypeName?.Trim().ToUpperInvariant().Contains("MSFD") ?? false;
                
                // Check UI overrides first
                if (uiClearances != null)
                {
                    string clearanceKey = GetClearanceKey(damper);
                    if (uiClearances.TryGetValue(clearanceKey, out double uiClearance))
                    {
                        return UnitUtils.ConvertToInternalUnits(uiClearance, UnitTypeId.Millimeters);
                    }
                }
                
                // UI should always provide default values, but if somehow missing, use reasonable defaults
                return isMSFD ? 
                    UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters) : // MSFD: 100mm
                    UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);    // Standard: 50mm
            }
            
            return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // Default standard damper clearance
        }
        
        public string GetCategory()
        {
            return "Fire Damper";
        }
        
        public bool SupportsInsulationDetection()
        {
            return false; // Fire dampers don't use insulation detection
        }
        
        public string GetClearanceKey(Element mepElement)
        {
            if (mepElement is FamilyInstance damper)
            {
                string familyTypeName = damper.Symbol.Name;
                bool isMSFD = familyTypeName?.Trim().ToUpperInvariant().Contains("MSFD") ?? false;
                
                return isMSFD ? "fire_damper_msfd_clearance" : "fire_damper_standard_clearance";
            }
            
            return "fire_damper_standard_clearance";
        }
    }
}
