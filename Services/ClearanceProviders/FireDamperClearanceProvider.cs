using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Clearance provider for fire dampers with special MSFD logic
    /// Supports UI overrides while maintaining existing FireDamperSleevePlacerService logic
    /// </summary>
    public class FireDamperClearanceProvider : IClearanceProvider
    {
        public double GetClearance(Element mepElement, Dictionary<string, double>? uiClearances = null)
        {
            if (mepElement is FamilyInstance damper)
            {
                // Use existing FireDamperSleevePlacerService logic
                string familyTypeName = damper.Symbol.Name;
                string typeNameUpper = familyTypeName?.Trim().ToUpperInvariant() ?? "";
                bool isMSFD = typeNameUpper.Contains("MSFD");
                bool isMSD = typeNameUpper.Contains("MSD");
                // ✅ FIX: Check for standalone "MS" (not MSD or MSFD)
                bool isMS = typeNameUpper.Contains("MS") && !typeNameUpper.Contains("MSD") && !typeNameUpper.Contains("MSFD");
                bool isMD = typeNameUpper.Contains("MD");
                // ✅ FIX: Check for both "MOTORIZED" (US spelling) and "MOTORISED" (British spelling)
                bool isMotorized = typeNameUpper.Contains("MOTORIZED") || typeNameUpper.Contains("MOTORISED");
                
                // MD dampers should get MEP side clearance (same as MSFD/MSD/MS)
                bool needsMepSideClearance = isMSFD || isMSD || isMS || isMD || isMotorized;
                
                // Check UI overrides first
                if (uiClearances != null)
                {
                    string clearanceKey = GetClearanceKey(damper);
                    if (uiClearances.TryGetValue(clearanceKey, out double uiClearance))
                    {
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(uiClearance);
                    }
                }
                
                // UI should always provide default values, but if somehow missing, use reasonable defaults
                return needsMepSideClearance ? 
                    RevitUnitConversionService.Instance.ToInternalMillimeters(100.0) : // MSFD/MSD/MD/Motorized: 100mm MEP side clearance
                    RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);    // Standard: 50mm
            }
            
            return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0); // Default standard damper clearance
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
                string typeNameUpper = familyTypeName?.Trim().ToUpperInvariant() ?? "";
                bool isMSFD = typeNameUpper.Contains("MSFD");
                bool isMSD = typeNameUpper.Contains("MSD");
                // ✅ FIX: Check for standalone "MS" (not MSD or MSFD)
                bool isMS = typeNameUpper.Contains("MS") && !typeNameUpper.Contains("MSD") && !typeNameUpper.Contains("MSFD");
                bool isMD = typeNameUpper.Contains("MD");
                // ✅ FIX: Check for both "MOTORIZED" (US spelling) and "MOTORISED" (British spelling)
                bool isMotorized = typeNameUpper.Contains("MOTORIZED") || typeNameUpper.Contains("MOTORISED");
                
                // MD dampers should use MEP side clearance (same as MSFD/MSD/MS)
                bool needsMepSideClearance = isMSFD || isMSD || isMS || isMD || isMotorized;
                
                return needsMepSideClearance ? "fire_damper_msfd_clearance" : "fire_damper_standard_clearance";
            }
            
            return "fire_damper_standard_clearance";
        }
    }
}
