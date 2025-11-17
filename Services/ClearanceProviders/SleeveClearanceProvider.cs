using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Clearance provider that wraps the existing SleeveClearanceHelper logic
    /// Supports UI overrides while maintaining sophisticated insulation detection
    /// </summary>
    public class SleeveClearanceProvider : IClearanceProvider
    {
        public double GetClearance(Element mepElement, Dictionary<string, double>? uiClearances = null)
        {
            // Override with UI settings if provided
            if (uiClearances != null)
            {
                string clearanceKey = GetClearanceKey(mepElement);
                
                if (uiClearances.TryGetValue(clearanceKey, out double uiClearance))
                {
                    // Convert UI clearance to internal units and use it
                    return UnitUtils.ConvertToInternalUnits(uiClearance, UnitTypeId.Millimeters);
                }
            }
            
            // UI should always provide default values, but if somehow missing, use reasonable defaults
            string category = GetMepCategory(mepElement);
            return category switch
            {
                "Ducts" => UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters), // Default duct clearance
                "Pipes" => UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters), // Default pipe clearance
                "Cable Trays" => UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters), // Default cable tray clearance
                _ => UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters) // Default clearance
            };
        }
        
        public string GetCategory()
        {
            return "Standard";
        }
        
        public bool SupportsInsulationDetection()
        {
            return true;
        }
        
        public string GetClearanceKey(Element mepElement)
        {
            string category = GetMepCategory(mepElement);
            bool isInsulated = IsInsulated(mepElement);
            string insulationType = isInsulated ? "insulated" : "normal";
            return $"{category.ToLower().Replace(" ", "_")}_{insulationType}_clearance";
        }
        
        private string GetMepCategory(Element mepElement)
        {
            return mepElement switch
            {
                Duct => "Ducts",
                Pipe => "Pipes", 
                CableTray => "Cable Trays",
                _ => "Default"
            };
        }
        
        private bool IsInsulated(Element mepElement)
        {
            // Use existing SleeveClearanceHelper logic to determine if insulated
            // This is a simplified check - the actual logic is in SleeveClearanceHelper
            if (mepElement is Duct duct)
            {
                // Check for duct insulation
                var insulation = new FilteredElementCollector(duct.Document)
                    .OfClass(typeof(DuctInsulation))
                    .Cast<DuctInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == duct.Id);
                return insulation != null;
            }
            else if (mepElement is Pipe pipe)
            {
                // Check for pipe insulation
                var insulation = new FilteredElementCollector(pipe.Document)
                    .OfClass(typeof(PipeInsulation))
                    .Cast<PipeInsulation>()
                    .FirstOrDefault(ins => ins.HostElementId == pipe.Id);
                return insulation != null;
            }
            
            // Cable trays don't have insulation in standard Revit
            return false;
        }
    }
}
