using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Duct-specific placement strategy
    /// Handles round and rectangular ducts with insulation-aware clearance
    /// </summary>
    public class DuctPlacementStrategy : ISleevePlacementStrategy
    {
        public MepElementSize GetMepElementSize(Element mepElement, System.Collections.Generic.Dictionary<string, string>? parameters = null)
        {
            var duct = mepElement as Duct;
            if (duct == null)
            {
                DebugLogger.Warning($"[DuctStrategy] Element {mepElement?.Id} is not a Duct");
                return new MepElementSize();
            }
            
            var size = new MepElementSize();
            
            // Check if round
            if (TryGetDoubleParameter(duct, parameters, "Diameter", out double diameter) && diameter > 0.001)
            {
                size.Shape = "Round";
                size.Diameter = diameter;
                size.Width = diameter;
                size.Height = diameter;
            }
            else
            {
                // Rectangular
                size.Shape = "Rectangular";
                TryGetDoubleParameter(duct, parameters, "Width", out double w);
                TryGetDoubleParameter(duct, parameters, "Height", out double h);
                size.Width = w;
                size.Height = h;
            }
            
            // Check insulation
            if (TryGetDoubleParameter(mepElement, parameters, "Insulation Thickness", out double thickness))
            {
                double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: Insulation thickness = {thicknessMm:F1}mm ({thickness:F6}ft)");
                
                if (thickness > 0.001) 
                {
                    size.IsInsulated = true;
                    size.InsulationThickness = thickness;
                    DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: ✅ INSULATED");
                }
            }
            
            // ⚠️ DIAGNOSTIC: Log final insulation status for XML storage
            DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: FINAL STATUS - Shape='{size.Shape}', IsInsulated={size.IsInsulated}, InsulationThickness={size.InsulationThickness:F6}ft");
            
            return size;
        }

        private bool TryGetDoubleParameter(Element element, System.Collections.Generic.Dictionary<string, string>? parameters, string name, out double value)
        {
            value = 0;
            if (parameters != null && parameters.TryGetValue(name, out var strVal) && !string.IsNullOrEmpty(strVal))
            {
                if (double.TryParse(strVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value))
                    return true;
            }
            
            var p = element.LookupParameter(name);
            if (p != null && p.HasValue)
            {
                value = p.AsDouble();
                return true;
            }
            return false;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            if (conditions?.ClearanceSettings == null)
            {
                // Default: 50mm
                DebugLogger.Info($"[DuctStrategy] Clearance: Using default 50mm (conditions or ClearanceSettings is null)");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
            
            double clearanceMm;
            
            // Duct-specific: Different clearance for round vs rectangular, normal vs insulated
            if (mepSize.Shape == "Round" || mepSize.Shape == "Circular")
            {
                clearanceMm = mepSize.IsInsulated 
                    ? conditions.ClearanceSettings.RoundInsulated 
                    : conditions.ClearanceSettings.RoundNormal;
            }
            else
            {
                clearanceMm = mepSize.IsInsulated 
                    ? conditions.ClearanceSettings.RectangularInsulated 
                    : conditions.ClearanceSettings.RectangularNormal;
            }
            
            // ✅ FIX: Check if clearance is 0.0 or invalid (database may have NULL values that return 0.0)
            // If clearance is 0.0 or less, fall back to default values
            if (clearanceMm <= 0.0 || double.IsNaN(clearanceMm) || double.IsInfinity(clearanceMm))
            {
                DebugLogger.Warning($"[DuctStrategy] Clearance: Invalid value {clearanceMm}mm from database, using default 50mm");
                clearanceMm = 50.0; // Default clearance for ducts
            }
            
            DebugLogger.Info($"[DuctStrategy] Clearance: Shape={mepSize.Shape}, Insulated={mepSize.IsInsulated}, Clearance={clearanceMm}mm");
            
            return UnitUtils.ConvertToInternalUnits(clearanceMm, UnitTypeId.Millimeters);
        }
        
        public string GetSystemAbbreviation(Element mepElement, System.Collections.Generic.Dictionary<string, string>? parameters = null)
        {
            if (parameters != null && parameters.TryGetValue("System Abbreviation", out var abbr) && !string.IsNullOrEmpty(abbr))
            {
                return abbr;
            }
            var duct = mepElement as Duct;
            if (duct?.MEPSystem != null)
            {
                var systemAbbrParam = duct.MEPSystem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                return systemAbbrParam?.AsString() ?? "HVAC";
            }
            return "HVAC";
        }
        
        public string GetCategoryName()
        {
            return "Ducts";
        }
    }
}

