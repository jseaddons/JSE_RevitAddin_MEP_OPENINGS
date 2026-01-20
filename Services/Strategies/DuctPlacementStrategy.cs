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
        public MepElementSize GetMepElementSize(Element mepElement)
        {
            var duct = mepElement as Duct;
            if (duct == null)
            {
                DebugLogger.Warning($"[DuctStrategy] Element {mepElement?.Id} is not a Duct");
                return new MepElementSize();
            }
            
            var size = new MepElementSize();
            
            // Check if round
            var diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (diamParam != null && diamParam.HasValue && diamParam.AsDouble() > 0.001)
            {
                size.Shape = "Round";
                size.Diameter = diamParam.AsDouble();
                size.Width = size.Diameter;
                size.Height = size.Diameter;
            }
            else
            {
                // Rectangular
                size.Shape = "Rectangular";
                var widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                
                size.Width = widthParam?.AsDouble() ?? 0.0;
                size.Height = heightParam?.AsDouble() ?? 0.0;
            }
            
            // Check insulation
            var insulationParam = mepElement.LookupParameter("Insulation Thickness") 
                               ?? mepElement.LookupParameter("InsulationThickness");
            
            // ⚠️ DIAGNOSTIC: Log insulation detection for debugging
            DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: Checking insulation parameters");
            DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: InsulationThickness param = {insulationParam?.AsDouble() ?? -1:F6}ft");
            
            if (insulationParam != null && insulationParam.HasValue)
            {
                double thickness = insulationParam.AsDouble();
                double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: Insulation thickness = {thicknessMm:F1}mm ({thickness:F6}ft)");
                
                if (thickness > 0.001) // > ~0.3mm
                {
                    size.IsInsulated = true;
                    size.InsulationThickness = thickness;
                    DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: ✅ INSULATED (thickness > 0.3mm)");
                }
                else
                {
                    DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: ❌ NOT INSULATED (thickness ≤ 0.3mm)");
                }
            }
            else
            {
                DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: ❌ NOT INSULATED (no insulation parameter)");
            }
            
            // ⚠️ DIAGNOSTIC: Log final insulation status for XML storage
            DebugLogger.Info($"[DuctStrategy] Duct {mepElement.Id}: FINAL STATUS - Shape='{size.Shape}', IsInsulated={size.IsInsulated}, InsulationThickness={size.InsulationThickness:F6}ft");
            
            return size;
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
        
        public string GetSystemAbbreviation(Element mepElement)
        {
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

