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
            if (insulationParam != null && insulationParam.HasValue)
            {
                double thickness = insulationParam.AsDouble();
                if (thickness > 0.001) // > ~0.3mm
                {
                    size.IsInsulated = true;
                    size.InsulationThickness = thickness;
                }
            }
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            if (conditions?.ClearanceSettings == null)
            {
                // Default: 50mm
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

