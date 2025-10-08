using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Pipe-specific placement strategy
    /// Handles round pipes (always circular) with simple clearance
    /// </summary>
    public class PipePlacementStrategy : ISleevePlacementStrategy
    {
        public MepElementSize GetMepElementSize(Element mepElement)
        {
            var pipe = mepElement as Pipe;
            if (pipe == null)
            {
                DebugLogger.Warning($"[PipeStrategy] Element {mepElement?.Id} is not a Pipe");
                return new MepElementSize();
            }
            
            var size = new MepElementSize { Shape = "Round" }; // Pipes are always round
            
            var diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
            if (diamParam != null && diamParam.HasValue)
            {
                size.Diameter = diamParam.AsDouble();
                size.Width = size.Diameter;
                size.Height = size.Diameter;
            }
            
            // Check insulation
            var insulationParam = mepElement.LookupParameter("Insulation Thickness") 
                               ?? mepElement.LookupParameter("InsulationThickness");
            if (insulationParam != null && insulationParam.HasValue)
            {
                double thickness = insulationParam.AsDouble();
                if (thickness > 0.001)
                {
                    size.IsInsulated = true;
                    size.InsulationThickness = thickness;
                }
            }
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            // Pipes: Simple clearance (same for all pipes)
            // Use round normal clearance as default for pipes
            double clearanceMm = conditions?.ClearanceSettings?.RoundNormal ?? 50.0;
            
            DebugLogger.Info($"[PipeStrategy] Clearance: {clearanceMm}mm (simple pipe clearance)");
            
            return UnitUtils.ConvertToInternalUnits(clearanceMm, UnitTypeId.Millimeters);
        }
        
        public string GetSystemAbbreviation(Element mepElement)
        {
            var pipe = mepElement as Pipe;
            if (pipe?.MEPSystem != null)
            {
                var systemAbbrParam = pipe.MEPSystem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                return systemAbbrParam?.AsString() ?? "PLB";
            }
            return "PLB";
        }
        
        public string GetCategoryName()
        {
            return "Pipes";
        }
    }
}
