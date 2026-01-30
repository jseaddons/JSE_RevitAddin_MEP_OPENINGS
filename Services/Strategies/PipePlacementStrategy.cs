using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;

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
            
            // ⚠️ DIAGNOSTIC: Log insulation detection for debugging
            DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: Checking insulation parameters");
            DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: InsulationThickness param = {insulationParam?.AsDouble() ?? -1:F6}ft");
            
            if (insulationParam != null && insulationParam.HasValue)
            {
                double thickness = insulationParam.AsDouble();
                double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: Insulation thickness = {thicknessMm:F1}mm ({thickness:F6}ft)");
                
                if (thickness > 0.001) // > ~0.3mm
                {
                    size.IsInsulated = true;
                    size.InsulationThickness = thickness;
                    DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: ✅ INSULATED (thickness > 0.3mm)");
                }
                else
                {
                    DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: ❌ NOT INSULATED (thickness ≤ 0.3mm)");
                }
            }
            else
            {
                DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: ❌ NOT INSULATED (no insulation parameter)");
            }
            
            // ⚠️ DIAGNOSTIC: Log final insulation status for XML storage
            DebugLogger.Info($"[PipeStrategy] Pipe {mepElement.Id}: FINAL STATUS - Shape='{size.Shape}', IsInsulated={size.IsInsulated}, InsulationThickness={size.InsulationThickness:F6}ft");
            
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
        
        /// <summary>
        /// Determines the opening type for a pipe based on global configuration rules
        /// Implements the architecture: Global rules > UI preferences
        /// </summary>
        /// <param name="mepSize">MEP element size information</param>
        /// <param name="uiPreference">UI preference for opening type</param>
        /// <param name="clearanceMm">Clearance to include in diameter calculation (mm)</param>
        /// <returns>Resolved opening type: "Circular" or "Rectangular"</returns>
        public string GetResolvedOpeningType(MepElementSize mepSize, string uiPreference = "Circular", string hostType = null, double clearanceMm = 50.0)
        {
            try
            {
                // Create element properties for configuration resolution
                var elementProps = new ElementProperties
                {
                    Diameter = mepSize.Diameter,
                    Width = mepSize.Width,
                    Height = mepSize.Height,
                    Shape = mepSize.Shape,
                    IsInsulated = mepSize.IsInsulated,
                    InsulationThickness = mepSize.InsulationThickness,
                    Angle = 0.0 // Pipes typically don't have angle issues
                };
                
                // Create UI preferences
                var uiPreferences = new UIUserPreferences
                {
                    OpeningType = uiPreference,
                    Clearance = clearanceMm // Use provided clearance
                };
                
                // Resolve configuration using global rules
                var resolvedConfig = ConfigurationResolutionService.Instance
                    .ResolveConfiguration("Pipes", elementProps, uiPreferences, hostType);
                
                // Recalculate opening type directly (ResolveConfiguration calls ResolveOpeningType, but we want to ensure clearance is passed)
                // Actually ResolveConfiguration creates a ResolvedConfiguration object which calls ResolveOpeningType without clearance?
                // Wait, ResolveConfiguration method in ConfigurationResolutionService calls ResolveOpeningType.
                // I need to update ConfigurationResolutionService.ResolveConfiguration to pass clearance too!
                
                // Direct call for opening type with clearance
                return ConfigurationResolutionService.Instance.ResolveOpeningType("Pipes", elementProps, uiPreference, hostType, clearanceMm);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PipeStrategy] Error resolving opening type: {ex.Message}");
                return uiPreference; // Fallback to UI preference
            }
        }
    }
}
