using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Cable Tray-specific placement strategy
    /// Handles rectangular cable trays with NO clearance addition (uses actual size)
    /// </summary>
    public class CableTrayPlacementStrategy : ISleevePlacementStrategy
    {
        public MepElementSize GetMepElementSize(Element mepElement)
        {
            var cableTray = mepElement as CableTray;
            if (cableTray == null)
            {
                DebugLogger.Warning($"[CableTrayStrategy] Element {mepElement?.Id} is not a CableTray");
                return new MepElementSize();
            }
            
            var size = new MepElementSize { Shape = "Rectangular" }; // Cable trays are always rectangular
            
            var widthParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
            var heightParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
            
            size.Width = widthParam?.AsDouble() ?? 0.0;
            size.Height = heightParam?.AsDouble() ?? 0.0;
            
            // Cable trays typically don't have insulation
            size.IsInsulated = false;
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            // ⚠️ CABLE TRAY SPECIAL: NO clearance addition
            // Opening size = exact cable tray size (no addition)
            // This is different from ducts/pipes which add clearance
            
            DebugLogger.Info($"[CableTrayStrategy] Clearance: 0mm (cable trays use exact size, no addition)");
            
            return 0.0;  // No clearance addition for cable trays
        }
        
        public string GetSystemAbbreviation(Element mepElement)
        {
            // Cable trays don't have MEPSystem like ducts/pipes
            // Default to electrical
            return "ELEC";
        }
        
        public string GetCategoryName()
        {
            return "Cable Trays";
        }
        
        /// <summary>
        /// Calculate cable tray placement adjustment (always offset upward)
        /// Returns (offsetVector, finalWidth, finalHeight)
        /// </summary>
        public (XYZ offsetVector, double finalWidth, double finalHeight) GetCableTrayPlacementAdjustment(
            ClashZone clashZone, 
            OpeningConditions conditions)
        {
            try
            {
                double trayWidth = clashZone.MepElementWidth;
                double trayHeight = clashZone.MepElementHeight;
                
                // Cable trays: top clearance is typically larger, other 3 sides smaller
                double baseClearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // 50mm default
                double topClearance = baseClearance * 2.0; // Top side gets 2x clearance (100mm default)
                double otherClearance = baseClearance; // Other 3 sides get base clearance (50mm default)
                
                // Calculate offset upward (always toward top)
                double offsetAmount = (topClearance - otherClearance) / 2.0;
                XYZ offsetVector = new XYZ(0, 0, offsetAmount); // Always offset upward
                
                // Calculate final size with asymmetric clearances
                double finalWidth = trayWidth + (2 * otherClearance); // Left and right use other clearance
                double finalHeight = trayHeight + topClearance + otherClearance; // Top uses top clearance, bottom uses other
                
                DebugLogger.Info($"[CableTrayStrategy] Top={topClearance:F4}ft, Other={otherClearance:F4}ft, Offset={offsetAmount:F4}ft upward");
                DebugLogger.Info($"[CableTrayStrategy] Size: {finalWidth:F4} x {finalHeight:F4}, Offset: {offsetVector}");
                
                return (offsetVector, finalWidth, finalHeight);
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[CableTrayStrategy] Error calculating placement adjustment: {ex.Message}");
                return (XYZ.Zero, clashZone.MepElementWidth, clashZone.MepElementHeight);
            }
        }
    }
}
