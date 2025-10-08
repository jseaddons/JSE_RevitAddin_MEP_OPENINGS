using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Damper (Duct Accessory) placement strategy
    /// Handles fire dampers and other duct accessories with NO clearance addition (uses actual size)
    /// </summary>
    public class DamperPlacementStrategy : ISleevePlacementStrategy
    {
        public MepElementSize GetMepElementSize(Element mepElement)
        {
            var damper = mepElement as FamilyInstance;
            if (damper == null || damper.Category?.Id.IntegerValue != (int)BuiltInCategory.OST_DuctAccessory)
            {
                DebugLogger.Warning($"[DamperStrategy] Element {mepElement?.Id} is not a Duct Accessory");
                return new MepElementSize();
            }
            
            var size = new MepElementSize { Shape = "Rectangular" }; // Most dampers are rectangular
            
            // Get size from damper-specific parameters (not generic Width/Height)
            // Fire dampers use "Damper Width" and "Damper Height" parameters
            var widthParam = damper.LookupParameter("Damper Width") ?? 
                            damper.LookupParameter("Width") ?? 
                            damper.LookupParameter("width");
            var heightParam = damper.LookupParameter("Damper Height") ?? 
                             damper.LookupParameter("Height") ?? 
                             damper.LookupParameter("height");
            
            size.Width = widthParam?.AsDouble() ?? 0.0;
            size.Height = heightParam?.AsDouble() ?? 0.0;
            
            DebugLogger.Info($"[DamperStrategy] Damper {damper.Id}: Width={size.Width} ft, Height={size.Height} ft (from '{widthParam?.Definition.Name}' and '{heightParam?.Definition.Name}' parameters)");
            
            // Dampers don't have insulation
            size.IsInsulated = false;
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            // ⚠️ DAMPER SPECIAL: NO clearance addition
            // Opening size = exact damper size (no addition)
            // Fire dampers must fit precisely in fire-rated assemblies
            
            DebugLogger.Info($"[DamperStrategy] Clearance: 0mm (dampers use exact size for fire rating compliance)");
            
            return 0.0;  // No clearance addition for dampers
        }
        
        public string GetSystemAbbreviation(Element mepElement)
        {
            var damper = mepElement as FamilyInstance;
            if (damper != null)
            {
                // Try to get connected duct system
                var connectors = damper.MEPModel?.ConnectorManager?.Connectors;
                if (connectors != null)
                {
                    foreach (Connector connector in connectors)
                    {
                        if (connector.MEPSystem != null)
                        {
                            var systemAbbrParam = connector.MEPSystem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
                            if (systemAbbrParam != null)
                            {
                                return systemAbbrParam.AsString() ?? "HVAC";
                            }
                        }
                    }
                }
            }
            
            return "HVAC"; // Default for duct accessories
        }
        
        public string GetCategoryName()
        {
            return "Duct Accessories";
        }
        
        /// <summary>
        /// Calculate damper-specific placement adjustment for MSFD dampers
        /// Returns (offsetVector, finalWidth, finalHeight)
        /// </summary>
        public (XYZ offsetVector, double finalWidth, double finalHeight) GetDamperPlacementAdjustment(
            ClashZone clashZone, 
            OpeningConditions conditions)
        {
            try
            {
                double damperWidth = clashZone.MepElementWidth;
                double damperHeight = clashZone.MepElementHeight;
                
                // Get clearance from ClearanceManager (reads from UI, same as old FireDamperPlaceCommand)
                // We can't access the actual damper element, so use a default clearance approach
                // The clearance values should be in the UI clearance dictionary passed through conditions
                double baseClearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // 50mm default
                
                if (clashZone.IsMSFDDamper && !string.IsNullOrEmpty(clashZone.DamperConnectorSide))
                {
                    // MSFD Damper: Asymmetric clearance
                    // MEP side (connector) = 2x base clearance, other sides = base clearance
                    double mepSideClearance = baseClearance * 2.0;
                    double otherSideClearance = baseClearance;
                    
                    // Calculate offset toward connector side
                    double offsetAmount = (mepSideClearance - otherSideClearance) / 2.0;
                    
                    // Determine which dimension gets the asymmetric clearance
                    double left = otherSideClearance, right = otherSideClearance;
                    double top = otherSideClearance, bottom = otherSideClearance;
                    
                    XYZ offsetVector = XYZ.Zero;
                    
                    switch (clashZone.DamperConnectorSide)
                    {
                        case "Left":
                            left = mepSideClearance;
                            offsetVector = new XYZ(-offsetAmount, 0, 0); // Offset left
                            break;
                        case "Right":
                            right = mepSideClearance;
                            offsetVector = new XYZ(offsetAmount, 0, 0); // Offset right
                            break;
                        case "Top":
                            top = mepSideClearance;
                            offsetVector = new XYZ(0, 0, offsetAmount); // Offset up
                            break;
                        case "Bottom":
                            bottom = mepSideClearance;
                            offsetVector = new XYZ(0, 0, -offsetAmount); // Offset down
                            break;
                    }
                    
                    double finalWidth = damperWidth + left + right;
                    double finalHeight = damperHeight + top + bottom;
                    
                    DebugLogger.Info($"[DamperStrategy] MSFD - Connector={clashZone.DamperConnectorSide}, MEP={mepSideClearance:F4}ft, Other={otherSideClearance:F4}ft, Offset={offsetAmount:F4}ft");
                    DebugLogger.Info($"[DamperStrategy] MSFD - Size: {finalWidth:F4} x {finalHeight:F4}, Offset: {offsetVector}");
                    
                    return (offsetVector, finalWidth, finalHeight);
                }
                else
                {
                    // Standard Damper: Symmetric clearance (same on all sides)
                    double finalWidth = damperWidth + (2 * baseClearance);
                    double finalHeight = damperHeight + (2 * baseClearance);
                    
                    DebugLogger.Info($"[DamperStrategy] Standard - Clearance={baseClearance:F4}ft, Size: {finalWidth:F4} x {finalHeight:F4}, No offset");
                    
                    return (XYZ.Zero, finalWidth, finalHeight);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[DamperStrategy] Error calculating placement adjustment: {ex.Message}");
                return (XYZ.Zero, clashZone.MepElementWidth, clashZone.MepElementHeight);
            }
        }
    }
}
