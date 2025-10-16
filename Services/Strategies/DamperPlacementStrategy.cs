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
        private readonly Document _doc;
        
        public DamperPlacementStrategy(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }
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
            
            // Store damper type for clearance calculation
            string familyTypeName = damper.Symbol?.Name ?? "";
            string typeNameUpper = familyTypeName.Trim().ToUpperInvariant();
            
            // Detect specific damper types (case insensitive)
            if (typeNameUpper.Contains("MSFD"))
            {
                size.DamperType = "MSFD";
            }
            else if (typeNameUpper.Contains("MSD"))
            {
                size.DamperType = "MSD";
            }
            else if (typeNameUpper.Contains("MD"))
            {
                size.DamperType = "MD";
            }
            else if (typeNameUpper.Contains("MOTORIZED"))
            {
                size.DamperType = "Motorized";
            }
            else
            {
                size.DamperType = "Standard";
            }
            
            DebugLogger.Info($"[DamperStrategy] Detected damper type: '{size.DamperType}' from family: '{familyTypeName}'");
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            // Use the damper type stored in MepElementSize during GetMepElementSize
            string damperType = mepSize.DamperType ?? "";
            DebugLogger.Info($"[DamperStrategy] Damper type from MepElementSize: '{damperType}'");
            
            // Apply appropriate clearance based on damper type
            double clearanceInMm;
            if (damperType.Contains("MSFD"))
            {
                clearanceInMm = conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0; // MSFD uses MEP side clearance
                DebugLogger.Info($"[DamperStrategy] MSFD damper - using MEP side clearance: {clearanceInMm}mm");
            }
            else if (damperType.Contains("MSD") || damperType.Contains("MD") || damperType.Contains("Motorized"))
            {
                clearanceInMm = conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0; // MSD, MD, Motorized use MEP side clearance
                DebugLogger.Info($"[DamperStrategy] {(damperType.Contains("MSD") ? "MSD" : damperType.Contains("MD") ? "MD" : "Motorized")} damper - using MEP side clearance: {clearanceInMm}mm");
            }
            else
            {
                clearanceInMm = conditions?.ClearanceSettings?.DuctAccessoryOtherNormal ?? 50.0; // Standard dampers use other side clearance
                DebugLogger.Info($"[DamperStrategy] Standard damper - using other side clearance: {clearanceInMm}mm");
            }
            
            // Convert from mm to feet (Revit internal units)
            double clearanceInFeet = UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
            DebugLogger.Info($"[DamperStrategy] Final clearance: {clearanceInMm}mm = {clearanceInFeet:F6}ft");
            
            return clearanceInFeet;
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
                
                // Get clearance from OpeningConditions (loaded from UI via CONDITIONS XML)
                double otherClearanceMm = conditions.ClearanceSettings.DuctAccessoryOtherNormal;
                double mepClearanceMm = conditions.ClearanceSettings.DuctAccessoryMepNormal;
                
                double baseClearance = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                
                DebugLogger.Info($"[DamperStrategy] Clearances from conditions: MEP={mepClearanceMm}mm, Other={otherClearanceMm}mm");
                
                if (clashZone.IsMSFDDamper && !string.IsNullOrEmpty(clashZone.DamperConnectorSide))
                {
                    // MSFD Damper: Asymmetric clearance
                    double mepSideClearance = UnitUtils.ConvertToInternalUnits(mepClearanceMm, UnitTypeId.Millimeters);
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
