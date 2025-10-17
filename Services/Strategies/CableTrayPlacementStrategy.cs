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
            OpeningConditions conditions,
            Dictionary<string, double> uiClearanceSettings = null)
        {
            try
            {
                // 🔥 DEBUG: Log that we're entering the cable tray strategy method
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cabletray_clearance_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] GetCableTrayPlacementAdjustment START for ClashZone {clashZone.Id}\n");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cabletray_clearance_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] clearanceSettings.Count = {uiClearanceSettings?.Count ?? 0}\n");
                
                if (uiClearanceSettings != null)
                {
                    foreach (var kvp in uiClearanceSettings)
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\cabletray_clearance_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}]   Key='{kvp.Key}', Value={kvp.Value}mm\n");
                    }
                }
                
                double trayWidth = clashZone.MepElementWidth;
                double trayHeight = clashZone.MepElementHeight;
                
                // ⚠️ DIAGNOSTIC: Log cable tray dimensions from ClashZone
                double trayWidthMm = UnitUtils.ConvertFromInternalUnits(trayWidth, UnitTypeId.Millimeters);
                double trayHeightMm = UnitUtils.ConvertFromInternalUnits(trayHeight, UnitTypeId.Millimeters);
                DebugLogger.Info($"[CableTrayStrategy] CZ={clashZone.Id}: Read from XML: Width={trayWidthMm:F1}mm ({trayWidth:F6}ft), Height={trayHeightMm:F1}mm ({trayHeight:F6}ft)");
                
                // ✅ PRIORITY SYSTEM: UI Settings > XML Conditions > Default
                double topClearanceMm = 100.0; // Default fallback
                double otherClearanceMm = 50.0; // Default fallback
                
                // 1. Try UI clearance settings first (highest priority)
                if (uiClearanceSettings != null && uiClearanceSettings.Count > 0)
                {
                    // 🔥 DEBUG: Log all UI clearance settings to see what keys are available
                    DebugLogger.Info($"[CableTrayStrategy] 🔥 UI Clearance Settings ({uiClearanceSettings.Count} items):");
                    foreach (var kvp in uiClearanceSettings)
                    {
                        DebugLogger.Info($"[CableTrayStrategy] 🔥 Key: '{kvp.Key}' = {kvp.Value}mm");
                    }
                    
                    // ✅ FIX: Handle insulation-specific keys for cable trays
                    // Cable trays are typically not insulated, so use normal clearance
                    if (uiClearanceSettings.ContainsKey("cabletray_top_normal"))
                    {
                        topClearanceMm = uiClearanceSettings["cabletray_top_normal"];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI top clearance (normal): {topClearanceMm}mm");
                    }
                    else if (uiClearanceSettings.ContainsKey("cabletray_top_insulated"))
                    {
                        topClearanceMm = uiClearanceSettings["cabletray_top_insulated"];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI top clearance (insulated): {topClearanceMm}mm");
                    }
                    
                    if (uiClearanceSettings.ContainsKey("cabletray_other_normal"))
                    {
                        otherClearanceMm = uiClearanceSettings["cabletray_other_normal"];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI other clearance (normal): {otherClearanceMm}mm");
                    }
                    else if (uiClearanceSettings.ContainsKey("cabletray_other_insulated"))
                    {
                        otherClearanceMm = uiClearanceSettings["cabletray_other_insulated"];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI other clearance (insulated): {otherClearanceMm}mm");
                    }
                }
                // 2. Fallback to XML conditions (second priority)
                else if (conditions?.ClearanceSettings != null)
                {
                    topClearanceMm = conditions.ClearanceSettings.CableTrayTop;
                    otherClearanceMm = conditions.ClearanceSettings.CableTrayOther;
                    DebugLogger.Info($"[CableTrayStrategy] Using XML clearances: Top={topClearanceMm}mm, Other={otherClearanceMm}mm");
                }
                
                DebugLogger.Info($"[CableTrayStrategy] FINAL clearances: Top={topClearanceMm}mm, Other={otherClearanceMm}mm");
                DebugLogger.Info($"[CableTrayStrategy] CONDITIONS object: {conditions?.FilterName ?? "NULL"}, Category: {conditions?.Category ?? "NULL"}");
                DebugLogger.Info($"[CableTrayStrategy] CONDITIONS XML file: {conditions?.FilterName ?? "NULL"}_{conditions?.Category ?? "NULL"}_CONDITIONS.xml");
                
                double topClearance = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                double otherClearance = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                
                // Calculate offset upward (always toward top)
                double offsetAmount = (topClearance - otherClearance) / 2.0;
                XYZ offsetVector = new XYZ(0, 0, offsetAmount); // Always offset upward
                
                // Calculate final size with asymmetric clearances
                double finalWidth = trayWidth + (2 * otherClearance); // Left and right use other clearance
                double finalHeight = trayHeight + topClearance + otherClearance; // Top uses top clearance, bottom uses other
                
                // ⚠️ DIAGNOSTIC: Log final calculated dimensions
                double finalWidthMm = UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters);
                double finalHeightMm = UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters);
                DebugLogger.Info($"[CableTrayStrategy] Top={topClearance:F4}ft, Other={otherClearance:F4}ft, Offset={offsetAmount:F4}ft upward");
                DebugLogger.Info($"[CableTrayStrategy] FINAL SIZE: Width={finalWidthMm:F1}mm ({finalWidth:F6}ft) x Height={finalHeightMm:F1}mm ({finalHeight:F6}ft), Offset: {offsetVector}");
                
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
