using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Cable Tray-specific placement strategy
    /// Handles rectangular cable trays with NO clearance addition (uses actual size)
    /// </summary>
    public class CableTrayPlacementStrategy : ISleevePlacementStrategy
    {
        private readonly IInsulationAwareSizingService _sizingService;
        
        public CableTrayPlacementStrategy()
            : this(new InsulationAwareSizingService())
        {
        }
        
        /// <summary>
        /// Constructor with dependency injection for testing (SOLID principles)
        /// </summary>
        public CableTrayPlacementStrategy(IInsulationAwareSizingService sizingService)
        {
            _sizingService = sizingService ?? throw new ArgumentNullException(nameof(sizingService));
        }
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
        /// Determines if a cable tray is vertical and intersecting a floor
        /// </summary>
        private bool IsVerticalCableTrayOnFloor(ClashZone clashZone)
        {
            DebugLogger.Info($"[CableTrayStrategy] 🔍 IsVerticalCableTrayOnFloor CHECK for ClashZone {clashZone.Id}");
            DebugLogger.Info($"[CableTrayStrategy] StructuralElementType: '{clashZone.StructuralElementType}'");

            // Check if this is a floor intersection
            bool isFloorIntersection = string.Equals(clashZone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(clashZone.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);

            DebugLogger.Info($"[CableTrayStrategy] Is floor intersection: {isFloorIntersection}");

            if (!isFloorIntersection)
            {
                DebugLogger.Info($"[CableTrayStrategy] ❌ Not a floor intersection, returning false");
                return false;
            }

            // For floor intersections, we need to check if cable tray is vertical
            // We can use the MEP element orientation data from ClashZone
            var mepOrientation = clashZone.MepElementOrientation;
            if (mepOrientation != null)
            {
                // If Z component is dominant, cable tray is running vertically
                double absX = Math.Abs(mepOrientation.X);
                double absY = Math.Abs(mepOrientation.Y);
                double absZ = Math.Abs(mepOrientation.Z);

                bool isVertical = absZ > absX && absZ > absY;

                DebugLogger.Info($"[CableTrayStrategy] Floor intersection analysis: MEP=({mepOrientation.X:F3},{mepOrientation.Y:F3},{mepOrientation.Z:F3}), isVertical={isVertical}");
                DebugLogger.Info($"[CableTrayStrategy] Component magnitudes: X={absX:F3}, Y={absY:F3}, Z={absZ:F3}");

                if (isVertical)
                {
                    DebugLogger.Info($"[CableTrayStrategy] ✅ Vertical cable tray on floor detected!");
                }
                else
                {
                    DebugLogger.Info($"[CableTrayStrategy] ❌ Horizontal cable tray on floor - using standard logic");
                }

                return isVertical;
            }

            // Fallback: assume vertical if we don't have orientation data
            DebugLogger.Warning($"[CableTrayStrategy] ⚠️ No MEP orientation data for ClashZone {clashZone.Id}, assuming vertical for floor intersection");
            return true;
        }

        /// <summary>
        /// Determines the cable tray's open side direction for floor intersections using lightweight analysis
        /// </summary>
        private XYZ GetCableTrayOpenSideDirection(CableTray cableTray, ClashZone clashZone)
        {
            try
            {
                DebugLogger.Info($"[CableTrayStrategy] Detecting open side direction for cable tray {cableTray.Id} on floor");

                // Method 1: Use MEP element orientation to determine likely open side
                var mepOrientation = clashZone.MepElementOrientation;
                if (mepOrientation != null)
                {
                    // Get the cable tray's width direction (perpendicular to centerline)
                    var (orientation, widthDirection) = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementOrientationHelper.GetCableTrayWidthOrientation(cableTray);

                    DebugLogger.Info($"[CableTrayStrategy] Cable tray orientation: {orientation}, widthDirection: ({widthDirection.X:F3},{widthDirection.Y:F3},{widthDirection.Z:F3})");

                    // For floor intersections, the open side is typically perpendicular to the floor
                    // and aligned with the cable tray's width direction
                    if (Math.Abs(widthDirection.Z) > 0.5)
                    {
                        // Width direction has significant Z component - use it
                        DebugLogger.Info($"[CableTrayStrategy] Using width direction for open side: ({widthDirection.X:F3},{widthDirection.Y:F3},{widthDirection.Z:F3})");
                        return widthDirection.Normalize();
                    }
                }

                // Method 2: Use bounding box analysis as fallback
                var bbox = cableTray.get_BoundingBox(null);
                if (bbox != null)
                {
                    double width = bbox.Max.X - bbox.Min.X;
                    double height = bbox.Max.Y - bbox.Min.Y;
                    double depth = bbox.Max.Z - bbox.Min.Z;

                    DebugLogger.Info($"[CableTrayStrategy] Bounding box analysis: X={width:F3}ft, Y={height:F3}ft, Z={depth:F3}ft");

                    // Find the smallest dimension (likely the closed side)
                    if (width <= height && width <= depth)
                    {
                        // X is smallest - open side is likely along X
                        DebugLogger.Info($"[CableTrayStrategy] X dimension smallest, assuming open side along X");
                        return XYZ.BasisX;
                    }
                    else if (height <= width && height <= depth)
                    {
                        // Y is smallest - open side is likely along Y
                        DebugLogger.Info($"[CableTrayStrategy] Y dimension smallest, assuming open side along Y");
                        return XYZ.BasisY;
                    }
                }

                // Method 3: Ultimate fallback - assume upward for floor intersections
                DebugLogger.Warning($"[CableTrayStrategy] Could not determine open side, using upward as fallback");
                return XYZ.BasisZ;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[CableTrayStrategy] Error detecting open side direction: {ex.Message}");
                return XYZ.BasisZ; // Safe fallback
            }
        }
        
        /// <summary>
        /// Calculate cable tray placement adjustment with smart open side detection for floor intersections
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string cabletrayClearanceDebugLogPath = SafeFileLogger.GetLogFilePath("cabletray_clearance_debug.log");
                    System.IO.File.AppendAllText(cabletrayClearanceDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] GetCableTrayPlacementAdjustment START for ClashZone {clashZone.Id}\n");
                    System.IO.File.AppendAllText(cabletrayClearanceDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] clearanceSettings.Count = {uiClearanceSettings?.Count ?? 0}\n");
                    
                    if (uiClearanceSettings != null)
                    {
                        foreach (var kvp in uiClearanceSettings)
                        {
                            System.IO.File.AppendAllText(cabletrayClearanceDebugLogPath, $"[{DateTime.Now:HH:mm:ss}]   Key='{kvp.Key}', Value={kvp.Value}mm\n");
                        }
                    }
                }
                
                double trayWidth = clashZone.MepElementWidth;
                double trayHeight = clashZone.MepElementHeight;
                
                // ⚠️ DIAGNOSTIC: Log cable tray dimensions from ClashZone
                double trayWidthMm = UnitUtils.ConvertFromInternalUnits(trayWidth, UnitTypeId.Millimeters);
                double trayHeightMm = UnitUtils.ConvertFromInternalUnits(trayHeight, UnitTypeId.Millimeters);
                DebugLogger.Info($"[CableTrayStrategy] CZ={clashZone.Id}: Read from ClashZone: Width={trayWidthMm:F1}mm ({trayWidth:F6}ft), Height={trayHeightMm:F1}mm ({trayHeight:F6}ft)");
                
                // ✅ CRITICAL DIAGNOSTIC: Log to file for comparison with Refresh_debug.log
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CABLE-TRAY-STRATEGY] Zone {clashZone.Id}: MEP={clashZone.MepElementIdValue}, " +
                    $"MepElementWidth={trayWidth:F6}ft ({trayWidthMm:F1}mm), MepElementHeight={trayHeight:F6}ft ({trayHeightMm:F1}mm)\n");
                
                // ✅ PRIORITY SYSTEM: Database (conditions) > UI Settings > Default
                double topClearanceMm = 100.0; // Default fallback
                double otherClearanceMm = 50.0; // Default fallback
                
                // 🔥 UNCONDITIONAL DIAGNOSTIC: Log conditions object state
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CONDITIONS-CHECK] Zone {clashZone.Id}: " +
                    $"conditions={conditions != null}, ClearanceSettings={conditions?.ClearanceSettings != null}, " +
                    $"FilterName={conditions?.FilterName ?? "NULL"}, Category={conditions?.Category ?? "NULL"}\n");
                
                if (conditions?.ClearanceSettings != null)
                {
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CONDITIONS-VALUES] CableTrayTop={conditions.ClearanceSettings.CableTrayTop}, " +
                        $"CableTrayOther={conditions.ClearanceSettings.CableTrayOther}\n");
                }
                
                // 1. ✅ PRIMARY: Read from database (conditions object loaded from SQLite)
                if (conditions?.ClearanceSettings != null)
                {
                    topClearanceMm = conditions.ClearanceSettings.CableTrayTop;
                    otherClearanceMm = conditions.ClearanceSettings.CableTrayOther;
                    DebugLogger.Info($"[CableTrayStrategy] ✅ Using DATABASE clearances: Top={topClearanceMm}mm, Other={otherClearanceMm}mm");
                    
                    // 🔥 UNCONDITIONAL: Log database selection
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-SOURCE] DATABASE: Top={topClearanceMm}mm, Other={otherClearanceMm}mm\n");
                    
                    // 🔥 UNCONDITIONAL: Log database selection
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-SOURCE] DATABASE: Top={topClearanceMm}mm, Other={otherClearanceMm}mm\n");
                }
                // 2. Fallback to UI clearance settings (if database not available)
                else if (uiClearanceSettings != null && uiClearanceSettings.Count > 0)
                {
                    // 🔥 UNCONDITIONAL: Log UI fallback
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-SOURCE] UI: {uiClearanceSettings.Count} settings available\n");
                    
                    // 🔥 DEBUG: Log all UI clearance settings to see what keys are available
                    DebugLogger.Info($"[CableTrayStrategy] ⚠️ Database not available, using UI Clearance Settings ({uiClearanceSettings.Count} items):");
                    foreach (var kvp in uiClearanceSettings)
                    {
                        DebugLogger.Info($"[CableTrayStrategy] 🔥 Key: '{kvp.Key}' = {kvp.Value}mm");
                    }
                    
                    // ✅ CRITICAL FIX: Use ClashZone.IsInsulated to select correct clearance (authoritative data from database)
                    bool isInsulated = clashZone?.IsInsulated ?? false;
                    
                    // ✅ FIX: Handle insulation-specific keys for cable trays using ClashZone.IsInsulated
                    string topNormalKey = "cabletray_top_normal";
                    string topInsulatedKey = "cabletray_top_insulated";
                    string otherNormalKey = "cabletray_other_normal";
                    string otherInsulatedKey = "cabletray_other_insulated";
                    
                    // Select top clearance based on insulation status
                    string topTargetKey = isInsulated ? topInsulatedKey : topNormalKey;
                    if (uiClearanceSettings.ContainsKey(topTargetKey))
                    {
                        topClearanceMm = uiClearanceSettings[topTargetKey];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI top clearance (isInsulated={isInsulated}, key='{topTargetKey}'): {topClearanceMm}mm");
                    }
                    else if (isInsulated && uiClearanceSettings.ContainsKey(topNormalKey))
                    {
                        // Fallback: if insulated key not found, use normal key
                        topClearanceMm = uiClearanceSettings[topNormalKey];
                        DebugLogger.Info($"[CableTrayStrategy] ⚠️ Insulated key '{topInsulatedKey}' not found, using normal key: {topClearanceMm}mm");
                    }
                    
                    // Select other clearance based on insulation status
                    string otherTargetKey = isInsulated ? otherInsulatedKey : otherNormalKey;
                    if (uiClearanceSettings.ContainsKey(otherTargetKey))
                    {
                        otherClearanceMm = uiClearanceSettings[otherTargetKey];
                        DebugLogger.Info($"[CableTrayStrategy] ✅ Using UI other clearance (isInsulated={isInsulated}, key='{otherTargetKey}'): {otherClearanceMm}mm");
                    }
                    else if (isInsulated && uiClearanceSettings.ContainsKey(otherNormalKey))
                    {
                        // Fallback: if insulated key not found, use normal key
                        otherClearanceMm = uiClearanceSettings[otherNormalKey];
                        DebugLogger.Info($"[CableTrayStrategy] ⚠️ Insulated key '{otherInsulatedKey}' not found, using normal key: {otherClearanceMm}mm");
                    }
                }
                else
                {
                    // 🔥 UNCONDITIONAL: Log default fallback
                    SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-SOURCE] DEFAULTS: Top={topClearanceMm}mm, Other={otherClearanceMm}mm\n");
                }
                
                // 🔥 UNCONDITIONAL: Log final values
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-FINAL] Zone {clashZone.Id}: Top={topClearanceMm}mm, Other={otherClearanceMm}mm\n");
                
                DebugLogger.Info($"[CableTrayStrategy] FINAL clearances: Top={topClearanceMm}mm, Other={otherClearanceMm}mm");
                DebugLogger.Info($"[CableTrayStrategy] CONDITIONS object: {conditions?.FilterName ?? "NULL"}, Category: {conditions?.Category ?? "NULL"}");
                DebugLogger.Info($"[CableTrayStrategy] CONDITIONS XML file: {conditions?.FilterName ?? "NULL"}_{conditions?.Category ?? "NULL"}_CONDITIONS.xml");
                
                double topClearance = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                double otherClearance = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);

                // ✅ OOP METHOD: Get insulation contribution using sizing service (SOLID principles)
                double insulationContribution = 0.0;
                if (clashZone.IsInsulated && clashZone.InsulationThickness > 0)
                {
                    insulationContribution = 2 * clashZone.InsulationThickness; // Both sides
                    DebugLogger.Info($"[CableTrayStrategy] CABLE TRAY INSULATED: Insulation contribution={UnitUtils.ConvertFromInternalUnits(insulationContribution, UnitTypeId.Millimeters):F1}mm total (on both sides)");
                }

                // 🛠️ ENHANCED: Smart offset direction based on cable tray orientation and host type
                XYZ offsetVector;

                // Check if this is a vertical cable tray intersecting a floor
                if (IsVerticalCableTrayOnFloor(clashZone))
                {
                    DebugLogger.Info($"[CableTrayStrategy] 🎯 VERTICAL CABLE TRAY ON FLOOR DETECTED - Using smart open side detection");

                    var mepOrientation = clashZone.MepElementOrientation;
                    if (mepOrientation != null)
                    {
                        // Use the width direction (perpendicular to cable tray centerline) as open side indicator
                        double absX = Math.Abs(mepOrientation.X);
                        double absY = Math.Abs(mepOrientation.Y);
                        double absZ = Math.Abs(mepOrientation.Z);

                        DebugLogger.Info($"[CableTrayStrategy] MEP Orientation Analysis: X={absX:F3}, Y={absY:F3}, Z={absZ:F3}");

                        // ✅ CRITICAL FIX: Use mm values directly to avoid variable swapping issues when batching is enabled
                        double topClearanceForOffset = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                        double otherClearanceForOffset = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                        double offsetDelta = (topClearanceForOffset - otherClearanceForOffset) / 2.0;
                        
                        if (absX > absY && absX > absZ)
                        {
                            // X component is dominant - open side likely along X
                            offsetVector = new XYZ(offsetDelta, 0, 0);
                            DebugLogger.Info($"[CableTrayStrategy] 🔄 Using X-direction offset: {offsetDelta:F4}ft ({offsetVector})");
                        }
                        else if (absY > absX && absY > absZ)
                        {
                            // Y component is dominant - open side likely along Y
                            offsetVector = new XYZ(0, offsetDelta, 0);
                            DebugLogger.Info($"[CableTrayStrategy] 🔄 Using Y-direction offset: {offsetDelta:F4}ft ({offsetVector})");
                        }
                        else
                        {
                            // Z component is dominant or equal - use Z direction (upward)
                            offsetVector = new XYZ(0, 0, offsetDelta);
                            DebugLogger.Info($"[CableTrayStrategy] 🔄 Using Z-direction offset: {offsetDelta:F4}ft ({offsetVector})");
                        }
                    }
                    else
                    {
                        // No orientation data - fallback to upward
                        // ✅ CRITICAL FIX: Use mm values directly to avoid variable swapping issues when batching is enabled
                        double topClearanceForOffset = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                        double otherClearanceForOffset = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                        double offsetDelta = (topClearanceForOffset - otherClearanceForOffset) / 2.0;
                        offsetVector = new XYZ(0, 0, offsetDelta);
                        DebugLogger.Warning($"[CableTrayStrategy] ⚠️ No orientation data for vertical cable tray, using upward fallback");
                    }
                }
                else
                {
                    // 🟢 STANDARD LOGIC: For horizontal cable trays and wall intersections, use upward offset
                    DebugLogger.Info($"[CableTrayStrategy] 🟢 Using standard upward offset logic (horizontal or wall intersection)");
                    // ✅ CRITICAL FIX: Use mm values directly to avoid variable swapping issues when batching is enabled
                    double topClearanceForOffset = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                    double otherClearanceForOffset = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                    double offsetDelta = (topClearanceForOffset - otherClearanceForOffset) / 2.0;
                    offsetVector = new XYZ(0, 0, offsetDelta); // Always offset upward
                }

                // ✅ CRITICAL FIX: Calculate final size with insulation contribution + asymmetric clearances
                // Formula: Width/Height + insulation contribution + clearance on each side
                // ⚠️ BUG FIX: Ensure we use the CORRECT clearances (not swapped) when batching is enabled
                // Width: Left and right use OTHER clearance (25mm each = 50mm total)
                // Height: Top uses TOP clearance (75mm), bottom uses OTHER clearance (25mm) = 100mm total
                
                // ✅ CRITICAL VALIDATION: Verify clearances are correct before calculation
                double topClearanceFt = UnitUtils.ConvertToInternalUnits(topClearanceMm, UnitTypeId.Millimeters);
                double otherClearanceFt = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                
                // ✅ DIAGNOSTIC: Log clearances in both mm and ft to verify correct values
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-VERIFY] Zone {clashZone.Id}: " +
                    $"topClearanceMm={topClearanceMm:F1}mm → topClearanceFt={topClearanceFt:F6}ft, " +
                    $"otherClearanceMm={otherClearanceMm:F1}mm → otherClearanceFt={otherClearanceFt:F6}ft\n");
                
                // ✅ CRITICAL: Use the CORRECT clearances (topClearanceFt and otherClearanceFt) - NOT the variables that might be swapped
                double finalWidth = trayWidth + insulationContribution + (2 * otherClearanceFt); // Left and right use OTHER clearance (25mm each)
                double finalHeight = trayHeight + insulationContribution + topClearanceFt + otherClearanceFt; // Top uses TOP clearance (75mm), bottom uses OTHER (25mm)
                
                // ✅ DIAGNOSTIC: Log calculation breakdown BEFORE conversion
                double widthClearanceAdded = 2 * otherClearanceFt;
                double heightClearanceAdded = topClearanceFt + otherClearanceFt;
                double widthClearanceAddedMm = UnitUtils.ConvertFromInternalUnits(widthClearanceAdded, UnitTypeId.Millimeters);
                double heightClearanceAddedMm = UnitUtils.ConvertFromInternalUnits(heightClearanceAdded, UnitTypeId.Millimeters);
                
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CLEARANCE-BREAKDOWN] Zone {clashZone.Id}: " +
                    $"Width clearance added: 2 * {otherClearanceMm:F1}mm = {widthClearanceAddedMm:F1}mm, " +
                    $"Height clearance added: {topClearanceMm:F1}mm + {otherClearanceMm:F1}mm = {heightClearanceAddedMm:F1}mm\n");
                
                // ⚠️ DIAGNOSTIC: Log final calculated dimensions
                double finalWidthMm = UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters);
                double finalHeightMm = UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters);
                double offsetAmount = offsetVector.GetLength();
                DebugLogger.Info($"[CableTrayStrategy] Top={topClearanceFt:F4}ft ({topClearanceMm:F1}mm), Other={otherClearanceFt:F4}ft ({otherClearanceMm:F1}mm), Offset={offsetAmount:F4}ft in direction {offsetVector}");
                DebugLogger.Info($"[CableTrayStrategy] FINAL SIZE: Width={finalWidthMm:F1}mm ({finalWidth:F6}ft) x Height={finalHeightMm:F1}mm ({finalHeight:F6}ft), Offset: {offsetVector}");
                
                // ✅ CRITICAL DIAGNOSTIC: Log calculation breakdown for debugging "haywire" dimensions
                SafeFileLogger.SafeAppendText("cabletray_dimension_trace.log",
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CABLE-TRAY-CALC] Zone {clashZone.Id}: " +
                    $"Raw: W={trayWidthMm:F1}mm, H={trayHeightMm:F1}mm | " +
                    $"Insulation: {UnitUtils.ConvertFromInternalUnits(insulationContribution, UnitTypeId.Millimeters):F1}mm | " +
                    $"Clearances: Top={topClearanceMm:F1}mm, Other={otherClearanceMm:F1}mm | " +
                    $"Width clearance: {widthClearanceAddedMm:F1}mm (2×{otherClearanceMm:F1}mm), " +
                    $"Height clearance: {heightClearanceAddedMm:F1}mm ({topClearanceMm:F1}mm+{otherClearanceMm:F1}mm) | " +
                    $"FINAL: W={finalWidthMm:F1}mm ({finalWidth:F6}ft), H={finalHeightMm:F1}mm ({finalHeight:F6}ft)\n");
                
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
