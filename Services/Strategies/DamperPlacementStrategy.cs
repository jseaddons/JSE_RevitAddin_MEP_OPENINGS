using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Damper (Duct Accessory) placement strategy
    /// Handles fire dampers and other duct accessories with NO clearance addition (uses actual size)
    /// </summary>
    public class DamperPlacementStrategy : ISleevePlacementStrategy
    {
        private readonly Document _doc;
        private readonly IDamperTypeDetector _damperTypeDetector;
        private readonly IInsulationAwareSizingService _sizingService;
        
        public DamperPlacementStrategy(Document doc)
            : this(doc, new DamperTypeDetector(), new InsulationAwareSizingService())
        {
        }
        
        /// <summary>
        /// Constructor with dependency injection for testing (SOLID principles)
        /// </summary>
        public DamperPlacementStrategy(Document doc, IDamperTypeDetector typeDetector, IInsulationAwareSizingService sizingService)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _damperTypeDetector = typeDetector ?? throw new ArgumentNullException(nameof(typeDetector));
            _sizingService = sizingService ?? throw new ArgumentNullException(nameof(sizingService));
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
            
            // ✅ OOP METHOD: Use IDamperTypeDetector to detect damper type
            string familyTypeName = damper.Symbol?.Name ?? "";
            size.DamperType = _damperTypeDetector.DetectDamperType(familyTypeName);
            
            DebugLogger.Info($"[DamperStrategy] Detected damper type: '{size.DamperType}' from family: '{familyTypeName}'");
            
            return size;
        }
        
        public double GetClearance(MepElementSize mepSize, OpeningConditions conditions)
        {
            // Use the damper type stored in MepElementSize during GetMepElementSize
            string damperType = mepSize.DamperType ?? "";
            DebugLogger.Info($"[DamperStrategy] Damper type from MepElementSize: '{damperType}'");
            
            // ✅ OOP METHOD: Use IDamperTypeDetector to determine clearance requirements
            double clearanceInMm;
            if (_damperTypeDetector.RequiresMepSideClearance(damperType))
            {
                clearanceInMm = conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0; // Non-standard dampers use MEP side clearance
                DebugLogger.Info($"[DamperStrategy] {damperType} damper - using MEP side clearance: {clearanceInMm}mm");
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
                
                // ✅ DIRECT LOGGING: Always log what values we're reading from ClashZone
                DebugLogger.Info($"[DamperStrategy] Zone {clashZone.Id}: Reading MepElementWidth={damperWidth:F6}ft ({damperWidth * 304.8:F1}mm), MepElementHeight={damperHeight:F6}ft ({damperHeight * 304.8:F1}mm)");
                SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-READ] Zone {clashZone.Id}: MepElementWidth={damperWidth:F6}ft ({damperWidth * 304.8:F1}mm), MepElementHeight={damperHeight:F6}ft ({damperHeight * 304.8:F1}mm)\n");
                
                // Get clearance from OpeningConditions (loaded from UI via CONDITIONS XML)
                double otherClearanceMm = conditions.ClearanceSettings.DuctAccessoryOtherNormal;
                double mepClearanceMm = conditions.ClearanceSettings.DuctAccessoryMepNormal;
                
                double otherClearance = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                double mepClearance = UnitUtils.ConvertToInternalUnits(mepClearanceMm, UnitTypeId.Millimeters);
                
                // ✅ OOP METHOD: Get insulation contribution using sizing service (SOLID principles)
                double insulationContribution = 0.0;
                if (clashZone.IsInsulated && clashZone.InsulationThickness > 0)
                {
                    insulationContribution = 2 * clashZone.InsulationThickness; // Both sides
                    DebugLogger.Info($"[DamperStrategy] DAMPER INSULATED: Insulation contribution={RevitUnitConversionService.Instance.FromInternalMillimeters(insulationContribution):F1}mm total (on both sides)");
                }
                
                // ✅ OOP METHOD: Log connector detection results (connector-based logic, not damper type)
                DebugLogger.Info($"[DamperStrategy] 🔍 CONNECTOR DETECTION: HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}'");
                SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-CONNECTOR-DETECTION] Zone {clashZone.Id}: HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}'\n");
                
                DebugLogger.Info($"[DamperStrategy] Clearances from conditions: MEP={mepClearanceMm}mm, Other={otherClearanceMm}mm");
                SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-CLEARANCE] Zone {clashZone.Id}: MEP={mepClearanceMm}mm, Other={otherClearanceMm}mm\n");
                
                // ✅ OOP METHOD: Check if connector was detected (regardless of damper type)
                // If connector exists, use MEP+Other on width (MEP side 100mm + Other side 50mm = 150mm total)
                // If no connector, use Other clearance on all sides (50mm + 50mm = 100mm total)
                if (clashZone.HasMepConnector && !string.IsNullOrEmpty(clashZone.DamperConnectorSide))
                {
                    // MSFD Damper: Asymmetric clearance
                    double mepSideClearance = mepClearance;
                    double otherSideClearance = otherClearance;
                    
                    // Determine which dimension gets the asymmetric clearance
                    double left = otherSideClearance, right = otherSideClearance;
                    double top = otherSideClearance, bottom = otherSideClearance;
                    
                    // ✅ OFFSET ALONG WALL AXIS: Calculate offset to achieve correct clearance distribution
                    // Methodology: 
                    // 1. Sleeve is sized: Base (500mm) + MEP clearance (100mm) + Other clearance (50mm) = 650mm total
                    // 2. When centered on damper: clearance is 75mm on each side
                    // 3. To achieve 100mm on connector side and 50mm on other side, move by difference/2
                    //    Offset = (mepClearance - otherClearance) / 2 = (100 - 50) / 2 = 25mm toward connector
                    // 4. After move: Connector side = 75 + 25 = 100mm ✓, Other side = 75 - 25 = 50mm ✓
                    double offsetAmount = (mepSideClearance - otherSideClearance) / 2.0;
                    XYZ offsetVector = XYZ.Zero;
                    
                    // Get wall orientation to determine offset direction
                    string hostOrientation = clashZone.HostOrientation ?? string.Empty;
                    bool isXWall = string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase);
                    bool isYWall = string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase);
                    
                    // Only apply offset for walls (not floors/framing)
                    bool isWallHost = string.Equals(clashZone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(clashZone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                    
                    if (isWallHost && offsetAmount > 0.0001) // Only if there's a significant offset
                    {
                        // Determine offset direction based on connector side
                        // For Left/Right connectors: offset along wall axis
                        // For Top/Bottom connectors: offset along wall axis
                        // X-wall: offset along X-axis | Y-wall: offset along Y-axis
                        bool offsetPositive = clashZone.DamperConnectorSide == "Right" || clashZone.DamperConnectorSide == "Top";
                        
                        if (isXWall)
                        {
                            // X-wall: offset along X-axis (wall runs along X-axis)
                            offsetVector = offsetPositive 
                                ? new XYZ(offsetAmount, 0, 0) 
                                : new XYZ(-offsetAmount, 0, 0);
                        }
                        else if (isYWall)
                        {
                            // Y-wall: offset along Y-axis (wall runs along Y-axis)
                            offsetVector = offsetPositive 
                                ? new XYZ(0, offsetAmount, 0) 
                                : new XYZ(0, -offsetAmount, 0);
                        }
                    }
                    
                    switch (clashZone.DamperConnectorSide)
                    {
                        case "Left":
                            left = mepSideClearance;
                            break;
                        case "Right":
                            right = mepSideClearance;
                            break;
                        case "Top":
                            top = mepSideClearance;
                            break;
                        case "Bottom":
                            bottom = mepSideClearance;
                            break;
                    }
                    
                    // ✅ OOP METHOD: Formula: Base + insulation contribution + asymmetric clearance on each side
                    double finalWidth = damperWidth + insulationContribution + left + right;
                    double finalHeight = damperHeight + insulationContribution + top + bottom;
                    
                    // ✅ DETAILED LOGGING: Log offset calculation details for debugging
                    string offsetInfo = offsetVector.GetLength() > 0.0001 
                        ? $"Offset={offsetAmount*304.8:F1}mm along wall axis ({hostOrientation}-wall)" 
                        : "NO OFFSET";
                    DebugLogger.Info($"[DamperStrategy] MSFD - Connector={clashZone.DamperConnectorSide}, MEP={mepSideClearance:F4}ft ({mepSideClearance*304.8:F1}mm), Other={otherSideClearance:F4}ft ({otherSideClearance*304.8:F1}mm)");
                    DebugLogger.Info($"[DamperStrategy] MSFD - Offset Calculation: (MEP={mepSideClearance*304.8:F1}mm - Other={otherSideClearance*304.8:F1}mm) / 2 = {offsetAmount*304.8:F1}mm toward connector side");
                    DebugLogger.Info($"[DamperStrategy] MSFD - Wall Info: HostOrientation='{hostOrientation}', IsXWall={isXWall}, IsYWall={isYWall}, IsWallHost={isWallHost}");
                    DebugLogger.Info($"[DamperStrategy] MSFD - Offset Vector: {offsetVector} (Length={offsetVector.GetLength()*304.8:F1}mm)");
                    DebugLogger.Info($"[DamperStrategy] MSFD - Calculation: Base({damperWidth:F6}ft={damperWidth*304.8:F1}mm) + Clearance({(left+right):F6}ft={(left+right)*304.8:F1}mm) = Final({finalWidth:F6}ft={finalWidth*304.8:F1}mm)");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-MSFD-FINAL] Zone {clashZone.Id}: Base({damperWidth*304.8:F1}mm) + Clearance({(left+right)*304.8:F1}mm) = Final({finalWidth*304.8:F1}mm x {finalHeight*304.8:F1}mm), {offsetInfo}\n");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-OFFSET-DETAIL] Zone {clashZone.Id}: OffsetAmount={offsetAmount*304.8:F1}mm, HostOrientation='{hostOrientation}', OffsetVector=({offsetVector.X*304.8:F1}, {offsetVector.Y*304.8:F1}, {offsetVector.Z*304.8:F1})mm\n");
                    
                    return (offsetVector, finalWidth, finalHeight);
                }
                else
                {
                    // Damper has NO connector - use symmetric clearance on all sides
                    // ✅ OOP METHOD: No connector = symmetric Other clearance on all 4 sides
                    // ✅ OOP METHOD: Use sizing service for symmetric clearance case (SOLID principles)
                    (double finalW, double finalH, _) = _sizingService.CalculateFinalDimensionsFromClashZone(
                        damperWidth, damperHeight, 0, clashZone, otherClearance);
                    double finalWidth = finalW;
                    double finalHeight = finalH;
                    
                    // ✅ OOP METHOD: Show actual calculation from sizing service (includes insulation if present)
                    double insulationThicknessMm = clashZone.IsInsulated && clashZone.InsulationThickness > 0 
                        ? RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.InsulationThickness) 
                        : 0.0;
                    double clearanceMm = RevitUnitConversionService.Instance.FromInternalMillimeters(otherClearance);
                    
                    if (clashZone.IsInsulated && insulationThicknessMm > 0)
                    {
                        DebugLogger.Info($"[DamperStrategy] Standard (OOP): Base({damperWidth*304.8:F1}mm) + Insulation({insulationThicknessMm*2:F1}mm) + Clearance({clearanceMm*2:F1}mm) = Final({finalWidth*304.8:F1}mm)");
                        SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-STANDARD-FINAL] Zone {clashZone.Id}: Base({damperWidth*304.8:F1}mm) + Insulation({insulationThicknessMm*2:F1}mm) + Clearance({clearanceMm*2:F1}mm) = Final({finalWidth*304.8:F1}mm x {finalHeight*304.8:F1}mm)\n");
                    }
                    else
                    {
                        DebugLogger.Info($"[DamperStrategy] Standard (OOP): Base({damperWidth*304.8:F1}mm) + Clearance({clearanceMm*2:F1}mm) = Final({finalWidth*304.8:F1}mm)");
                        SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-STANDARD-FINAL] Zone {clashZone.Id}: Base({damperWidth*304.8:F1}mm) + Clearance({clearanceMm*2:F1}mm) = Final({finalWidth*304.8:F1}mm x {finalHeight*304.8:F1}mm)\n");
                    }
                    
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
