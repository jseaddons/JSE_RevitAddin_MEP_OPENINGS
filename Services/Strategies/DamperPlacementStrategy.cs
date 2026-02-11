using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    /// <summary>
    /// Damper (Duct Accessory) placement strategy
    /// Handles fire dampers and other duct accessories with NO clearance addition (uses actual size)
    /// </summary>
    public class DamperPlacementStrategy : ISleevePlacementStrategy
    {
        /// <summary>
        /// PERFORMANCE: Enable/disable detailed debug logging (default: false for production)
        /// Set to true only when debugging damper placement issues
        /// </summary>
        public static bool EnableDebugLogging { get; set; } = true;
        
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
            if (damper == null || damper.Category?.Id.GetIntegerValue() != (int)BuiltInCategory.OST_DuctAccessory)
            {
                DebugLogger.Warning($"[DamperStrategy] Element {mepElement?.Id} is not a Duct Accessory");
                return new MepElementSize();
            }
            
            var size = new MepElementSize { Shape = "Rectangular" }; // Most dampers are rectangular
            
            // Get size from damper-specific parameters (not generic Width/Height)
            // Fire dampers use "Damper Width" and "Damper Height" parameters
            // User Request (2025-12-31): Add "Dimensions Width" / "Dimension Width" / "Dimensions_Width" support
            var widthParam = damper.LookupParameter("Damper Width") ?? 
                            damper.LookupParameter("Width") ?? 
                            damper.LookupParameter("width") ??
                            damper.LookupParameter("Dimensions_Width") ??
                            damper.LookupParameter("Dimensions Width") ??
                            damper.LookupParameter("Dimension Width") ??
                            damper.LookupParameter("dimensions width") ??
                            damper.LookupParameter("dimension width");
                            
            var heightParam = damper.LookupParameter("Damper Height") ?? 
                             damper.LookupParameter("Height") ?? 
                             damper.LookupParameter("height") ??
                             damper.LookupParameter("Dimensions_Height") ??
                             damper.LookupParameter("Dimensions Height") ??
                             damper.LookupParameter("Dimension Height") ??
                             damper.LookupParameter("dimensions height") ??
                             damper.LookupParameter("dimension height");
            
            size.Width = widthParam?.AsDouble() ?? 0.0;
            size.Height = heightParam?.AsDouble() ?? 0.0;
            
            // ✅ R2024 FIX: If parameters fail (0.0), use Geometry Bounding Box
            if (size.Width <= 0.001 || size.Height <= 0.001)
            {
                DebugLogger.Warning($"[DamperStrategy] Damper {damper.Id}: Zero dimensions from parameters. Attempting BoundingBox calculation.");
                var bbox = damper.get_BoundingBox(null);
                if (bbox != null)
                {
                    // Height is Z-delta
                    size.Height = bbox.Max.Z - bbox.Min.Z;
                    
                    // Width depends on Wall Orientation
                    double deltaX = bbox.Max.X - bbox.Min.X;
                    double deltaY = bbox.Max.Y - bbox.Min.Y;
                    
                    // Try to get Host Wall to determine Orientation
                    // Note: This is a strategy-level detection, independent of ClashZoneService
                    Element host = damper.Host; 
                    // If host is null (e.g. invalid) check if it's face hosted or try to find intersecting wall? 
                    // For now, assume Max lateral dimension is Width (Safest default for square/rectangular dampers)
                    // If we want to be smarter:
                    string orientation = "Unknown";
                    if (host is Wall wall)
                    {
                         // Basic orientation check
                         XYZ normal = wall.Orientation;
                         if (Math.Abs(normal.X) > Math.Abs(normal.Y)) orientation = "X"; // Normal X -> Wall runs Y -> Width is Y? No.
                         // Wall runs PERPENDICULAR to Normal.
                         // If Normal is X, Wall runs along Y. Width is along Y.
                         // If Normal is Y, Wall runs along X. Width is along X.
                         
                         if (Math.Abs(normal.X) > Math.Abs(normal.Y)) 
                         {
                             // Wall is Y-aligned (runs North-South)
                             size.Width = deltaY;
                         }
                         else
                         {
                             // Wall is X-aligned (runs East-West)
                             size.Width = deltaX; 
                         }
                    }
                    else
                    {
                         // Fallback: Max lateral dimension
                         size.Width = Math.Max(deltaX, deltaY);
                    }
                    
                    DebugLogger.Info($"[DamperStrategy] Calculated from BBox: Width={size.Width} ft, Height={size.Height} ft");
                }
            }
            
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
            // CRASH-PROOF: Validate inputs
            if (clashZone == null)
            {
                DebugLogger.Error("[DamperStrategy] ClashZone is null, returning zero offset and dimensions");
                return (XYZ.Zero, 0, 0);
            }
            
                // ✅ COMPREHENSIVE LOGGING: Log conditions object state
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-STRATEGY-ENTRY] Zone {clashZone.Id}, " +
                        $"Conditions={conditions != null}, ClearanceSettings={conditions?.ClearanceSettings != null}, " +
                        $"IsInsulated={clashZone.IsInsulated}\n");
                }
                
                if (conditions == null || conditions.ClearanceSettings == null)
            {
                DebugLogger.Error($"[DamperStrategy] Zone {clashZone.Id}: Conditions or ClearanceSettings is null, returning zero offset and original dimensions");
                return (XYZ.Zero, clashZone.MepElementWidth, clashZone.MepElementHeight);
            }
            
            try
            {
                double damperWidth = clashZone.MepElementWidth;
                double damperHeight = clashZone.MepElementHeight;
                
                // CRASH-PROOF: Validate dimensions
                if (double.IsNaN(damperWidth) || double.IsInfinity(damperWidth) || damperWidth < 0)
                {
                    DebugLogger.Warning($"[DamperStrategy] Zone {clashZone.Id}: Invalid damperWidth={damperWidth}, using 0");
                    damperWidth = 0;
                }
                if (double.IsNaN(damperHeight) || double.IsInfinity(damperHeight) || damperHeight < 0)
                {
                    DebugLogger.Warning($"[DamperStrategy] Zone {clashZone.Id}: Invalid damperHeight={damperHeight}, using 0");
                    damperHeight = 0;
                }
                
                // ✅ PERFORMANCE OPTIMIZATION: Conditional logging (eliminates 6+ file writes per damper)
                if (EnableDebugLogging)
                {
                    DebugLogger.Info($"[DamperStrategy] Zone {clashZone.Id}: Reading MepElementWidth={damperWidth:F6}ft ({damperWidth * 304.8:F1}mm), MepElementHeight={damperHeight:F6}ft ({damperHeight * 304.8:F1}mm)");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-READ] Zone {clashZone.Id}: MepElementWidth={damperWidth:F6}ft ({damperWidth * 304.8:F1}mm), MepElementHeight={damperHeight:F6}ft ({damperHeight * 304.8:F1}mm)\n");
                }
                
                // Get clearance from OpeningConditions (loaded from UI via CONDITIONS XML) - CRASH-PROOF: validate clearances
                // Get clearance from OpeningConditions (loaded from UI via CONDITIONS XML) - CRASH-PROOF: validate clearances
                // ✅ FIX: Force uninsulated for dampers (Duct Accessories usually don't have insulation or it's part of casing)
                // This prevents phantom insulation thickness (e.g. 50mm) from being added to the sleeve width
                bool isInsulated = false; // Forced false to match GetMepElementSize logic
                double otherClearanceMm = conditions.ClearanceSettings.DuctAccessoryOtherNormal;
                double mepClearanceMm = conditions.ClearanceSettings.DuctAccessoryMepNormal;
                
                // CRASH-PROOF: Validate clearance values
                if (double.IsNaN(otherClearanceMm) || double.IsInfinity(otherClearanceMm) || otherClearanceMm < 0)
                {
                    DebugLogger.Warning($"[DamperStrategy] Zone {clashZone.Id}: Invalid otherClearanceMm={otherClearanceMm}, using 50mm default");
                    otherClearanceMm = 50.0;
                }
                if (double.IsNaN(mepClearanceMm) || double.IsInfinity(mepClearanceMm) || mepClearanceMm < 0)
                {
                    DebugLogger.Warning($"[DamperStrategy] Zone {clashZone.Id}: Invalid mepClearanceMm={mepClearanceMm}, using 100mm default");
                    mepClearanceMm = 100.0;
                }
                
                // ✅ COMPREHENSIVE LOGGING: Log clearance values read from database BEFORE conversion
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-STRATEGY-CLEARANCE] Zone {clashZone.Id}, " +
                        $"MEP Clearance={mepClearanceMm}mm, Other Clearance={otherClearanceMm}mm, " +
                        $"IsInsulated={isInsulated} (Forced False), " +
                        $"Raw MEP W={clashZone.MepElementWidth * 304.8:F1}mm, H={clashZone.MepElementHeight * 304.8:F1}mm\n");
                }
                
                double otherClearance = UnitUtils.ConvertToInternalUnits(otherClearanceMm, UnitTypeId.Millimeters);
                double mepClearance = UnitUtils.ConvertToInternalUnits(mepClearanceMm, UnitTypeId.Millimeters);
                
                // ✅ OOP METHOD: Insulation contribution is ZERO for dampers (already forced uninsulated)
                double insulationContribution = 0.0;
                // if (clashZone.IsInsulated && clashZone.InsulationThickness > 0)
                // {
                //      insulationContribution = 2 * clashZone.InsulationThickness; 
                // }
                
                // ✅ PERFORMANCE OPTIMIZATION: Conditional logging
                if (EnableDebugLogging)
                {
                    // ✅ OOP METHOD: Log connector detection results (connector-based logic, not damper type)
                    // ✅ NEW: DamperConnectorSide now returns world coordinate directions: "+X", "-X", "+Y", "-Y", "+Z", "-Z"
                    DebugLogger.Info($"[DamperStrategy] 🔍 CONNECTOR DETECTION: HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}' (World Coordinate Direction)");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                        $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-CONNECTOR-DETECTION] Zone {clashZone.Id}: " +
                        $"HasMepConnector={clashZone.HasMepConnector}, " +
                        $"DamperConnectorSide='{clashZone.DamperConnectorSide}' (World Coordinate Direction), " +
                        $"HostOrientation='{clashZone.HostOrientation ?? "Unknown"}', " +
                        $"StructuralElementType='{clashZone.StructuralElementType ?? "Unknown"}'\n");
                    
                    DebugLogger.Info($"[DamperStrategy] Clearances from conditions: MEP={mepClearanceMm}mm, Other={otherClearanceMm}mm");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-CLEARANCE] Zone {clashZone.Id}: MEP={mepClearanceMm}mm, Other={otherClearanceMm}mm\n");
                }
                
                // ✅ SIMPLIFIED LOGIC: Check if damper has MEP connector
                // If HasMepConnector = true → use asymmetric clearance based on DamperConnectorSide
                // If HasMepConnector = false → use symmetric clearance

                // Always log branching decision for diagnostics
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[DamperStrategy] Zone {clashZone.Id}: HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}'");
                    SafeFileLogger.SafeAppendText("damper_placement_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-BRANCHING] Zone {clashZone.Id}: " +
                        $"HasMepConnector={clashZone.HasMepConnector}, DamperConnectorSide='{clashZone.DamperConnectorSide}'\n");
                }

                string branchReason = clashZone.HasMepConnector && !string.IsNullOrEmpty(clashZone.DamperConnectorSide)
                    ? "MEP CONNECTOR PATH (asymmetric clearance)"
                    : "SYMMETRIC PATH (no MEP connector)";

                DebugLogger.Info($"[DamperStrategy] Zone {clashZone.Id}: BRANCHING DECISION → {branchReason}");
                SafeFileLogger.SafeAppendText("damper_placement_trace.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-BRANCH-DECISION] Zone {clashZone.Id}: " +
                    $"Branch={branchReason}, HasMepConnector={clashZone.HasMepConnector}, " +
                    $"DamperConnectorSide='{clashZone.DamperConnectorSide}'\n");

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
                    
                    // ✅ NEW: Get connector direction (world coordinate: "+X", "-X", "+Y", "-Y", "+Z", "-Z")
                    string connectorDir = clashZone.DamperConnectorSide ?? string.Empty;
                    
                    // Only apply offset for walls (not floors/framing)
                    bool isWallHost = string.Equals(clashZone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(clashZone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                    
                    // ✅ WALL-ONLY FIX: If connector is vertical (+Z/-Z), the damper's family Width/Height are rotated
                    // relative to the wall axes. Swap the base damperWidth/damperHeight before adding clearances.
                    if (isWallHost && (connectorDir == "+Z" || connectorDir == "-Z"))
                    {
                        double originalWidth = damperWidth;
                        double originalHeight = damperHeight;
                        damperWidth = originalHeight;
                        damperHeight = originalWidth;
                        if (EnableDebugLogging)
                        {
                            DebugLogger.Info($"[DamperStrategy] WALL Z-CONNECTOR: Swapped base dimensions for vertical connector: Width {originalWidth:F6}ft→{damperWidth:F6}ft, Height {originalHeight:F6}ft→{damperHeight:F6}ft");
                            SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-Z-SWAP] Zone {clashZone.Id}: WALL host, vertical connector ({connectorDir}) → swapped base dims W {originalWidth*304.8:F1}mm→{damperWidth*304.8:F1}mm, H {originalHeight*304.8:F1}mm→{damperHeight*304.8:F1}mm\n");
                        }
                    }

                    if (isWallHost && offsetAmount > 0.0001) // Only if there's a significant offset
                    {
                        // ✅ CRITICAL FIX: Offset should ONLY be along wall axis (width direction), NOT perpendicular to wall
                        // For X-wall: width is along X-axis, so offset only in X direction
                        // For Y-wall: width is along Y-axis, so offset only in Y direction
                        // Z direction (vertical) is always valid for offset
                        // Perpendicular directions (through wall depth) should NOT cause offset
                        
                        // ✅ DIAGNOSTIC LOG: Trace variables before switch
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                                $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-OFFSET-PRE-SWITCH] Zone {clashZone.Id}: " +
                                $"ConnectorDir='{connectorDir}', HostOrientation='{hostOrientation}', " +
                                $"IsXWall={isXWall}, IsYWall={isYWall}, OffsetAmount={offsetAmount * 304.8:F1}mm\n");
                        }
                        
                        switch (connectorDir)
                        {
                            case "+X":
                                // +X direction: Map to wall width axis based on wall orientation
                                if (isXWall)
                                {
                                    // X-wall: width is along X-axis, so +X = right side → offset in +X direction
                                    offsetVector = new XYZ(offsetAmount, 0, 0);
                                }
                                else if (isYWall)
                                {
                                    // Y-wall: width is along Y-axis, +X connector maps to "right" side → offset in +Y direction
                                    offsetVector = new XYZ(0, offsetAmount, 0);
                                }
                                break;
                            
                            case "-X":
                                // -X direction: Map to wall width axis based on wall orientation
                                if (isXWall)
                                {
                                    // X-wall: width is along X-axis, so -X = left side → offset in -X direction
                                    offsetVector = new XYZ(-offsetAmount, 0, 0);
                                }
                                else if (isYWall)
                                {
                                    // Y-wall: width is along Y-axis, -X connector maps to "left" side → offset in -Y direction
                                    offsetVector = new XYZ(0, -offsetAmount, 0);
                                }
                                break;
                            
                            case "+Y":
                                // +Y direction: Map to wall width axis based on wall orientation
                                if (isYWall)
                                {
                                    // Y-wall: width is along Y-axis, so +Y = right side → offset in +Y direction
                                    offsetVector = new XYZ(0, offsetAmount, 0);
                                }
                                else if (isXWall)
                                {
                                    // X-wall: width is along X-axis, +Y connector maps to "right" side → offset in +X direction
                                    offsetVector = new XYZ(offsetAmount, 0, 0);
                                }
                                break;
                            
                            case "-Y":
                                // -Y direction: Map to wall width axis based on wall orientation
                                if (isYWall)
                                {
                                    // Y-wall: width is along Y-axis, so -Y = left side → offset in -Y direction
                                    offsetVector = new XYZ(0, -offsetAmount, 0);
                                }
                                else if (isXWall)
                                {
                                    // X-wall: width is along X-axis, -Y connector maps to "left" side → offset in -X direction
                                    offsetVector = new XYZ(-offsetAmount, 0, 0);
                                }
                                break;
                            
                            case "+Z":
                                // +Z direction: vertical offset (always valid, not affected by wall orientation)
                                offsetVector = new XYZ(0, 0, offsetAmount);
                                break;
                            
                            case "-Z":
                                // -Z direction: vertical offset (always valid, not affected by wall orientation)
                                offsetVector = new XYZ(0, 0, -offsetAmount);
                                break;
                            
                            default:
                                // Fallback for old format ("Left", "Right", etc.) - keep for backward compatibility during transition
                                bool isLeftRight = connectorDir == "Left" || connectorDir == "Right";
                                bool isTopBottom = connectorDir == "Top" || connectorDir == "Bottom";
                                bool offsetPositive = connectorDir == "Right" || connectorDir == "Top";
                                
                                if (isLeftRight)
                                {
                                    // Left/Right: offset along wall width axis
                                    if (isXWall)
                                        offsetVector = offsetPositive ? new XYZ(offsetAmount, 0, 0) : new XYZ(-offsetAmount, 0, 0);
                                    else if (isYWall)
                                        offsetVector = offsetPositive ? new XYZ(0, offsetAmount, 0) : new XYZ(0, -offsetAmount, 0);
                                }
                                else if (isTopBottom)
                                {
                                    // Top/Bottom: vertical offset (always valid)
                                    offsetVector = offsetPositive ? new XYZ(0, 0, offsetAmount) : new XYZ(0, 0, -offsetAmount);
                                }
                                break;
                        }
                        
                        if (EnableDebugLogging)
                        {
                            DebugLogger.Info($"[DamperStrategy] MSFD - Offset calculation: ConnectorDirection='{connectorDir}', HostOrientation='{hostOrientation}', IsXWall={isXWall}, IsYWall={isYWall}, IsWallHost={isWallHost}, OffsetVector=({offsetVector.X*304.8:F1}, {offsetVector.Y*304.8:F1}, {offsetVector.Z*304.8:F1})mm");
                            SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                                $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-OFFSET-CALC] Zone {clashZone.Id}: " +
                                $"ConnectorDirection='{connectorDir}', " +
                                $"HostOrientation='{hostOrientation}', " +
                                $"OffsetAmount={offsetAmount*304.8:F1}mm, " +
                                $"OffsetVector=({offsetVector.X*304.8:F1}, {offsetVector.Y*304.8:F1}, {offsetVector.Z*304.8:F1})mm\n");
                        }
                    }
                    
                    // ✅ NEW: Map world coordinate directions to clearance sides (wall-aware)
                    // Direct mapping: connector direction in WCS → MEP clearance on that side
                    // For X/Y directions: depends on wall orientation (which axis is width)
                    // For Z direction: always affects height (top/bottom)
                    switch (connectorDir)
                    {
                        case "+X":
                            // +X direction: MEP clearance on +X side
                            if (isXWall)
                            {
                                // X-wall: width is along X-axis, so +X = right side of width
                                right = mepSideClearance;
                            }
                            else if (isYWall)
                            {
                                // Y-wall: +X is through-wall (depth), but width is along Y-axis
                                // For Y-wall, +X direction means right side of width (perpendicular to wall)
                                right = mepSideClearance;
                            }
                            else
                            {
                                // Fallback: assume right side
                                right = mepSideClearance;
                            }
                            break;
                        
                        case "-X":
                            // -X direction: MEP clearance on -X side
                            if (isXWall)
                            {
                                // X-wall: width is along X-axis, so -X = left side of width
                                left = mepSideClearance;
                            }
                            else if (isYWall)
                            {
                                // Y-wall: -X is through-wall (depth), but width is along Y-axis
                                // For Y-wall, -X direction means left side of width (perpendicular to wall)
                                left = mepSideClearance;
                            }
                            else
                            {
                                // Fallback: assume left side
                                left = mepSideClearance;
                            }
                            break;
                        
                        case "+Y":
                            // +Y direction: MEP clearance on +Y side
                            if (isYWall)
                            {
                                // Y-wall: width is along Y-axis, so +Y = right side of width
                                right = mepSideClearance;
                            }
                            else if (isXWall)
                            {
                                // X-wall: +Y is through-wall (depth), but width is along X-axis
                                // For X-wall, +Y direction means right side of width (perpendicular to wall)
                                right = mepSideClearance;
                            }
                            else
                            {
                                // Fallback: assume right side
                                right = mepSideClearance;
                            }
                            break;
                        
                        case "-Y":
                            // -Y direction: MEP clearance on -Y side
                            if (isYWall)
                            {
                                // Y-wall: width is along Y-axis, so -Y = left side of width
                                left = mepSideClearance;
                            }
                            else if (isXWall)
                            {
                                // X-wall: -Y is through-wall (depth), but width is along X-axis
                                // For X-wall, -Y direction means left side of width (perpendicular to wall)
                                left = mepSideClearance;
                            }
                            else
                            {
                                // Fallback: assume left side
                            left = mepSideClearance;
                            }
                            break;
                        
                        case "+Z":
                            // +Z direction: MEP clearance on top (vertical, always affects height)
                            top = mepSideClearance;
                            break;
                        
                        case "-Z":
                            // -Z direction: MEP clearance on bottom (vertical, always affects height)
                            bottom = mepSideClearance;
                            break;
                        
                        // Backward compatibility: old format ("Left", "Right", "Top", "Bottom")
                        case "Right":
                            right = mepSideClearance;
                            break;
                        case "Left":
                            left = mepSideClearance;
                            break;
                        case "Top":
                            top = mepSideClearance;
                            break;
                        case "Bottom":
                            bottom = mepSideClearance;
                            break;
                    }
                    
                    // ✅ CRITICAL: Store individual clearance values in ClashZone for sleeve parameter setting
                    // These values are based on world coordinate directions (+X, -X, +Y, -Y, +Z, -Z)
                    clashZone.ClearanceLeft = left;
                    clashZone.ClearanceRight = right;
                    clashZone.ClearanceTop = top;
                    clashZone.ClearanceBottom = bottom;
                    
                    double finalWidth = damperWidth + insulationContribution + left + right;
                    double finalHeight = damperHeight + insulationContribution + top + bottom;
                    
                    // ✅ COMPREHENSIVE LOGGING: Always log clearance breakdown to trace file (not just when debug enabled)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-CLEARANCE-BREAKDOWN] Zone {clashZone.Id}, " +
                            $"HasConnector=True, ConnectorDir='{connectorDir}', " +
                            $"Left={left * 304.8:F1}mm, Right={right * 304.8:F1}mm, " +
                            $"Top={top * 304.8:F1}mm, Bottom={bottom * 304.8:F1}mm, " +
                            $"WidthTotal={((left + right) * 304.8):F1}mm, HeightTotal={((top + bottom) * 304.8):F1}mm, " +
                            $"Base W={damperWidth * 304.8:F1}mm, H={damperHeight * 304.8:F1}mm, " +
                            $"Final W={finalWidth * 304.8:F1}mm, H={finalHeight * 304.8:F1}mm\n");
                    }
                    
                    // ✅ PERFORMANCE OPTIMIZATION: Conditional detailed logging
                    if (EnableDebugLogging)
                    {
                        // ✅ DETAILED LOGGING: Log clearance assignment and offset calculation details
                        string clearanceSide = "Unknown";
                        if (left == mepSideClearance) clearanceSide = "Left";
                        else if (right == mepSideClearance) clearanceSide = "Right";
                        else if (top == mepSideClearance) clearanceSide = "Top";
                        else if (bottom == mepSideClearance) clearanceSide = "Bottom";
                        
                        string offsetInfo = offsetVector.GetLength() > 0.0001 
                            ? $"Offset={offsetAmount*304.8:F1}mm toward {connectorDir} direction" 
                            : "NO OFFSET";
                        
                        DebugLogger.Info($"[DamperStrategy] MSFD - ConnectorDirection='{connectorDir}' (WCS), HostOrientation='{hostOrientation}', MEPClearanceSide='{clearanceSide}'");
                        DebugLogger.Info($"[DamperStrategy] MSFD - Clearance: MEP={mepSideClearance:F4}ft ({mepSideClearance*304.8:F1}mm) on {clearanceSide}, Other={otherSideClearance:F4}ft ({otherSideClearance*304.8:F1}mm) on opposite side");
                        DebugLogger.Info($"[DamperStrategy] MSFD - Offset Calculation: (MEP={mepSideClearance*304.8:F1}mm - Other={otherSideClearance*304.8:F1}mm) / 2 = {offsetAmount*304.8:F1}mm toward {connectorDir} direction");
                        DebugLogger.Info($"[DamperStrategy] MSFD - Wall Info: HostOrientation='{hostOrientation}', IsXWall={isXWall}, IsYWall={isYWall}, IsWallHost={isWallHost}");
                        DebugLogger.Info($"[DamperStrategy] MSFD - Offset Vector: {offsetVector} (Length={offsetVector.GetLength()*304.8:F1}mm)");
                        DebugLogger.Info($"[DamperStrategy] MSFD - Calculation: Base({damperWidth:F6}ft={damperWidth*304.8:F1}mm) + Clearance({(left+right):F6}ft={(left+right)*304.8:F1}mm width, {(top+bottom):F6}ft={(top+bottom)*304.8:F1}mm height) = Final({finalWidth:F6}ft={finalWidth*304.8:F1}mm x {finalHeight*304.8:F1}mm)");
                        SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-MSFD-FINAL] Zone {clashZone.Id}: " +
                            $"ConnectorDirection='{connectorDir}' (WCS), HostOrientation='{hostOrientation}', MEPClearanceSide='{clearanceSide}', " +
                            $"Base({damperWidth*304.8:F1}mm) + Clearance({(left+right)*304.8:F1}mm width, {(top+bottom)*304.8:F1}mm height) = " +
                            $"Final({finalWidth*304.8:F1}mm x {finalHeight*304.8:F1}mm), {offsetInfo}\n");
                        SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-CLEARANCE-ASSIGNMENT] Zone {clashZone.Id}: " +
                            $"ConnectorDirection='{connectorDir}', HostOrientation='{hostOrientation}', " +
                            $"MEPClearance={mepSideClearance*304.8:F1}mm on {clearanceSide}, " +
                            $"OtherClearance={otherSideClearance*304.8:F1}mm on opposite side\n");
                        SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                            $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-OFFSET-DETAIL] Zone {clashZone.Id}: " +
                            $"OffsetAmount={offsetAmount*304.8:F1}mm, ConnectorDirection='{connectorDir}', " +
                            $"OffsetVector=({offsetVector.X*304.8:F1}, {offsetVector.Y*304.8:F1}, {offsetVector.Z*304.8:F1})mm\n");
                    }
                    
                    return (offsetVector, finalWidth, finalHeight);
                }
                else
                {
                    // ✅ REQUIREMENT: "Standard" dampers → same clearance all sides
                    // Non-standard (MSFD/MSD/MD) without Motorized in family name → symmetric clearance
                    // Non-standard with Motorized but no connector → fallback to symmetric clearance
                    // ✅ OOP METHOD: Standard or no connector = symmetric Other clearance on all 4 sides
                    // ✅ CRITICAL: Store symmetric clearance values in ClashZone for sleeve parameter setting
                    clashZone.ClearanceLeft = otherClearance;
                    clashZone.ClearanceRight = otherClearance;
                    clashZone.ClearanceTop = otherClearance;
                    clashZone.ClearanceBottom = otherClearance;
                    
                    // ✅ OOP METHOD: Use sizing service with rounding (RoundAlwaysUp, RoundingValue) for symmetric clearance case
                    var settings = JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService.Instance.GetCurrentSettings();
                    (double finalW, double finalH, _) = _sizingService.CalculateFinalDimensionsFromClashZoneRounded(
                        damperWidth, damperHeight, 0, clashZone, otherClearance, settings.RoundingValue, settings.RoundAlwaysUp);
                    double finalWidth = finalW;
                    double finalHeight = finalH;
                    
                    // ✅ COMPREHENSIVE LOGGING: Always log clearance breakdown for symmetric clearance case
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-CLEARANCE-BREAKDOWN] Zone {clashZone.Id}, " +
                            $"HasConnector=False, " +
                            $"SymmetricClearance={otherClearance * 304.8:F1}mm (all sides), " +
                            $"WidthTotal={((otherClearance * 2) * 304.8):F1}mm, HeightTotal={((otherClearance * 2) * 304.8):F1}mm, " +
                            $"Base W={damperWidth * 304.8:F1}mm, H={damperHeight * 304.8:F1}mm, " +
                            $"Final W={finalWidth * 304.8:F1}mm, H={finalHeight * 304.8:F1}mm\n");
                    }
                    
                    // ✅ OOP METHOD: Show actual calculation from sizing service (includes insulation if present)
                    if (EnableDebugLogging)
                    {
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