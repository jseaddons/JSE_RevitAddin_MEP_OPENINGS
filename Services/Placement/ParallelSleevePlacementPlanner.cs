using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Parallel implementation of sleeve placement planning. Performs pure computations
    /// (sizes, clearance classification, skip rules, rotation heuristics) prior to sequential Revit operations.
    /// </summary>
    public class ParallelSleevePlacementPlanner : ISleevePlacementPlanner
    {
        private readonly int _minParallelCount;
        private readonly int _maxDegree;
        private readonly OpeningConditions _conditions;
        private readonly Dictionary<string, double> _clearanceSettings;
        private readonly IInsulationAwareSizingService _sizingService;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveRotationService _rotationService;

        public ParallelSleevePlacementPlanner(
            OpeningConditions conditions = null,
            Dictionary<string, double> clearanceSettings = null,
            int minParallelCount = 12,
            int? maxDegree = null,
            IInsulationAwareSizingService sizingService = null)
        {
            _minParallelCount = minParallelCount < 1 ? 1 : minParallelCount;
            _maxDegree = maxDegree ?? Environment.ProcessorCount;
            _conditions = conditions ?? new OpeningConditions();
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            // ✅ OOP METHOD: Initialize sizing service (SOLID principles)
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            _rotationService = new JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveRotationService();
        }

        public SleevePlacementPlanningResult Plan(IEnumerable<ClashZone> clashZones)
        {
            var list = clashZones?.ToList() ?? new List<ClashZone>();
            if (list.Count == 0)
            {
                return new SleevePlacementPlanningResult(Array.Empty<SleevePlacementPlanningDto>(), 0, 0, 0, 0);
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var bag = new ConcurrentBag<SleevePlacementPlanningDto>();
            int skippedCount = 0; int highRisk = 0; int criticalRisk = 0;

            Action<ClashZone> work = zone =>
            {
                var dto = PlanSingle(zone);
                if (dto != null)
                {
                    bag.Add(dto);
                    if (dto.ShouldSkip) System.Threading.Interlocked.Increment(ref skippedCount);
                    else
                    {
                        if (dto.Risk == ClearanceRiskClassification.High) System.Threading.Interlocked.Increment(ref highRisk);
                        if (dto.Risk == ClearanceRiskClassification.Critical) System.Threading.Interlocked.Increment(ref criticalRisk);
                    }
                }
                else
                {
                    System.Threading.Interlocked.Increment(ref skippedCount);
                }
            };

            if (list.Count >= _minParallelCount)
            {
                System.Threading.Tasks.Parallel.ForEach(list, new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = _maxDegree }, work);
            }
            else
            {
                foreach (var z in list) work(z);
            }

            sw.Stop();
            var ordered = bag.OrderBy(b => b.ShouldSkip).ThenByDescending(b => b.RawSleeveSizeFt).ToList();
            return new SleevePlacementPlanningResult(ordered, skippedCount, highRisk, criticalRisk, sw.ElapsedMilliseconds);
        }

        /// <summary>
        /// ✅ FIX #1: Calculate damper connector-side offset vector for asymmetric clearance.
        /// EXACTLY REPLICATES DamperPlacementStrategy.GetDamperPlacementAdjustment() logic.
        /// Uses DB properties: HasMepConnector, DamperConnectorSide, HostOrientation, StructuralElementType
        /// 
        /// Offset = (MEP Clearance - Other Clearance) / 2
        /// For 100mm MEP, 50mm Other: (100 - 50) / 2 = 25mm toward connector side
        /// </summary>
        private XYZ CalculateDamperOffsetVector(ClashZone zone)
        {
            if (zone == null) return XYZ.Zero;
            
            // ✅ CHECK: Does this damper have a MEP connector? (DB property)
            if (!zone.HasMepConnector || string.IsNullOrEmpty(zone.DamperConnectorSide))
            {
                // No connector = symmetric clearance = no offset
                return XYZ.Zero;
            }
            
            // ✅ GET CLEARANCES: From conditions (same as DamperPlacementStrategy)
            double mepClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0;
            double otherClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryOtherNormal ?? 50.0;
            
            // Note: DamperPlacementStrategy forces isInsulated=false for dampers
            // So we don't check zone.IsInsulated here
            
            // ✅ CALCULATE OFFSET: (MEP - Other) / 2
            double offsetAmountMm = (mepClearanceMm - otherClearanceMm) / 2.0;
            double offsetAmountFt = offsetAmountMm / 304.8;
            
            // If clearances are equal, no offset needed
            if (Math.Abs(offsetAmountFt) < 0.0001)
                return XYZ.Zero;
            
            // ✅ GET WALL ORIENTATION: From DB (same as DamperPlacementStrategy)
            string hostOrientation = zone.HostOrientation ?? string.Empty;
            bool isXWall = string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase);
            bool isYWall = string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase);
            
            // ✅ GET CONNECTOR DIRECTION: From DB (world coordinate: "+X", "-X", "+Y", "-Y", "+Z", "-Z")
            string connectorDir = zone.DamperConnectorSide ?? string.Empty;
            
            // ✅ CHECK: Only apply offset for walls (not floors/framing) - same as DamperPlacementStrategy
            bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
            
            XYZ offsetVector = XYZ.Zero;
            
            if (isWallHost && offsetAmountFt > 0.0001)
            {
                // ✅ EXACT COPY OF DamperPlacementStrategy SWITCH STATEMENT:
                switch (connectorDir)
                {
                    case "+X":
                        if (isXWall)
                            offsetVector = new XYZ(offsetAmountFt, 0, 0);
                        else if (isYWall)
                            offsetVector = new XYZ(0, offsetAmountFt, 0);
                        break;
                    
                    case "-X":
                        if (isXWall)
                            offsetVector = new XYZ(-offsetAmountFt, 0, 0);
                        else if (isYWall)
                            offsetVector = new XYZ(0, -offsetAmountFt, 0);
                        break;
                    
                    case "+Y":
                        if (isYWall)
                            offsetVector = new XYZ(0, offsetAmountFt, 0);
                        else if (isXWall)
                            offsetVector = new XYZ(offsetAmountFt, 0, 0);
                        break;
                    
                    case "-Y":
                        if (isYWall)
                            offsetVector = new XYZ(0, -offsetAmountFt, 0);
                        else if (isXWall)
                            offsetVector = new XYZ(-offsetAmountFt, 0, 0);
                        break;
                    
                    case "+Z":
                        offsetVector = new XYZ(0, 0, offsetAmountFt);
                        break;
                    
                    case "-Z":
                        offsetVector = new XYZ(0, 0, -offsetAmountFt);
                        break;
                    
                    default:
                        // Backward compatibility for old format ("Left", "Right", etc.)
                        bool isLeftRight = connectorDir == "Left" || connectorDir == "Right";
                        bool isTopBottom = connectorDir == "Top" || connectorDir == "Bottom";
                        bool offsetPositive = connectorDir == "Right" || connectorDir == "Top";
                        
                        if (isLeftRight)
                        {
                            if (isXWall)
                                offsetVector = offsetPositive ? new XYZ(offsetAmountFt, 0, 0) : new XYZ(-offsetAmountFt, 0, 0);
                            else if (isYWall)
                                offsetVector = offsetPositive ? new XYZ(0, offsetAmountFt, 0) : new XYZ(0, -offsetAmountFt, 0);
                        }
                        else if (isTopBottom)
                        {
                            offsetVector = offsetPositive ? new XYZ(0, 0, offsetAmountFt) : new XYZ(0, 0, -offsetAmountFt);
                        }
                        break;
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("planner_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PLANNER] ✅ DAMPER OFFSET: Zone {zone.Id}, " +
                    $"HasMepConnector={zone.HasMepConnector}, ConnectorDir='{connectorDir}', " +
                    $"HostOrientation='{hostOrientation}', IsWallHost={isWallHost}, " +
                    $"MEP={mepClearanceMm:F1}mm, Other={otherClearanceMm:F1}mm, " +
                    $"OffsetAmount={offsetAmountMm:F2}mm, " +
                    $"OffsetVector=({offsetVector.X*304.8:F2}, {offsetVector.Y*304.8:F2}, {offsetVector.Z*304.8:F2})mm\n");
            }
            
            return offsetVector;
        }

        /// <summary>
        /// ✅ FIX #2: Calculate damper dimensions with asymmetric clearance.
        /// EXACTLY REPLICATES DamperPlacementStrategy.GetDamperPlacementAdjustment() logic.
        /// Uses DB properties: HasMepConnector, DamperConnectorSide, HostOrientation
        /// 
        /// For MSFD (with connector): MEP clearance on connector side, Other clearance on other sides
        /// For Standard (no connector): Other clearance on all sides (symmetric)
        /// Then applies rounding using RoundAlwaysUp setting (like NewSleevePlacerService)
        /// </summary>
        private (double finalWidth, double finalHeight) CalculateDamperDimensionsWithAsymmetricClearance(ClashZone zone)
        {
            double damperWidth = zone.MepElementWidth;
            double damperHeight = zone.MepElementHeight;
            
            // ✅ GET CLEARANCES: From conditions (same as DamperPlacementStrategy)
            double mepClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0;
            double otherClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryOtherNormal ?? 50.0;
            
            // Convert to feet
            double mepClearanceFt = mepClearanceMm / 304.8;
            double otherClearanceFt = otherClearanceMm / 304.8;
            
            // DamperPlacementStrategy forces insulation to false for dampers
            double insulationContribution = 0.0;
            
            double finalWidth, finalHeight;
            
            // ✅ CHECK: Does this damper have a MEP connector? (DB property)
            if (zone.HasMepConnector && !string.IsNullOrEmpty(zone.DamperConnectorSide))
            {
                // MSFD Damper: Asymmetric clearance
                double left = otherClearanceFt, right = otherClearanceFt;
                double top = otherClearanceFt, bottom = otherClearanceFt;
                
                string connectorDir = zone.DamperConnectorSide ?? string.Empty;
                string hostOrientation = zone.HostOrientation ?? string.Empty;
                bool isXWall = string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase);
                bool isYWall = string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase);
                bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                
                // ✅ WALL Z-CONNECTOR FIX: Swap base dimensions for vertical connector (same as DamperPlacementStrategy)
                if (isWallHost && (connectorDir == "+Z" || connectorDir == "-Z"))
                {
                    double originalWidth = damperWidth;
                    damperWidth = damperHeight;
                    damperHeight = originalWidth;
                }
                
                // ✅ MAP CONNECTOR DIRECTION TO CLEARANCE SIDE (exact copy of DamperPlacementStrategy)
                switch (connectorDir)
                {
                    case "+X":
                        right = mepClearanceFt;
                        break;
                    case "-X":
                        left = mepClearanceFt;
                        break;
                    case "+Y":
                        right = mepClearanceFt;
                        break;
                    case "-Y":
                        left = mepClearanceFt;
                        break;
                    case "+Z":
                        top = mepClearanceFt;
                        break;
                    case "-Z":
                        bottom = mepClearanceFt;
                        break;
                    // Backward compatibility
                    case "Right":
                        right = mepClearanceFt;
                        break;
                    case "Left":
                        left = mepClearanceFt;
                        break;
                    case "Top":
                        top = mepClearanceFt;
                        break;
                    case "Bottom":
                        bottom = mepClearanceFt;
                        break;
                }
                
                // ✅ CALCULATE FINAL DIMENSIONS (same formula as DamperPlacementStrategy line 616-617)
                finalWidth = damperWidth + insulationContribution + left + right;
                finalHeight = damperHeight + insulationContribution + top + bottom;
            }
            else
            {
                // Standard Damper: Symmetric clearance (Other clearance on all sides)
                finalWidth = damperWidth + insulationContribution + (2 * otherClearanceFt);
                finalHeight = damperHeight + insulationContribution + (2 * otherClearanceFt);
            }
            
            // ✅ APPLY ROUNDING: Same as NewSleevePlacerService lines 1595-1597
            // Uses RoundAlwaysUp setting for consistent behavior
            var roundingSettings = JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService.Instance.GetCurrentSettings();
            double roundingValueMm = roundingSettings.RoundingValue;
            bool roundAlwaysUp = roundingSettings.RoundAlwaysUp;
            double roundingValueFt = roundingValueMm / 304.8;
            
            double roundedWidth, roundedHeight;
            if (roundAlwaysUp)
            {
                roundedWidth = Math.Ceiling(finalWidth / roundingValueFt) * roundingValueFt;
                roundedHeight = Math.Ceiling(finalHeight / roundingValueFt) * roundingValueFt;
            }
            else
            {
                roundedWidth = Math.Round(finalWidth / roundingValueFt) * roundingValueFt;
                roundedHeight = Math.Round(finalHeight / roundingValueFt) * roundingValueFt;
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("planner_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PLANNER] ✅ DAMPER DIMS: Zone {zone.Id}, " +
                    $"HasMepConnector={zone.HasMepConnector}, ConnectorDir='{zone.DamperConnectorSide}', " +
                    $"Raw=({damperWidth*304.8:F1}, {damperHeight*304.8:F1})mm, " +
                    $"MEP={mepClearanceMm:F1}mm, Other={otherClearanceMm:F1}mm, " +
                    $"PreRound=({finalWidth*304.8:F1}, {finalHeight*304.8:F1})mm, " +
                    $"RoundingValue={roundingValueMm}mm, RoundAlwaysUp={roundAlwaysUp}, " +
                    $"Rounded=({roundedWidth*304.8:F1}, {roundedHeight*304.8:F1})mm\n");
            }
            
            return (roundedWidth, roundedHeight);
        }

        public SleevePlacementPlanningDto PlanSingle(ClashZone zone)
        {
            try
            {
                if (zone == null) return null;

                // 1. Basic skip rules (no mutation of original object)
                bool shouldSkip = false;
                string skipReason = null;
                if (zone.IsResolved && zone.SleeveInstanceId > 0) { shouldSkip = true; skipReason = "AlreadyResolved"; }
                else if (zone.IsClusterResolved && zone.ClusterSleeveInstanceId > 0) { shouldSkip = true; skipReason = "ClusterResolved"; }
                else if (zone.HasDamperNearby) { shouldSkip = true; skipReason = "DamperSkip"; }

                // 2. MEP Metadata & Categories
                string mepCategory = zone.MepElementCategory ?? "Unknown";
                string hostType = zone.StructuralElementType ?? "Wall";
                bool isPipesCategory = string.Equals(mepCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                bool isDuctsCategory = string.Equals(mepCategory, "Ducts", StringComparison.OrdinalIgnoreCase);
                
                // ✅ FIX: Detect damper category
                bool isDamperCategory = string.Equals(mepCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);

                double rawWidth = zone.MepElementOuterDiameter > 0 ? zone.MepElementOuterDiameter : zone.MepElementWidth;
                double rawHeight = zone.MepElementHeight > 0 ? zone.MepElementHeight : rawWidth;
                double rawDiameter = isPipesCategory ? rawWidth : Math.Max(rawWidth, rawHeight);

                // 3. Resolve Opening Type (Circular vs Rectangular) - CENTRALIZED RULE
                // ✅ REUSE: ConfigurationResolutionService is the absolute source of truth for rules
                var elementProps = new JSE_RevitAddin_MEP_OPENINGS.Services.Configuration.ElementProperties
                {
                    Diameter = rawDiameter,
                    Width = rawWidth,
                    Height = rawHeight,
                    Shape = (isPipesCategory || (isDuctsCategory && (
                        string.Equals(zone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) || 
                        string.Equals(zone.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase) ||
                        (zone.MepElementSizeData != null && (
                            string.Equals(zone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(zone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)
                        ))
                    ))) ? "Round" : "Rectangular",
                    IsInsulated = zone.IsInsulated,
                    InsulationThickness = zone.InsulationThickness
                };

                string uiPreference = "Rectangular";
                double pertinentClearanceMm = 0.0;

                if (isPipesCategory) 
                {
                    uiPreference = _conditions?.OpeningTypePreferences?.Pipes ?? "Circular";
                    pertinentClearanceMm = _conditions?.ClearanceSettings?.PipesNormal ?? 50.0;
                }
                else if (isDuctsCategory) 
                {
                    // ✅ CRITICAL FIX: Only apply "RoundDucts" preference if the duct is actually Round/Circular
                    // Rectangular ducts should naturally default to "Rectangular"
                    bool isRoundDuct = string.Equals(zone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) || 
                                      string.Equals(zone.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase) ||
                                      (zone.MepElementSizeData != null && (
                                          string.Equals(zone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(zone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)
                                      ));
                                      
                    if (isRoundDuct)
                    {
                        uiPreference = _conditions?.OpeningTypePreferences?.RoundDucts ?? "Circular";
                        pertinentClearanceMm = _conditions?.ClearanceSettings?.RoundNormal ?? 50.0;
                    }
                    else
                    {
                        uiPreference = "Rectangular"; // Rectangular ducts are always rectangular preference
                    }
                }
                
                // Fix for user request: Calculate total diameter (MEP + 2*Clearance) inside ResolveOpeningType
                string resolvedType = JSE_RevitAddin_MEP_OPENINGS.Services.Configuration.ConfigurationResolutionService.Instance
                    .ResolveOpeningType(mepCategory, elementProps, uiPreference, hostType, pertinentClearanceMm);
                
                bool isCircular = string.Equals(resolvedType, "Circular", StringComparison.OrdinalIgnoreCase);

                // 4. Resolve Dimensions - CENTRALIZED SIZING (with rounding per user settings)
                double targetW, targetH;
                double clearance;
                
                // ✅ FIX: Special handling for dampers with asymmetric clearance
                if (isDamperCategory)
                {
                    // Use asymmetric clearance calculation for dampers
                    var (damperWidth, damperHeight) = CalculateDamperDimensionsWithAsymmetricClearance(zone);
                    targetW = damperWidth;
                    targetH = damperHeight;
                    
                    // Get average clearance for risk classification
                    double mepClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryMepNormal ?? 100.0;
                    double otherClearanceMm = _conditions?.ClearanceSettings?.DuctAccessoryOtherNormal ?? 50.0;
                    clearance = ((mepClearanceMm + otherClearanceMm) / 2.0) / 304.8; // Average clearance in feet
                }
                else
                {
                    // Standard sizing for non-dampers
                    clearance = GetClearanceFromConditions(mepCategory, rawDiameter * 304.8);
                    var roundingSettings = JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService.Instance.GetCurrentSettings();
                    var dims = _sizingService.CalculateFinalDimensionsFromClashZoneRounded(
                        rawWidth, rawHeight, rawDiameter, zone, clearance, roundingSettings.RoundingValue, roundingSettings.RoundAlwaysUp);
                    targetW = dims.Item1;
                    targetH = dims.Item2;
                }

                // 5. Resolve Family Name - CENTRALIZED MAPPING
                // ✅ REUSE: FamilyManager maps host+shape to the core 4 families
                string familyName = FamilyManager.SelectUniversalFamily(hostType, resolvedType);

                // 6. Resolve Rotation - CENTRALIZED ROTATION
                // ✅ REUSE: SleeveRotationService handles X/Y Wall and Floor heuristics
                double rotationRad = _rotationService.DetermineRotation(zone);

                // 7. Resolve Placement Point - CENTRALIZED COORDINATES
                // ✅ FIX #1: For dampers, apply connector-side offset to placement point.
                // ✅ FIX #2: Do NOT double-apply offset when re-using a DB-saved placement point.
                //
                // RULE:
                // - If SleevePlacementPoint* is zero → we are planning from the raw intersection → apply damper offset.
                // - If SleevePlacementPoint* is non-zero → it already includes any previous offset → DO NOT add it again.
                bool hasSavedPlacementPoint =
                    Math.Abs(zone.SleevePlacementPointX) > 1e-6 ||
                    Math.Abs(zone.SleevePlacementPointY) > 1e-6 ||
                    Math.Abs(zone.SleevePlacementPointZ) > 1e-6;

                XYZ placementPoint = new XYZ(
                    hasSavedPlacementPoint ? zone.SleevePlacementPointX : zone.IntersectionPointX,
                    hasSavedPlacementPoint ? zone.SleevePlacementPointY : zone.IntersectionPointY,
                    hasSavedPlacementPoint ? zone.SleevePlacementPointZ : zone.IntersectionPointZ);
                
                // ✅ CRITICAL FIX: Apply damper offset for asymmetric clearance ONLY once (on fresh points).
                if (isDamperCategory && !hasSavedPlacementPoint)
                {
                    XYZ damperOffset = CalculateDamperOffsetVector(zone);
                    if (damperOffset.GetLength() > 0.0001) // Only apply if significant
                    {
                        placementPoint = placementPoint + damperOffset;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("planner_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [PLANNER] ✅ APPLIED DAMPER OFFSET: Zone {zone.Id}, " +
                                $"Offset=({damperOffset.X*304.8:F2}, {damperOffset.Y*304.8:F2}, {damperOffset.Z*304.8:F2})mm, " +
                                $"Final=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6}) (fresh intersection point)\n");
                        }
                    }
                }

                // 8. Host Thickness & Depth — structural thickness for all (Floor, Wall, Framing)
                double hostThickness = zone.StructuralElementThickness;
                if (hostThickness <= 0) hostThickness = 0.5; // Fallback 6"
                double sleeveDepthFt = zone.StructuralElementThickness > 0.001 ? zone.StructuralElementThickness : hostThickness;

                // 9. Risk assessment
                var risk = ClassifyRisk(hostThickness, clearance, rawDiameter);

                string logLine = $"PLAN ClashZone={zone.Id} Host={hostType} Cat={mepCategory} Type={resolvedType} Family={familyName} Skip={shouldSkip}";

                return new SleevePlacementPlanningDto(
                    zone.Id,
                    hostType,
                    rawDiameter,
                    zone.InsulationThickness,
                    targetW,
                    targetH,
                    clearance,
                    sleeveDepthFt,
                    rotationRad,
                    risk,
                    shouldSkip,
                    skipReason,
                    logLine,
                    familyName,
                    placementPoint,
                    isCircular);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    SafeFileLogger.SafeAppendText("planner_errors.log", $"[{DateTime.Now:HH:mm:ss}] Error planning zone {zone?.Id}: {ex.Message}\n");
                return null;
            }
        }

        private static ClearanceRiskClassification ClassifyRisk(double hostThickness, double clearance, double rawSize)
        {
            if (rawSize <= 0 || hostThickness <= 0) return ClearanceRiskClassification.Unknown;
            double ratio = clearance / rawSize; // relative clearance
            if (ratio < 0.08) return ClearanceRiskClassification.Critical;  // < ~1" on 1ft size
            if (ratio < 0.10) return ClearanceRiskClassification.High;
            if (ratio < 0.15) return ClearanceRiskClassification.Medium;
            return ClearanceRiskClassification.Low;
        }

        /// <summary>
        /// Get clearance for duct based on ClashZone.IsInsulated (uses authoritative data from database)
        /// This bypasses mepSize which might be stale and uses clashZone.IsInsulated directly
        /// </summary>
        private double GetClearanceForDuctFromClashZone(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                
                // ✅ PRIORITY 1: Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    string normalKey = "ducts_normal_clearance";
                    string insulatedKey = "ducts_insulated_clearance";
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearanceMm = _clearanceSettings[targetKey];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                    
                    // Fallback to generic duct clearance
                    if (_clearanceSettings.ContainsKey("ducts_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["ducts_clearance"];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions (check shape for round vs rectangular)
                if (_conditions?.ClearanceSettings != null)
                {
                    // Check if round or rectangular (default to rectangular if unknown)
                    bool isRound = clashZone?.MepElementSizeData?.Shape?.Equals("Round", StringComparison.OrdinalIgnoreCase) == true ||
                                   clashZone?.MepElementSizeData?.Shape?.Equals("Circular", StringComparison.OrdinalIgnoreCase) == true;
                    
                    double clearanceInMm;
                    if (isRound)
                    {
                        clearanceInMm = isInsulated 
                            ? _conditions.ClearanceSettings.RoundInsulated 
                            : _conditions.ClearanceSettings.RoundNormal;
                    }
                    else
                    {
                        clearanceInMm = isInsulated 
                            ? _conditions.ClearanceSettings.RectangularInsulated 
                            : _conditions.ClearanceSettings.RectangularNormal;
                    }
                    
                    // Convert mm to feet (internal units)
                    return clearanceInMm / 304.8;
                }
                
                // ✅ PRIORITY 3: Default fallback (50mm)
                return 50.0 / 304.8; // 50mm in feet
            }
            catch
            {
                // Fallback on error
                return 50.0 / 304.8; // 50mm in feet
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL FIX: Get ASYMMETRIC clearances for cable trays (Top and Other separately)
        /// Returns both topClearanceFt and otherClearanceFt for proper asymmetric calculation
        /// </summary>
        private (double topClearanceFt, double otherClearanceFt) GetCableTrayAsymmetricClearances(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                double topClearanceMm = 100.0; // Default fallback
                double otherClearanceMm = 50.0; // Default fallback
                
                // ✅ PRIORITY 1: Read from database (conditions object loaded from SQLite)
                if (_conditions?.ClearanceSettings != null)
                {
                    topClearanceMm = _conditions.ClearanceSettings.CableTrayTop;
                    otherClearanceMm = _conditions.ClearanceSettings.CableTrayOther;
                }
                // ✅ PRIORITY 2: Fallback to UI clearance settings
                else if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    // Check for insulated cable tray clearance first
                    if (isInsulated && _clearanceSettings.ContainsKey("cabletray_top_insulated"))
                    {
                        topClearanceMm = _clearanceSettings["cabletray_top_insulated"];
                    }
                    else if (_clearanceSettings.ContainsKey("cabletray_top_normal"))
                    {
                        topClearanceMm = _clearanceSettings["cabletray_top_normal"];
                    }
                    else if (_clearanceSettings.ContainsKey("cabletray_top_clearance"))
                    {
                        topClearanceMm = _clearanceSettings["cabletray_top_clearance"];
                    }
                    
                    if (isInsulated && _clearanceSettings.ContainsKey("cabletray_other_insulated"))
                    {
                        otherClearanceMm = _clearanceSettings["cabletray_other_insulated"];
                    }
                    else if (_clearanceSettings.ContainsKey("cabletray_other_normal"))
                    {
                        otherClearanceMm = _clearanceSettings["cabletray_other_normal"];
                    }
                    else if (_clearanceSettings.ContainsKey("cabletray_other_clearance"))
                    {
                        otherClearanceMm = _clearanceSettings["cabletray_other_clearance"];
                    }
                }
                
                // Convert mm to feet (internal units)
                double topClearanceFt = topClearanceMm / 304.8;
                double otherClearanceFt = otherClearanceMm / 304.8;
                
                return (topClearanceFt, otherClearanceFt);
            }
            catch
            {
                // Fallback on error (50mm default)
                return (50.0 / 304.8, 50.0 / 304.8);
            }
        }
        
        /// <summary>
        /// Get clearance for cable tray based on ClashZone.IsInsulated (uses authoritative data from database)
        /// This bypasses mepSize which might be stale and uses clashZone.IsInsulated directly
        /// ⚠️ DEPRECATED: Use GetCableTrayAsymmetricClearances instead for proper asymmetric clearance handling
        /// </summary>
        private double GetClearanceForCableTrayFromClashZone(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                
                // ✅ PRIORITY 1: Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    // Check for insulated cable tray clearance first
                    if (isInsulated && _clearanceSettings.ContainsKey("cabletray_insulated_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["cabletray_insulated_clearance"];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                    
                    // Fallback to top clearance (primary clearance for cable trays)
                    if (_clearanceSettings.ContainsKey("cabletray_top_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["cabletray_top_clearance"];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                    
                    // Fallback to other clearance
                    if (_clearanceSettings.ContainsKey("cabletray_other_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["cabletray_other_clearance"];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                    // Use top clearance as default (primary clearance for cable trays)
                    double clearanceInMm = _conditions.ClearanceSettings.CableTrayTop;
                    
                    // Convert mm to feet (internal units)
                    return clearanceInMm / 304.8;
                }
                
                // ✅ PRIORITY 3: Default fallback (50mm)
                return 50.0 / 304.8; // 50mm in feet
            }
            catch
            {
                // Fallback on error
                return 50.0 / 304.8; // 50mm in feet
            }
        }
        
        /// <summary>
        /// Get clearance for pipe based on ClashZone.IsInsulated (uses authoritative data from database)
        /// This bypasses mepSize which might be stale and uses clashZone.IsInsulated directly
        /// </summary>
        private double GetClearanceForPipeFromClashZone(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                
                // ✅ PRIORITY 1: Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    string normalKey = "pipes_normal_clearance";
                    string insulatedKey = "pipes_insulated_clearance";
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearanceMm = _clearanceSettings[targetKey];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                    
                    // Fallback to generic pipe clearance
                    if (_clearanceSettings.ContainsKey("pipes_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["pipes_clearance"];
                        // Convert mm to feet (internal units)
                        return clearanceMm / 304.8;
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                    double clearanceInMm = isInsulated 
                        ? _conditions.ClearanceSettings.PipesInsulated 
                        : _conditions.ClearanceSettings.PipesNormal;
                    
                    // Convert mm to feet (internal units)
                    return clearanceInMm / 304.8;
                }
                
                // ✅ PRIORITY 3: Default fallback (50mm)
                return 50.0 / 304.8; // 50mm in feet
            }
            catch
            {
                // Fallback on error
                return 50.0 / 304.8; // 50mm in feet
            }
        }
        
        /// <summary>
        /// Get clearance from OpeningConditions based on MEP category and size.
        /// Simplified version that matches UniversalSleevePlacerService logic.
        /// </summary>
        private double GetClearanceFromConditions(string category, double mepSizeInMm)
        {
            try
            {
                // 1. Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    // Direct category lookup
                    if (_clearanceSettings.TryGetValue(category, out double clearanceMm) && clearanceMm > 0)
                    {
                        // Convert mm to feet
                        return clearanceMm / 304.8;
                    }
                }

                // 2. Fall back to OpeningConditions XML defaults
                if (_conditions?.ClearanceSettings != null)
                {
                    double clearanceMm = 50.0; // default
                    
                    // Get clearance based on category
                    if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                    {
                        clearanceMm = _conditions.ClearanceSettings.RectangularNormal;
                    }
                    else if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                    {
                        // ⚠️ NOTE: This method should not be called for pipes - use GetClearanceForPipeFromClashZone instead
                        clearanceMm = _conditions.ClearanceSettings.PipesNormal;
                    }
                    else if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                    {
                        clearanceMm = _conditions.ClearanceSettings.CableTrayTop;
                    }
                    else if (string.Equals(category, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                    {
                        // ✅ FIX: Duct Accessories (dampers) should use average of MEP and Other clearance
                        double mepClearance = _conditions.ClearanceSettings.DuctAccessoryMepNormal;
                        double otherClearance = _conditions.ClearanceSettings.DuctAccessoryOtherNormal;
                        clearanceMm = (mepClearance + otherClearance) / 2.0;
                    }
                    
                    // Convert mm to feet
                    return clearanceMm / 304.8;
                }

                // 3. Fallback: 50mm (2") default clearance
                return 50.0 / 304.8; // 50mm in feet
            }
            catch
            {
                // Fallback on error
                return 50.0 / 304.8; // 50mm in feet
            }
        }
    }
}

/*
USAGE (to be wired later inside UniversalSleevePlacerService without modifying existing logic yet):

    // 1. Instantiate once (could be cached)
    var planner = new ParallelSleevePlacementPlanner();

    // 2. Run planning on the clash zones BEFORE sequential placement:
    var planningResult = planner.Plan(clashZones);

    // 3. Use planningResult.Items for ordered processing:
    foreach (var dto in planningResult.Items)
    {
        if (dto.ShouldSkip) continue; // Skip early
        // Use dto.TargetWidthFt, dto.TargetHeightFt, dto.RequiredDepthFt, dto.RotationAngleDeg
        // to drive sleeve family selection and parameter assignment.
    }

    // 4. Batch log:
    foreach (var line in planningResult.Items.Select(i => i.LogSummary))
    {
        SafeFileLogger.SafeAppendText("planning_debug.log", line + "\n");
    }

    // 5. Metrics (planningResult.HighRiskCount, etc.) can feed performance report.
*/
