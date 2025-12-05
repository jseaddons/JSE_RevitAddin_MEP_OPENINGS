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
            int skipped = 0; int highRisk = 0; int criticalRisk = 0;

            Action<ClashZone> work = zone =>
            {
                try
                {
                    if (zone == null)
                    {
                        System.Threading.Interlocked.Increment(ref skipped);
                        return;
                    }

                    // Basic skip rules (no mutation of original object)
                    string skipReason = null;
                    if (zone.IsResolved && zone.SleeveInstanceId > 0) skipReason = "AlreadyResolved";
                    else if (zone.IsClusterResolved && zone.ClusterSleeveInstanceId > 0) skipReason = "ClusterResolved";
                    else if (zone.HasDamperNearby) skipReason = "DamperSkip";

                    // Host type string
                    string hostType = zone.StructuralElementType ?? string.Empty;

                    // Raw MEP size (convert from internal units to feet)
                    // ✅ CRITICAL FIX: Use MepElementWidth/Height from ClashZone (already in database)
                    // These are populated during refresh and persisted, unlike MepElementSizeData
                    // ✅ PIPE FIX: For pipes, MepElementWidth should be OUTER DIAMETER (set during refresh)
                    string mepCategory = zone.MepElementCategory ?? "Unknown";
                    bool isPipesCategory = string.Equals(mepCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                    
                    double rawWidth = zone.MepElementWidth;  // Already in feet
                    double rawHeight = zone.MepElementHeight; // Already in feet
                    
                    // ✅ PIPE FIX: For pipes, use OUTER DIAMETER from database column (hardcoded route)
                    // ✅ HARDCODED: Always use MepElementOuterDiameter (RBS_PIPE_OUTER_DIAMETER) for accurate sizing
                    if (isPipesCategory)
                    {
                        // For pipes, use outer diameter from database column (preferred)
                        double pipeDiameter = zone.MepElementOuterDiameter > 0 
                            ? zone.MepElementOuterDiameter 
                            : (rawWidth > 0 ? rawWidth : rawHeight); // Fallback to MepElementWidth if outer diameter not available
                        
                        // Additional fallback to MepElementSizeData if both are missing
                        if (pipeDiameter <= 0 && zone.MepElementSizeData != null)
                        {
                            // Try to get diameter from MepElementSizeData
                            if (zone.MepElementSizeData.Diameter > 0)
                            {
                                pipeDiameter = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Diameter, UnitTypeId.Feet);
                            }
                        }
                        // Final fallback
                        if (pipeDiameter <= 0) pipeDiameter = zone.MepElementSize;
                        if (pipeDiameter <= 0) pipeDiameter = 0.25; // Minimum fallback 3"
                        
                        // For pipes, width = height = diameter (round pipes)
                        rawWidth = pipeDiameter;
                        rawHeight = pipeDiameter;
                    }
                    
                    double rawSize = Math.Max(rawWidth, rawHeight); // Use larger dimension
                    
                    // Fallback to MepElementSizeData if Width/Height are 0 (backward compatibility, non-pipes)
                    if (rawSize <= 0 && !isPipesCategory && zone.MepElementSizeData != null)
                    {
                        // MepElementSize stores in internal units, convert to feet
                        if (zone.MepElementSizeData.Diameter > 0)
                        {
                            rawSize = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Diameter, UnitTypeId.Feet);
                            rawWidth = rawSize;
                            rawHeight = rawSize;
                        }
                        else if (zone.MepElementSizeData.Width > 0)
                        {
                            rawWidth = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Width, UnitTypeId.Feet);
                            rawHeight = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Height, UnitTypeId.Feet);
                            rawSize = Math.Max(rawWidth, rawHeight);
                        }
                    }
                    // Final fallback to MepElementSize (already in feet)
                    if (rawSize <= 0) rawSize = zone.MepElementSize;
                    if (rawSize <= 0) rawSize = 0.25; // Minimum fallback 3"

                    // ✅ OOP METHOD: Get clearance from OpeningConditions (same logic as UniversalSleevePlacerService)
                    // ✅ CRITICAL FIX: Use ClashZone.IsInsulated for correct clearance selection (pipes, ducts, cable trays)
                    double clearance;
                    double clearanceInFeet;
                    double targetWidth = 0.0;
                    double targetHeight = 0.0;
                    
                    if (isPipesCategory)
                    {
                        clearance = GetClearanceForPipeFromClashZone(zone);
                    }
                    else if (string.Equals(mepCategory, "Ducts", StringComparison.OrdinalIgnoreCase))
                    {
                        clearance = GetClearanceForDuctFromClashZone(zone);
                    }
                    else if (string.Equals(mepCategory, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                    {
                        // ✅ CRITICAL FIX: Cable trays need ASYMMETRIC clearances (Top=75mm, Other=25mm)
                        // Width: 2 * OtherClearance (25mm * 2 = 50mm total)
                        // Height: TopClearance + OtherClearance (75mm + 25mm = 100mm total)
                        // CANNOT use single clearance value - must calculate separately
                        var cableTrayClearances = GetCableTrayAsymmetricClearances(zone);
                        double topClearanceFt = cableTrayClearances.topClearanceFt;
                        double otherClearanceFt = cableTrayClearances.otherClearanceFt;
                        
                        // ✅ CRITICAL: Calculate cable tray dimensions with asymmetric clearances
                        // Get insulation contribution
                        double insulationContribution = 0.0;
                        if (zone.IsInsulated && zone.InsulationThickness > 0)
                        {
                            insulationContribution = 2 * zone.InsulationThickness; // Both sides
                        }
                        
                        // Convert raw dimensions to internal units (feet) if needed
                        double rawWidthFt = rawWidth; // Already in feet from MepElementSizeData
                        double rawHeightFt = rawHeight; // Already in feet from MepElementSizeData
                        
                        // ✅ ASYMMETRIC CLEARANCE CALCULATION:
                        // Width: Left and right use OTHER clearance (25mm each = 50mm total)
                        // Height: Top uses TOP clearance (75mm), bottom uses OTHER clearance (25mm) = 100mm total
                        targetWidth = rawWidthFt + insulationContribution + (2 * otherClearanceFt);
                        targetHeight = rawHeightFt + insulationContribution + topClearanceFt + otherClearanceFt;
                        
                        // Set clearance to average for risk classification (not used for dimension calculation)
                        clearance = (topClearanceFt + otherClearanceFt) / 2.0;
                        clearanceInFeet = clearance;
                    }
                    else
                    {
                        clearance = GetClearanceFromConditions(mepCategory, rawSize * 304.8);
                        clearanceInFeet = clearance; // Already in internal units (feet)
                        
                        // ✅ OOP METHOD: Use sizing service for consistent calculation with insulation awareness
                        // Formula: Raw + (2 * insulation) + (2 * clearance) - handled by sizing service
                        // For pipes: Pass diameter as all three parameters (width, height, diameter)
                        double rawDiameter = isPipesCategory ? rawWidth : Math.Max(rawWidth, rawHeight);
                        (double targetW, double targetH, _) = _sizingService.CalculateFinalDimensionsFromClashZone(
                            rawWidth, rawHeight, rawDiameter, zone, clearanceInFeet);
                        targetWidth = targetW;
                        targetHeight = targetH;
                    }

                    // ✅ OOP METHOD: Get insulation thickness from ClashZone (already saved to DB during refresh)
                    double insulationThicknessFt = zone.IsInsulated && zone.InsulationThickness > 0 
                        ? zone.InsulationThickness 
                        : 0.0;

                    // Host thickness heuristics
                    double hostThickness = zone.StructuralElementThickness;
                    if (hostThickness <= 0 && hostType.Equals("Wall", StringComparison.OrdinalIgnoreCase)) hostThickness = zone.WallThickness;
                    if (hostThickness <= 0 && hostType.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)) hostThickness = zone.FramingThickness;
                    if (hostThickness <= 0 && hostType.Equals("Floor", StringComparison.OrdinalIgnoreCase)) hostThickness = 0.833; // 10" default
                    if (hostThickness <= 0) hostThickness = 0.5; // Generic fallback 6"

                    double requiredDepth = hostThickness + clearance;

                    // Rotation angle (convert from radians to degrees)
                    double rotationDeg = 0.0;
                    if (Math.Abs(zone.MepElementRotationAngle) > 1e-6)
                    {
                        rotationDeg = zone.MepElementRotationAngle * 180.0 / Math.PI; // Convert radians to degrees
                    }

                    // Risk classification
                    var risk = ClassifyRisk(hostThickness, clearance, rawSize);
                    if (risk == ClearanceRiskClassification.High) System.Threading.Interlocked.Increment(ref highRisk);
                    if (risk == ClearanceRiskClassification.Critical) System.Threading.Interlocked.Increment(ref criticalRisk);

                    bool shouldSkip = skipReason != null;

                    string logLine = $"PLAN ClashZone={zone.Id} Host={hostType} RawSizeFt={rawSize:F3} TargetFt={targetWidth:F3} ClearanceFt={clearance:F3} DepthFt={requiredDepth:F3} Risk={risk} Skip={shouldSkip} Reason={skipReason}";

                    bag.Add(new SleevePlacementPlanningDto(
                        zone.Id,
                        hostType,
                        rawSize,
                        insulationThicknessFt,
                        targetWidth,
                        targetHeight,
                        clearance,
                        requiredDepth,
                        rotationDeg,
                        risk,
                        shouldSkip,
                        skipReason,
                        logLine));

                    if (shouldSkip) System.Threading.Interlocked.Increment(ref skipped);
                }
                catch
                {
                    System.Threading.Interlocked.Increment(ref skipped);
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
            var ordered = bag.OrderBy(b => b.ShouldSkip).ThenByDescending(b => b.RawMepSizeFt).ToList();
            return new SleevePlacementPlanningResult(ordered, skipped, highRisk, criticalRisk, sw.ElapsedMilliseconds);
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
                        // ✅ FIX: Duct Accessories (dampers) should use DuctAccessoryOtherNormal clearance
                        clearanceMm = _conditions.ClearanceSettings.DuctAccessoryOtherNormal;
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
