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

                // 4. Resolve Dimensions - CENTRALIZED SIZING
                // ✅ REUSE: InsulationAwareSizingService handles insulation + clearance + rounding
                double clearance = GetClearanceFromConditions(mepCategory, rawDiameter * 304.8);
                
                var (targetW, targetH, _) = _sizingService.CalculateFinalDimensionsFromClashZone(
                    rawWidth, rawHeight, rawDiameter, zone, clearance);

                // 5. Resolve Family Name - CENTRALIZED MAPPING
                // ✅ REUSE: FamilyManager maps host+shape to the core 4 families
                string familyName = FamilyManager.SelectUniversalFamily(hostType, resolvedType);

                // 6. Resolve Rotation - CENTRALIZED ROTATION
                // ✅ REUSE: SleeveRotationService handles X/Y Wall and Floor heuristics
                double rotationRad = _rotationService.DetermineRotation(zone);

                // 7. Resolve Placement Point - CENTRALIZED COORDINATES
                // ✅ REUSE: Prioritize SleevePlacementPoint (set during refresh/adjustment) over IntersectionPoint
                XYZ placementPoint = new XYZ(
                    zone.SleevePlacementPointX != 0 ? zone.SleevePlacementPointX : zone.IntersectionPointX,
                    zone.SleevePlacementPointY != 0 ? zone.SleevePlacementPointY : zone.IntersectionPointY,
                    zone.SleevePlacementPointZ != 0 ? zone.SleevePlacementPointZ : zone.IntersectionPointZ);

                // 8. Host Thickness & Depth
                double hostThickness = zone.StructuralElementThickness;
                if (hostThickness <= 0) hostThickness = 0.5; // Fallback 6"
                double requiredDepth = hostThickness + (clearance * 2); // Internal units (feet)

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
                    requiredDepth,
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
