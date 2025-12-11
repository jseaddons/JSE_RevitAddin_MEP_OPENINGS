using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refactored;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using JSE_RevitAddin_MEP_OPENINGS.Services.Persistence;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For OpeningSettingsHelper

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// New service for placing sleeves, designed to replace UniversalSleevePlacerService.
    /// Implements SOLID principles and supports "Smart Replay" logic.
    /// </summary>
    public class NewSleevePlacerService
    {
        private readonly Document _doc;
        private readonly OpeningConditions _conditions;
        private readonly ISleevePlacementStrategy _strategy;
        private readonly Dictionary<string, double> _clearanceSettings;
        private readonly ISleeveRepository _sleeveRepository;
        private readonly IZoneFilterService _zoneFilterService;
        private readonly IFamilyManager _familyManager;
        private readonly Services.Interfaces.Refactor.IFlagManager _flagManager;
        private readonly bool _isReplayPath;
        private readonly string _filterName;
        
        // ✅ OOP METHOD: Insulation-aware sizing service (SOLID principles)
        private readonly IInsulationAwareSizingService _sizingService;
        
        // ✅ SRP COMPLIANCE: Clearance calculation service (delegates clearance logic)
        private readonly ClearanceCalculationService _clearanceService;
        
        // ✅ SRP COMPLIANCE: RCS bounding box service (delegates clustering geometry logic)
        private readonly RcsBoundingBoxService _rcsBoundingBoxService;
        
        // ✅ SRP COMPLIANCE: Sleeve persistence service (delegates all database persistence operations)
        private readonly SleevePersistenceService _persistenceService;
        
        // ✅ SRP COMPLIANCE: Sleeve rotation service (delegates rotation calculation logic)
        private readonly SleeveRotationService _rotationService;
        
        // ✅ SRP COMPLIANCE: Sleeve parameter service (delegates all parameter setting operations)
        // Note: Not readonly because it needs to be recreated with performance monitor when available
        private SleeveParameterService _parameterService;
        
        // ✅ SRP COMPLIANCE: Placement point adjustment service (delegates all placement point adjustment logic)
        // Note: Not readonly because it needs to be recreated with performance monitor when available
        private PlacementPointAdjustmentService _placementPointAdjustmentService;
        
        // ✅ SOLID REFACTORED: Optional refactored command services (injected when flag enabled)
        private readonly IFileNameNormalizer? _fileNameNormalizer;
        private readonly ISectionBoxChecker? _sectionBoxChecker;
        
        // ✅ CRASH-SAFE: Crash-safe executor for timeout protection and error handling
        private readonly CrashSafeExecutor? _crashSafeExecutor;
        
        // ✅ PERFORMANCE MONITORING: Performance monitor for tracking operations
        private Services.Placement.PlacementPerformanceMonitor? _performanceMonitor;
        
        // Family symbol cache for performance
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();
        
        // ✅ DAMPER CLEARANCE VALUES: Stores clearance values for damper parameter setting
        // Key: ClashZone ID (Guid - for matching zone to its calculated clearances)
        // Value: (finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom, offsetVector)
        // ✅ CRITICAL FIX: offsetVector is now stored to apply connector-side offset during placement
        private Dictionary<Guid, (double finalWidth, double finalHeight, double clearanceLeft, double clearanceRight, double clearanceTop, double clearanceBottom, XYZ offsetVector)> _damperPlacementAdjustments = 
            new Dictionary<Guid, (double, double, double, double, double, double, XYZ)>();
        
        // ✅ PARALLEL PLANNING: Optional planner for parallel pre-computation (OOP, DI-ready)
        // When enabled via DeploymentConfiguration.EnableParallelPlanning, pre-computes dimensions, clearance, and risk in parallel
        // Includes dampers (Duct Accessories) - ParallelSleevePlacementPlanner handles all categories
        private readonly ISleevePlacementPlanner? _planner;
        
        // ✅ FORCE DETECTION MODE: Flag to force recalculation of placement points
        private readonly bool _isForceDetectionMode;

        public NewSleevePlacerService(
            Document doc,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings,
            ISleeveRepository sleeveRepository,
            IZoneFilterService zoneFilterService,
            IFamilyManager familyManager,
            Services.Interfaces.Refactor.IFlagManager? flagManager,
            bool isReplayPath = false,
            string filterName = null,
            IInsulationAwareSizingService sizingService = null,  // ✅ OOP METHOD: Optional sizing service injection (SOLID)
            // ✅ SOLID REFACTORED: Optional refactored command services (injected when flag enabled)
            IFileNameNormalizer? fileNameNormalizer = null,
            ISectionBoxChecker? sectionBoxChecker = null,
            // ✅ CRASH-SAFE: Optional crash-safe executor (created if not provided when flag enabled)
            CrashSafeExecutor? crashSafeExecutor = null,
            // ✅ PARALLEL PLANNING: Optional planner for parallel pre-computation (enabled via safety flag)
            ISleevePlacementPlanner? planner = null,
            // ✅ FORCE DETECTION MODE
            bool isForceDetectionMode = false)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            
            // ✅ DIAGNOSTIC: Log conditions object state at construction
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[NewSleevePlacer] 🔍 CONDITIONS CHECK: conditions={(_conditions != null ? "NOT NULL" : "NULL")}, " +
                    $"ClearanceSettings={(_conditions?.ClearanceSettings != null ? "NOT NULL" : "NULL")}");
                if (_conditions?.ClearanceSettings != null)
                {
                    DebugLogger.Info($"[NewSleevePlacer] 🔍 CLEARANCE VALUES: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, " +
                        $"RoundNormal={_conditions.ClearanceSettings.RoundNormal}mm, " +
                        $"PipesNormal={_conditions.ClearanceSettings.PipesNormal}mm, " +
                        $"CableTrayTop={_conditions.ClearanceSettings.CableTrayTop}mm, " +
                        $"DuctAccessoryMepNormal={_conditions.ClearanceSettings.DuctAccessoryMepNormal}mm, " +
                        $"DuctAccessoryOtherNormal={_conditions.ClearanceSettings.DuctAccessoryOtherNormal}mm, " +
                        $"DuctAccessoryMepInsulated={_conditions.ClearanceSettings.DuctAccessoryMepInsulated}mm, " +
                        $"DuctAccessoryOtherInsulated={_conditions.ClearanceSettings.DuctAccessoryOtherInsulated}mm");
                }
            }
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _sleeveRepository = sleeveRepository ?? throw new ArgumentNullException(nameof(sleeveRepository));
            _zoneFilterService = zoneFilterService; // Can be null for now
            _familyManager = familyManager; // Can be null for now
            _flagManager = flagManager; // Passed as IFlagManager? (can be null)
            _isReplayPath = isReplayPath;
            _filterName = filterName;
            _isForceDetectionMode = isForceDetectionMode;
            
            // ✅ OOP METHOD: Initialize sizing service (create if not provided - Dependency Injection)
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            
            // ✅ SRP COMPLIANCE: Initialize clearance and RCS services (delegate to specialized services)
            _clearanceService = new ClearanceCalculationService();
            _rcsBoundingBoxService = new RcsBoundingBoxService();
            
            // ✅ SRP COMPLIANCE: Initialize persistence service
            _persistenceService = new SleevePersistenceService(doc);
            
            // ✅ SRP COMPLIANCE: Initialize rotation service (delegates wall/floor rotation logic)
            _rotationService = new SleeveRotationService();
            
            // ✅ SRP COMPLIANCE: Initialize parameter service (delegates all parameter setting operations)
            // Note: Performance monitor will be set later in PlaceAllSleevesInTransaction, so we pass null here
            _parameterService = new SleeveParameterService(doc, isReplayPath, null);
            
            // ✅ SRP COMPLIANCE: Initialize placement point adjustment service
            // Note: Performance monitor will be set later in PlaceAllSleevesInTransaction, so we pass null here
            _placementPointAdjustmentService = new PlacementPointAdjustmentService(doc, null, _isForceDetectionMode);
            
            // ✅ SOLID REFACTORED: Initialize refactored services (create if not provided when flag enabled)
            if (OptimizationFlags.UseRefactoredCommandServices)
            {
                _fileNameNormalizer = fileNameNormalizer ?? new Services.Refactored.FileNameNormalizerService();
                _sectionBoxChecker = sectionBoxChecker ?? new Services.Refactored.SectionBoxCheckerService();
            }
            else
            {
                _fileNameNormalizer = fileNameNormalizer;
                _sectionBoxChecker = sectionBoxChecker;
            }
            
            // ✅ CRASH-SAFE: Initialize crash-safe executor (create if not provided when flag enabled)
            if (OptimizationFlags.UseCrashSafeExecution)
            {
                _crashSafeExecutor = crashSafeExecutor ?? new CrashSafeExecutor();
            }
            else
            {
                _crashSafeExecutor = null;
            }
            
            // ✅ PARALLEL PLANNING: Initialize planner (create if not provided and flag enabled)
            // Safety flag: DeploymentConfiguration.EnableParallelPlanning controls whether planner is used
            if (DeploymentConfiguration.EnableParallelPlanning)
            {
                _planner = planner ?? new ParallelSleevePlacementPlanner(
                    conditions,
                    clearanceSettings,
                    minParallelCount: 12, // Minimum zones before parallel processing kicks in
                    maxDegree: null, // Use all CPU cores
                    sizingService: _sizingService);
            }
            else
            {
                _planner = null; // Parallel planning disabled via safety flag
            }
        }

        public (int placed, int skipped, int errors) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            // ✅ SAFE TRANSACTION MANAGEMENT: Validate document state before starting
            if (OptimizationFlags.UseSafeTransactionManagement)
            {
                if (!_doc.IsModifiable)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Document is not modifiable - cannot place sleeves");
                    }
                    return (0, 0, clashZones?.Count ?? 0);
                }
            }
            
            // ✅ PERFORMANCE MONITORING: Initialize performance monitor if enabled
            if (OptimizationFlags.UsePerformanceMonitoring)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string performanceLogName = $"NewSleevePlacer_{timestamp}.log";
                _performanceMonitor = new Services.Placement.PlacementPerformanceMonitor(performanceLogName);
                
                // Recreate services with performance monitor for proper tracking
                _parameterService = new SleeveParameterService(_doc, _isReplayPath, _performanceMonitor);
                _placementPointAdjustmentService = new PlacementPointAdjustmentService(_doc, _performanceMonitor, _isForceDetectionMode);
            }
            
            // ✅ PARAMETER BATCHING: Reset flags at start of each placement run
            _parameterService.ResetFlushFlag();
            
            // ✅ DAMPER PLACEMENT OFFSET: Clear stored offsets at start of each placement run
            _damperPlacementAdjustments?.Clear();
            
            // ✅ DIAGNOSTIC: Log batching flag status at placement start (ALWAYS log, even in deployment mode)
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ═══ PLACEMENT START ═══ UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, Zones={clashZones?.Count ?? 0}\n");
            
            // ✅ DIAGNOSTIC: Log flag status of all zones (ALWAYS log)
            if (clashZones != null && clashZones.Count > 0)
            {
                int resolvedCount = clashZones.Count(z => z.IsResolved);
                int clusterResolvedCount = clashZones.Count(z => z.IsClusterResolved);
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 📊 ZONE FLAGS: Total={clashZones.Count}, IsResolved={resolvedCount}, IsClusterResolved={clusterResolvedCount}\n");
            }
            
            int placed = 0;
            int skipped = 0;
            int errors = 0;
            
            // ✅ CRASH-SAFE: Execute with timeout protection if enabled
            if (OptimizationFlags.UseCrashSafeExecution && _crashSafeExecutor != null)
            {
                try
                {
                    var result = _crashSafeExecutor.ExecuteWithTimeout(() =>
                    {
                        var (p, s, e) = ExecutePlacementInternal(clashZones);
                        placed = p;
                        skipped = s;
                        errors = e;
                        return Result.Succeeded;
                    }, "Place All Sleeves");
                    
                    if (result != Result.Succeeded)
                    {
                        // Operation failed or was cancelled
                        return (placed, skipped, errors);
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Crash-safe execution failed: {ex.Message}");
                    }
                    errors = clashZones?.Count ?? 0;
                    return (placed, skipped, errors);
                }
            }
            else
            {
                // Normal execution without crash-safe wrapper
                (placed, skipped, errors) = ExecutePlacementInternal(clashZones);
            }
            
            // ✅ PERFORMANCE MONITORING: Generate report if enabled
            if (OptimizationFlags.UsePerformanceMonitoring && _performanceMonitor != null)
            {
                _performanceMonitor.GenerateReport(placed, 0); // Individual sleeves only, no clusters
            }
            
            return (placed, skipped, errors);
        }
        
        /// <summary>
        /// ✅ INTERNAL: Core placement logic (extracted for crash-safe wrapper)
        /// </summary>
        private (int placed, int skipped, int errors) ExecutePlacementInternal(List<ClashZone> clashZones)
        {
            int placed = 0;
            int skipped = 0;
            int errors = 0;
            
            // ✅ CRITICAL FIX: Store dimensions with placed sleeves for bounding box calculation when batching is enabled
            // When UseBatchedParameterWrites=true, sleeve.get_BoundingBox() returns STALE values because parameters
            // haven't been flushed yet. We need to calculate bounding boxes from the stored dimensions instead.
            var placedSleeveData = new List<(FamilyInstance sleeve, ClashZone zone, double width, double height, double depth)>();
            var processedZoneGuids = new List<Guid>();
            
            // ✅ DEDUPLICATION: Track placed locations to prevent duplicates
            var placedLocationKeys = new HashSet<string>();

            // ✅ Use injected ZoneFilterService if available to pre-filter zones
            List<ClashZone> filteredZones = clashZones;
            
            // ✅ FAMILY SYMBOL CACHING: Pre-cache family symbols with validation
            if (OptimizationFlags.UseFamilySymbolCache && filteredZones.Count > 0)
            {
                PreCacheFamilySymbols(filteredZones);
            }
            if (_zoneFilterService != null)
            {
                filteredZones = _zoneFilterService.PreFilterEligibleClashZones(_doc, clashZones);
                
                if (!DeploymentConfiguration.DeploymentMode && filteredZones.Count != clashZones.Count)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ZoneFilterService filtered {clashZones.Count} → {filteredZones.Count} zones");
                }
            }
            
            // ✅ PARALLEL PLANNING: Pre-compute all sleeve data in parallel (includes dampers)
            // Safety flag: DeploymentConfiguration.EnableParallelPlanning controls this feature
            // When enabled: Pre-computes dimensions, clearance, rotation, and risk classification in parallel
            // Benefits: Early skip detection, risk-based reordering, parallel computation, better diagnostics
            Dictionary<Guid, SleevePlacementPlanningDto>? planningMap = null;
            SleevePlacementPlanningResult? planningResult = null;
            
            if (DeploymentConfiguration.EnableParallelPlanning && _planner != null && filteredZones.Count > 0)
            {
                using (var planningTracker = _performanceMonitor?.TrackOperation("Parallel Sleeve Placement Planning"))
                {
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] 🚀 PARALLEL PLANNING: Starting parallel pre-computation for {filteredZones.Count} zones (including dampers)");
                        }
                        
                        // ✅ PARALLEL PROCESSING: Run planning in parallel (pure computations, no Revit API calls)
                        // This includes dampers (Duct Accessories) - ParallelSleevePlacementPlanner handles all categories
                        planningResult = _planner.Plan(filteredZones);
                        planningTracker?.SetItemCount(planningResult.TotalCount);
                        
                        // Create lookup map: ClashZoneId -> PlanningDto for fast access during placement
                        planningMap = planningResult.Items
                            .ToDictionary(dto => dto.ClashZoneId, dto => dto);
                        
                        // ✅ EARLY SKIP: Filter out zones marked for skip (avoids Revit API calls)
                        var skipGuids = planningResult.Items
                            .Where(dto => dto.ShouldSkip)
                            .Select(dto => dto.ClashZoneId)
                            .ToHashSet();
                        
                        // Update filtered zones to exclude skipped ones
                        filteredZones = filteredZones
                            .Where(cz => !skipGuids.Contains(cz.Id))
                            .ToList();
                        
                        // ✅ REORDERING: Sort by risk (low risk first, high risk last)
                        // This ensures problematic zones are processed last (less likely to block good zones)
                        filteredZones = filteredZones
                            .OrderBy(cz => planningMap.ContainsKey(cz.Id) 
                                ? (int)planningMap[cz.Id].ClearanceRisk 
                                : int.MaxValue)
                            .ThenByDescending(cz => planningMap.ContainsKey(cz.Id) 
                                ? planningMap[cz.Id].RawMepSizeFt 
                                : 0)
                            .ToList();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ✅ PARALLEL PLANNING: Completed in {planningResult.PlanningDurationMs:F1}ms - " +
                                $"Processed: {planningResult.TotalCount}, Skipped: {planningResult.SkippedCount}, " +
                                $"High Risk: {planningResult.HighRiskCount}, Critical: {planningResult.CriticalRiskCount}, " +
                                $"Remaining: {filteredZones.Count}");
                            
                            // ✅ BATCH LOGGING: Log all planning results at once (performance optimization)
                            if (planningResult.Items.Count > 0)
                            {
                                var logLines = planningResult.Items
                                    .Select(dto => $"[{DateTime.Now:HH:mm:ss.fff}] {dto.LogSummary}")
                                    .ToList();
                                SafeFileLogger.SafeAppendText("planning_debug.log", string.Join("\n", logLines) + "\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // ✅ GRACEFUL FALLBACK: If planning fails, continue with sequential processing
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] ⚠️ Parallel planning failed: {ex.Message} - Falling back to sequential processing");
                        }
                        planningMap = null;
                        planningResult = null;
                    }
                }
            }
            else if (!DeploymentConfiguration.EnableParallelPlanning && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[NewSleevePlacer] ⏭️ PARALLEL PLANNING: Disabled via safety flag (DeploymentConfiguration.EnableParallelPlanning=false) - Using sequential processing");
            }
            
            // ✅ TIMEOUT PROTECTION: Check timeout periodically
            foreach (var clashZone in filteredZones)
            {
                // ✅ TIMEOUT PROTECTION: Check if operation has exceeded timeout
                if (OptimizationFlags.UseTimeoutProtection && _crashSafeExecutor != null)
                {
                    if (_crashSafeExecutor.CheckTimeout("Place Individual Sleeves"))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] ⏱️ Timeout detected - stopping placement. Processed {placed + skipped + errors} out of {filteredZones.Count} zones");
                        }
                        break; // Stop processing on timeout
                    }
                }
                
                try
                {
                    // ✅ CRITICAL FIX: Skip if already resolved OR if SleeveInstanceId > 0 (even if IsResolved flag is not set)
                    // This prevents placing sleeves over existing sleeves (especially for dampers)
                    // ✅ CRITICAL: If IsClusterResolved=true, skip individual placement (zone is part of a cluster)
                    // ✅ CRITICAL: If ClusterSleeveInstanceId > 0, skip individual placement (cluster sleeve already exists)
                    bool shouldSkip = clashZone.IsResolved || 
                                     clashZone.IsClusterResolved || 
                                     clashZone.SleeveInstanceId > 0 || 
                                     clashZone.ClusterSleeveInstanceId > 0;
                    
                    if (shouldSkip)
                    {
                        // ✅ DIAGNOSTIC: Log why zone is being skipped (always log, even in deployment mode for debugging)
                        string skipReason = "";
                        if (clashZone.IsResolved) skipReason += "IsResolved=true, ";
                        if (clashZone.IsClusterResolved) skipReason += "IsClusterResolved=true, ";
                        if (clashZone.SleeveInstanceId > 0) skipReason += $"SleeveId={clashZone.SleeveInstanceId}, ";
                        if (clashZone.ClusterSleeveInstanceId > 0) skipReason += $"ClusterId={clashZone.ClusterSleeveInstanceId}, ";
                        skipReason = skipReason.TrimEnd(',', ' ');
                        
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⏭️ SKIP Zone {clashZone.Id}: {skipReason}\n");
                        skipped++;
                        continue;
                    }

                    // ✅ DEDUPLICATION: Check if we've already processed a zone at this location
                    // This prevents placing multiple sleeves at the exact same point (e.g. duplicate clashes or pre-clustered zones)
                    // Key includes: Category, HostID, and Rounded Coordinates (to 1mm approx)
                    string locationKey = $"{clashZone.MepElementCategory}_{clashZone.StructuralElementId}_" +
                                         $"{clashZone.IntersectionPointX:F4}_{clashZone.IntersectionPointY:F4}_{clashZone.IntersectionPointZ:F4}";
                                         
                    if (placedLocationKeys.Contains(locationKey))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⏭️ SKIP DUPLICATE: Zone {clashZone.Id} at same location as previous zone ({locationKey})\n");
                        }
                        skipped++;
                        continue;
                    }

                    bool isSleevePlaced = false;
                    FamilyInstance placedSleeve = null;

                    // SMART REPLAY LOGIC (Path 1)
                    if (_isReplayPath)
                    {
                        // Check if sleeve exists in Revit
                        bool sleeveExists = false;
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            var element = _doc.GetElement(new ElementId(clashZone.SleeveInstanceId));
                            if (element != null && element is FamilyInstance)
                            {
                                sleeveExists = true;
                            }
                        }

                        if (sleeveExists)
                        {
                            // Sleeve exists and we are in replay path -> Assume it's correct and skip (or update if needed)
                            // For this refactor, we'll skip placement but mark as processed
                            skipped++;
                            processedZoneGuids.Add(clashZone.Id);
                            continue;
                        }
                        else
                        {
                            // Sleeve is missing (deleted) but we have saved data
                            // Check if we can use saved data ("Smart Replay")
                            if (CanUseSavedData(clashZone))
                            {
                                placedSleeve = PlaceSleeveFromSavedData(clashZone);
                                if (placedSleeve != null)
                                {
                                    isSleevePlaced = true;
                                    placed++;
                                }
                            }
                        }
                    }

                    // If not placed via Smart Replay, proceed with Normal Placement
                    if (!isSleevePlaced)
                    {
                        // ✅ DIAGNOSTIC: Log before attempting normal placement
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 🔨 ATTEMPTING PLACEMENT: Zone {clashZone.Id}, IsReplayPath={_isReplayPath}, SleeveId={clashZone.SleeveInstanceId}\n");
                        
                        placedSleeve = PlaceSleeveNormal(clashZone);
                        if (placedSleeve != null)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ PLACED: Zone {clashZone.Id}, SleeveId={placedSleeve.Id.IntegerValue}\n");
                            placed++;
                        }
                        else
                        {
                            // ✅ DIAGNOSTIC: Log why placement failed
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ PLACEMENT FAILED: Zone {clashZone.Id}, PlaceSleeveNormal returned null\n");
                            skipped++; // Failed to place for some reason (e.g. invalid dimensions)
                        }
                    }

                    if (placedSleeve != null)
                    {
                        // ✅ SAFE ELEMENT VALIDATION: Validate element by ID (avoids document mismatch bug)
                        // ⚠️ CRITICAL: Do NOT compare documents by reference (causes false positives)
                        // Use element ID validation instead - if doc.GetElement() succeeds, element is in correct document
                        if (OptimizationFlags.UseSafeElementValidation)
                        {
                            try
                            {
                                // Validate element is still valid and accessible
                                var validationElement = _doc.GetElement(placedSleeve.Id);
                                if (validationElement == null || !validationElement.IsValidObject)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[NewSleevePlacer] ⚠️ Element {placedSleeve.Id.IntegerValue} is invalid after placement - skipping");
                                    }
                                    errors++;
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error($"[NewSleevePlacer] Error validating element {placedSleeve.Id.IntegerValue}: {ex.Message}");
                                }
                                errors++;
                                continue;
                            }
                        }
                        
                        // ✅ CRITICAL: Update ClashZone with new Sleeve ID and flag
                        // This ensures in-memory object is updated immediately (before batch flag update)
                        clashZone.SleeveInstanceId = placedSleeve.Id.IntegerValue;
                        clashZone.IsResolved = true;
                        
                        // ✅ DIAGNOSTIC: Log flag update for debugging
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ SET FLAG: Zone {clashZone.Id}: IsResolved=true, SleeveId={clashZone.SleeveInstanceId}\n");
                        }
                        
                        // ✅ CRITICAL FIX: Immediately update SleeveInstanceId in database
                        // This ensures the database is updated even if batch persistence is skipped or fails
                        // Required for cleanup service to identify individual sleeves correctly
                        try
                        {
                            UpdateSleeveInstanceIdImmediately(clashZone.Id, placedSleeve.Id.IntegerValue);
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[NewSleevePlacer] Failed to update SleeveInstanceId immediately for zone {clashZone.Id}: {ex.Message}");
                            }
                            // Continue - batch persistence will try again later
                        }
                        
                        // ✅ CRITICAL FIX: Store dimensions for bounding box calculation when batching is enabled
                        // Get dimensions from zone (already set in PlaceSleeveNormal or PlaceSleeveFromSavedData)
                        double storedWidth = clashZone.SleeveWidth > 0 ? clashZone.SleeveWidth : 0;
                        double storedHeight = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : 0;
                        
                        // ✅ CRITICAL FIX: Calculate depth from wall/structural thickness (same logic as SetSleeveParameters)
                        double storedDepth = 0.0;
                        if (clashZone.SleeveDiameter > 0)
                        {
                            // For circular, depth = diameter
                            storedDepth = clashZone.SleeveDiameter;
                        }
                        else
                        {
                            // For rectangular, calculate depth from structural element thickness
                            bool isWallHost = clashZone.StructuralElementType == "Wall" || clashZone.StructuralElementType == "Walls";
                            bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                            
                            if (isWallHost)
                            {
                                storedDepth = clashZone.WallThickness > 0 ? clashZone.WallThickness : clashZone.StructuralElementThickness;
                            }
                            else if (isFramingHost)
                            {
                                storedDepth = clashZone.FramingThickness > 0 ? clashZone.FramingThickness : clashZone.StructuralElementThickness;
                            }
                            else
                            {
                                storedDepth = clashZone.StructuralElementThickness;
                            }
                            
                            // ✅ ROBUST: No fallback - depth MUST be valid
                            if (storedDepth <= 0)
                            {
                                string errorMsg = $"[DEPTH-ERROR] Zone {clashZone.Id} (MEP={clashZone.MepElementId}, StructuralElement={clashZone.StructuralElementId}): " +
                                    $"Invalid depth - WallThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.WallThickness):F1}mm, " +
                                    $"FramingThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.FramingThickness):F1}mm, " +
                                    $"StructuralElementThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.StructuralElementThickness):F1}mm. " +
                                    $"Database corruption detected - run ClashZone refresh to fix.";
                                
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error(errorMsg);
                                
                                throw new InvalidOperationException(errorMsg);
                            }
                        }
                        
                        // If dimensions not set, try to get from deferred parameters
                        if (storedWidth <= 0 || storedHeight <= 0)
                        {
                            storedWidth = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Width", storedWidth);
                            storedHeight = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Height", storedHeight);
                        }
                        
                        // ✅ CRITICAL: Also check deferred parameters for Depth/Wall Width (set by SetSleeveParameters)
                        double depthFromParams = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Depth", 0.0);
                        if (depthFromParams <= 0.0)
                        {
                            depthFromParams = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Wall Width", 0.0);
                        }
                        if (depthFromParams > 0.0)
                        {
                            storedDepth = depthFromParams;
                        }
                        
                        placedSleeveData.Add((placedSleeve, clashZone, storedWidth, storedHeight, storedDepth));
                        processedZoneGuids.Add(clashZone.Id);
                        
                        // ✅ DEDUPLICATION: Add location to tracker
                        placedLocationKeys.Add(locationKey);
                        
                        // ✅ NOTE: Database persistence is now handled in batch after bounding boxes are calculated
                        // This ensures all data (instance ID, placement, bounding boxes, corners, snapshots) is saved together
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[NewSleevePlacer] Error placing sleeve for zone {clashZone.Id}: {ex.Message}");
                }
            }

            // ✅ PERFORMANCE OPTIMIZATION: Batch regeneration after ALL sleeves placed
            if (placedSleeveData.Count > 0)
            {
                var regenTimer = System.Diagnostics.Stopwatch.StartNew();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-REGEN] Regenerating document for {placedSleeveData.Count} sleeves...");
                }
                _doc.Regenerate(); // ✅ Single regeneration for all sleeves
                
                // ✅ CRITICAL FIX: Clear family symbol cache after regeneration
                // Regeneration invalidates ALL element references, including cached FamilySymbol objects
                if (OptimizationFlags.UseFamilySymbolCache)
                {
                    _familySymbolCache.Clear();
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[NewSleevePlacer] [FAMILY-CACHE] Cleared cache after document regeneration (all symbols invalidated)");
                }
                regenTimer.Stop();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-REGEN] ✅ Regenerated {placedSleeveData.Count} sleeves in {regenTimer.ElapsedMilliseconds}ms");
                }
                
                // ✅ CRITICAL FIX: Calculate bounding boxes from deferred parameters when batching is enabled
                // When UseBatchedParameterWrites=true, sleeve.get_BoundingBox() returns STALE values because parameters
                // haven't been flushed yet. We must calculate bounding boxes from stored dimensions instead.
                var bboxTimer = System.Diagnostics.Stopwatch.StartNew();
                int bboxCount = 0;
                foreach (var (sleeve, zone, storedWidth, storedHeight, storedDepth) in placedSleeveData)
                {
                    try
                    {
                        // ✅ Validate sleeve still exists (may have been deleted by clustering)
                        if (!sleeve.IsValidObject) continue;

                        // ✅ CRITICAL FIX: When batch writing is enabled, calculate bounding box from stored dimensions
                        // instead of reading from Revit (which returns stale values)
                        BoundingBoxXYZ actualBbox = null;
                        if (OptimizationFlags.UseBatchedParameterWrites && storedWidth > 0 && storedHeight > 0)
                        {
                            // ✅ BATCHING ENABLED: Calculate bounding box from stored dimensions
                            actualBbox = CalculateBoundingBoxFromDimensions(sleeve, storedWidth, storedHeight, storedDepth, zone);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] [BATCH-BBOX] Zone {zone.Id}: Calculated bbox from DEFERRED params - W={storedWidth * 304.8:F1}mm, H={storedHeight * 304.8:F1}mm, D={storedDepth * 304.8:F1}mm\n");
                            }
                        }
                        
                        // ✅ FALLBACK: If batching disabled or calculation failed, read from Revit
                        if (actualBbox == null)
                        {
                            actualBbox = sleeve.get_BoundingBox(null);
                        }
                        
                        if (actualBbox != null)
                        {
                            // ✅ SRP COMPLIANCE: Delegate RCS transformation to specialized service
                            _rcsBoundingBoxService.ProcessBoundingBox(zone, actualBbox);

                            // Update placement point from bounding box center
                            zone.SleevePlacementPoint = new XYZ(
                                (actualBbox.Min.X + actualBbox.Max.X) / 2,
                                (actualBbox.Min.Y + actualBbox.Max.Y) / 2,
                                (actualBbox.Min.Z + actualBbox.Max.Z) / 2
                            );

                            zone.SleevePlacementPointX = zone.SleevePlacementPoint.X;
                            zone.SleevePlacementPointY = zone.SleevePlacementPoint.Y;
                            zone.SleevePlacementPointZ = zone.SleevePlacementPoint.Z;

                            bboxCount++;
                        }
                    }
                    catch (Exception bboxEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] [BATCH-BBOX] Error retrieving bbox for sleeve {sleeve?.Id?.IntegerValue ?? -1}: {bboxEx.Message}");
                        }
                    }
                }
                bboxTimer.Stop();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-BBOX] ✅ Retrieved {bboxCount} bounding boxes in {bboxTimer.ElapsedMilliseconds}ms (avg: {bboxTimer.ElapsedMilliseconds / Math.Max(1, bboxCount):F1}ms per sleeve)");
                }
                
                // ✅ SRP COMPLIANCE: Persist all sleeve data to database in batch (instance ID, placement, bounding boxes, corners, snapshots)
                // ✅ CRITICAL: Batch save corners to database AFTER regeneration
                // - Corners are pre-calculated in parallel during PersistSleeveData
                // - All 4 corners (Corner1X/Y/Z through Corner4X/Y/Z) are saved to database
                // - Cluster calculation will use these saved corners directly (NO recalculation needed)
                // - This ensures cluster sizing uses accurate corner coordinates without expensive recalculation
                if (placedSleeveData.Count > 0)
                {
                    try
                    {
                        int persistedCount = _persistenceService.PersistSleeveData(placedSleeveData, _filterName);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ✅ Persisted {persistedCount} sleeves to database (instance ID, placement, bounding boxes, corners, snapshots)");
                        }
                    }
                    catch (Exception persistEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] Failed to persist sleeve data: {persistEx.Message}");
                        }
                    }
                }
            }
            
            // ✅ CRITICAL FIX: Batch update flags at the end (only if FlagManager is provided)
            // ✅ FLAG MANAGEMENT: Update IsResolvedFlag and IsClusterResolvedFlag in database after placement
            // This ensures flags are set correctly for dampers and all other MEP elements
            // ✅ FIX: Only update flags for zones that were actually placed (have SleeveInstanceId > 0)
            // Note: Cluster placement handles its own flag updates separately
            if (placedSleeveData.Count > 0 && _flagManager != null)
            {
                try
                {
                    // ✅ CRITICAL: Filter to only zones with SleeveInstanceId > 0 (actually placed)
                    // This ensures we don't try to update flags for zones that failed placement
                    var batchUpdates = placedSleeveData
                        .Where(x => x.zone != null && x.zone.SleeveInstanceId > 0)
                        .Select(x => (x.zone, x.zone.SleeveInstanceId))
                        .ToList();

                    if (batchUpdates.Count > 0)
                    {
                        _flagManager.BatchUpdateFlagsForPlacement(
                            batchUpdates,
                            isCluster: false,
                            placedSleeveData[0].zone.MepElementCategory,
                            _filterName
                        );
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ✅ Updated flags for {batchUpdates.Count} placed sleeves (Category: {placedSleeveData[0].zone.MepElementCategory})");
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ BATCH FLAG UPDATE: Updated IsResolvedFlag=1 for {batchUpdates.Count} sleeves\n");
                        }
                    }
                    else if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // ✅ DIAGNOSTIC: Log when no flags are updated (all zones have SleeveInstanceId = 0)
                        int zonesWithNoSleeveId = placedSleeveData.Count(x => x.zone == null || x.zone.SleeveInstanceId <= 0);
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⚠️ NO FLAG UPDATE: {zonesWithNoSleeveId} zones have SleeveInstanceId <= 0 (flags won't be updated)\n");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] ❌ Error updating flags: {ex.Message}");
                        DebugLogger.Error($"[NewSleevePlacer] Stack trace: {ex.StackTrace}");
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ FLAG UPDATE ERROR: {ex.Message}\n");
                    }
                }
            }
            else if (placedSleeveData.Count > 0 && _flagManager == null)
            {
                // ✅ DIAGNOSTIC: Log warning if FlagManager is not provided (flags won't be updated)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[NewSleevePlacer] ⚠️ FlagManager is null - flags will NOT be updated in database for {placedSleeveData.Count} placed sleeves. IsResolvedFlag will remain 0.");
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⚠️ FLAG MANAGER NULL: Flags will NOT be updated for {placedSleeveData.Count} sleeves\n");
                }
            }
            
            // ✅ PARAMETER BATCHING: Flush deferred parameters AFTER bounding box calculation
            // Bounding boxes must be calculated BEFORE flushing, so they use deferred parameter values
            if (OptimizationFlags.UseBatchedParameterWrites)
            {
                try
                {
                    int flushedCount = _parameterService.FlushDeferredParameters();
                    if (!DeploymentConfiguration.DeploymentMode && flushedCount > 0)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ✅ Flushed {flushedCount} deferred parameters via SleeveParameterService");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Error flushing deferred parameters: {ex.Message}");
                    }
                }
            }
            
            // Reset ReadyForPlacement flags
            if (processedZoneGuids.Count > 0)
            {
                ResetReadyForPlacementFlags(processedZoneGuids);
            }

            return (placed, skipped, errors);
        }

        private bool CanUseSavedData(ClashZone zone)
        {
            // Check if we have valid saved dimensions and placement point
            // Also verify clearance hasn't changed (simplified check for now)
            return zone.SleeveWidth > 0 && 
                   zone.SleeveHeight > 0 && 
                   zone.SleevePlacementPointX != 0;
        }

        private FamilyInstance PlaceSleeveFromSavedData(ClashZone zone)
        {
            // Use saved dimensions directly
            double width = zone.SleeveWidth;
            double height = zone.SleeveHeight;
            double diameter = zone.SleeveDiameter;
            
            // Determine shape based on dimensions
            bool isCircular = diameter > 0;
            
            // Select Family
            string familyName = GetSleeveFamilyName(zone, isCircular);
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null) return null;

            // Place Instance
            XYZ placementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
            
            // ✅ SRP: Use rotation service to determine correct rotation for host type
            double rotation = _rotationService.DetermineRotation(zone); 

            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance != null)
            {
                // ✅ CRITICAL FIX: Round dimensions ONCE here (for all categories including dampers)
                // This ensures both the Revit parameters AND the saved zone dimensions are rounded (consistent with cluster sleeves)
                // Rounding is applied to ALL categories (dampers, pipes, ducts, cable trays, etc.)
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                double roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(diameter);
                
                // ✅ CRITICAL FIX: Pass ROUNDED dimensions to SetSleeveParameters (no rounding inside SetSleeveParameters to prevent double rounding)
                // This ensures Revit parameters are set with rounded values
                _parameterService.SetSleeveParameters(instance, roundedWidth, roundedHeight, roundedDiameter, isCircular, zone);
                
                // ✅ CRITICAL FIX: Update zone with ROUNDED dimensions for bounding box calculation and database saving
                // This ensures dimensions are available when batching is enabled AND replay mode uses rounded dimensions
                zone.SleeveWidth = roundedWidth;
                zone.SleeveHeight = roundedHeight;
                zone.SleeveDiameter = roundedDiameter;
                zone.SleevePlacementPoint = placementPoint;
                zone.SleevePlacementPointX = placementPoint.X;
                zone.SleevePlacementPointY = placementPoint.Y;
                zone.SleevePlacementPointZ = placementPoint.Z;
            }
            
            return instance;
        }

        private FamilyInstance PlaceSleeveNormal(ClashZone zone, SleevePlacementPlanningDto? planningDto = null)
        {
            // ✅ DIAGNOSTIC: Log entry into PlaceSleeveNormal
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 🔨 PlaceSleeveNormal START: Zone {zone.Id}, HasPlanningDto={planningDto != null}\n");
            
            // ✅ PARALLEL PLANNING: Pass planningDto to dimension calculation
            // If planningDto is provided, CalculateSleeveDimensions will use pre-computed values (faster)
            // This includes dampers - parallel planning handles all categories
            var (width, height, diameter, isCircular) = CalculateSleeveDimensions(zone, planningDto);
            
            // ✅ DIAGNOSTIC: Log calculated dimensions (in both feet and mm for readability)
            double widthMm = width * 304.8;
            double heightMm = height * 304.8;
            double diameterMm = diameter * 304.8;
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 📐 FINAL DIMENSIONS: Zone {zone.Id}, " +
                $"W={width:F6}ft ({widthMm:F1}mm), H={height:F6}ft ({heightMm:F1}mm), D={diameter:F6}ft ({diameterMm:F1}mm), Circular={isCircular}\n");
            
            if (width <= 0 && height <= 0 && diameter <= 0)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ INVALID DIMENSIONS: Zone {zone.Id}, all dimensions <= 0\n");
                return null; // Invalid dimensions
            }

            // Select Family
            string familyName = GetSleeveFamilyName(zone, isCircular);
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 🏠 FAMILY: Zone {zone.Id}, FamilyName='{familyName}'\n");
            
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ SYMBOL NULL: Zone {zone.Id}, FamilyName='{familyName}' - LoadFamilySymbol returned null\n");
                return null;
            }

            // Determine Placement Point (Intersection Point)
            XYZ placementPoint = zone.IntersectionPoint ?? new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
            
            // ✅ SRP COMPLIANCE: Delegate placement point adjustment to dedicated service
            // PlacementPointAdjustmentService handles non-dampers, DamperPlacementPointService handles dampers
            // This keeps NewSleevePlacerService focused on orchestration, not geometric calculations
            
            // ✅ EXTRACT CLEARANCE VALUES: Store individual clearance values in zone for later parameter setting
            // This is done BEFORE placement point adjustment so values are available for parameter setting
            XYZ damperOffsetVector = XYZ.Zero;
            if (_damperPlacementAdjustments.ContainsKey(zone.Id))
            {
                var (finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom, offsetVector) = 
                    _damperPlacementAdjustments[zone.Id];
                
                // ✅ SOLID ISP: Store individual clearance values in zone for later parameter setting
                // This allows SetSleeveParameters to set clearance parameters without strategy dependency
                zone.ClearanceLeft = clearanceLeft;
                zone.ClearanceRight = clearanceRight;
                zone.ClearanceTop = clearanceTop;
                zone.ClearanceBottom = clearanceBottom;
                
                // ✅ CRITICAL FIX: Store offset vector for later application
                damperOffsetVector = offsetVector;
            }
            
            // ✅ DELEGATE TO SERVICE: PlacementPointAdjustmentService delegates to DamperPlacementPointService for dampers
            // This maintains SRP - one class handles damper placement, another handles other MEP elements
            placementPoint = _placementPointAdjustmentService.AdjustPlacementPoint(zone, placementPoint, null);
            
            // ✅ CRITICAL FIX: Apply connector-side offset for dampers with MEP connector
            // The offset shifts the sleeve toward the connector side to achieve:
            // - 100mm clearance on connector side (MEP clearance)
            // - 50mm clearance on other side (Other clearance)
            // Without offset, sleeve would be centered, giving 75mm/75mm (average)
            // The offset is calculated as (mepClearance - otherClearance) / 2 = (100mm - 50mm) / 2 = 25mm
            if (string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase) &&
                damperOffsetVector.GetLength() > 0.0001) // Only apply if offset is significant (> 0.1mm)
            {
                placementPoint = placementPoint + damperOffsetVector;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ APPLIED DAMPER OFFSET: Zone {zone.Id}, " +
                        $"Offset=({damperOffsetVector.X*304.8:F1}, {damperOffsetVector.Y*304.8:F1}, {damperOffsetVector.Z*304.8:F1})mm, " +
                        $"AfterOffset=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3})\n");
                }
            }
                    
                    // ✅ DIAGNOSTIC: Log placement point
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 📍 PLACEMENT POINT: Zone {zone.Id}, Point=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3})\n");
            
            // ✅ SRP: Use rotation service to determine correct rotation for host type
            double rotation = _rotationService.DetermineRotation(zone);

            // Place Instance
            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance == null)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ INSTANCE NULL: Zone {zone.Id}, PlaceSleeveInstance returned null\n");
                return null;
            }
            
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ INSTANCE CREATED: Zone {zone.Id}, InstanceId={instance.Id.IntegerValue}\n");
            
            if (instance != null)
            {
                // ✅ CRITICAL FIX: Round dimensions ONCE here (for all categories including dampers)
                // This ensures both the Revit parameters AND the saved zone dimensions are rounded (consistent with cluster sleeves)
                // Rounding is applied to ALL categories (dampers, pipes, ducts, cable trays, etc.)
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                double roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(diameter);
                
                // ✅ CRITICAL FIX: Pass ROUNDED dimensions to SetSleeveParameters (it will round again, but rounding already-rounded values is idempotent)
                // This ensures Revit parameters are set with rounded values
                _parameterService.SetSleeveParameters(instance, roundedWidth, roundedHeight, roundedDiameter, isCircular, zone);
                
                // ✅ CRITICAL FIX: Update zone with ROUNDED dimensions for saving to database
                // This ensures replay mode uses rounded dimensions (consistent with cluster sleeves)
                zone.SleeveWidth = roundedWidth;
                zone.SleeveHeight = roundedHeight;
                zone.SleeveDiameter = roundedDiameter;
                zone.SleevePlacementPoint = placementPoint;
                zone.SleevePlacementPointX = placementPoint.X;
                zone.SleevePlacementPointY = placementPoint.Y;
                zone.SleevePlacementPointZ = placementPoint.Z;
            }
            
            return instance;
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Calculate sleeve dimensions with proper clearance handling.
        /// When batch parameter writing is enabled, this ensures:
        /// 1. Cable trays use GetCableTrayPlacementAdjustment which correctly reads clearances from conditions.ClearanceSettings (database)
        /// 2. Other categories use GetClearance which may read from strategy or fallback
        /// 3. All parameter reads during placement use GetParameterValueWithBatchingSupport to prevent stale reads
        /// 
        /// ✅ PARALLEL PLANNING: If planningDto is provided (from parallel planning phase), use pre-computed dimensions.
        /// This includes dampers - parallel planning handles all categories including Duct Accessories.
        /// </summary>
        private (double width, double height, double diameter, bool isCircular) CalculateSleeveDimensions(ClashZone zone, SleevePlacementPlanningDto? planningDto = null)
        {
            // ✅ PERFORMANCE MONITORING: Track dimension calculation
            using (var tracker = _performanceMonitor?.TrackOperation("Calculate Sleeve Dimensions"))
            {
                // ✅ PARALLEL PLANNING: Use pre-computed dimensions if available (from parallel planning phase)
                // This includes dampers - ParallelSleevePlacementPlanner handles Duct Accessories category
                // Safety: Only use planning data if EnableParallelPlanning flag is enabled and planningDto is valid
                if (DeploymentConfiguration.EnableParallelPlanning && planningDto != null && !planningDto.ShouldSkip)
                {
                    // ✅ USE PRE-COMPUTED DIMENSIONS: From parallel planning (includes dampers)
                    // Note: For dampers with asymmetric clearance, we still need strategy for offset calculation
                    // But dimensions can come from planning phase
                    bool isDamper = string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                    
                    if (!isDamper)
                    {
                        // ✅ NON-DAMPER: Use pre-computed dimensions directly (faster, already calculated in parallel)
                        tracker?.SetItemCount(1);
                        return (planningDto.TargetWidthFt, planningDto.TargetHeightFt, Math.Max(planningDto.TargetWidthFt, planningDto.TargetHeightFt), false);
                    }
                    else
                    {
                        // ✅ DAMPER: Use pre-computed dimensions but still need strategy for offset calculation
                        // Planning phase calculated dimensions, but we need strategy for asymmetric clearance offset
                        // Fall through to strategy-based calculation below
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] 📋 DAMPER: Using pre-computed dimensions from planning: " +
                                $"W={planningDto.TargetWidthFt * 304.8:F1}mm, H={planningDto.TargetHeightFt * 304.8:F1}mm, " +
                                $"but still need strategy for offset calculation");
                        }
                    }
                }
                
                // ✅ DAMPER PLACEMENT STRATEGY: Handle dampers with connector-aware asymmetric clearance
                // Dampers with MEP connectors require special offset calculation to achieve 100mm on connector side, 50mm on other
                if (_strategy is DamperPlacementStrategy damperStrategy)
                {
                    // ✅ COMPREHENSIVE LOGGING: Log raw MEP dimensions and clearance settings BEFORE strategy call
                    double rawWidthMm = zone.MepElementWidth * 304.8;
                    double rawHeightMm = zone.MepElementHeight * 304.8;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-PRE-STRATEGY] Zone {zone.Id}, " +
                            $"RawMEP W={rawWidthMm:F1}mm, H={rawHeightMm:F1}mm, " +
                            $"DB Clearances: MEP={(zone.IsInsulated ? _conditions?.ClearanceSettings?.DuctAccessoryMepInsulated : _conditions?.ClearanceSettings?.DuctAccessoryMepNormal) ?? 0}mm, " +
                            $"Other={(zone.IsInsulated ? _conditions?.ClearanceSettings?.DuctAccessoryOtherInsulated : _conditions?.ClearanceSettings?.DuctAccessoryOtherNormal) ?? 0}mm, " +
                            $"IsInsulated={zone.IsInsulated}, HasConditions={_conditions != null}, HasClearanceSettings={_conditions?.ClearanceSettings != null}\n");
                    }
                    
                    // ✅ OOP METHOD: Get damper placement adjustment including offset and individual clearances (SOLID DIP)
                    // This method handles:
                    // 1. Connector detection (which side has MEP connector)
                    // 2. Asymmetric clearance mapping (100mm MEP side, 50mm other side)
                    // 3. Offset calculation (25mm toward connector to redistribute from 75/75 to 100/50)
                    // 4. Wall-specific handling (Z-axis swap for vertical connectors, axis-based offset)
                    var damperAdj = damperStrategy.GetDamperPlacementAdjustment(zone, _conditions);
                    
                    // ✅ COMPREHENSIVE LOGGING: Log final dimensions AFTER strategy call
                    double finalWidthMm = damperAdj.finalWidth * 304.8;
                    double finalHeightMm = damperAdj.finalHeight * 304.8;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-POST-STRATEGY] Zone {zone.Id}, " +
                            $"Final W={damperAdj.finalWidth:F6}ft ({finalWidthMm:F1}mm), H={damperAdj.finalHeight:F6}ft ({finalHeightMm:F1}mm), " +
                            $"ClearanceApplied W={finalWidthMm - rawWidthMm:F1}mm, H={finalHeightMm - rawHeightMm:F1}mm\n");
                    }
                    
                    // ✅ SOLID ISP: Extract individual clearance values from damper strategy result
                    // These will be stored in ClashZone and set as sleeve parameters later
                    // This allows placement service to use zone data without direct strategy dependency
                    double clearanceLeft = damperAdj.finalWidth > 0 ? damperAdj.finalWidth : zone.ClearanceLeft;
                    double clearanceRight = damperAdj.finalHeight > 0 ? damperAdj.finalHeight : zone.ClearanceRight;
                    double clearanceTop = damperAdj.finalWidth > 0 ? damperAdj.finalHeight : zone.ClearanceTop;
                    double clearanceBottom = damperAdj.finalHeight > 0 ? damperAdj.finalWidth : zone.ClearanceBottom;
                    
                    // ✅ CRITICAL: Store damper clearance values for later parameter setting
                    // Key: zone.Id (Guid) for matching, Value: final dimensions + individual clearances + offsetVector
                    // ✅ CRITICAL FIX: offsetVector is now stored to apply connector-side offset during placement
                    // The offset shifts sleeve toward connector side to achieve 100mm/50mm clearance (not 75mm/75mm)
                    _damperPlacementAdjustments[zone.Id] = (
                        damperAdj.finalWidth,
                        damperAdj.finalHeight,
                        clearanceLeft,
                        clearanceRight,
                        clearanceTop,
                        clearanceBottom,
                        damperAdj.offsetVector // ✅ Store offset for later application
                    );
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ✅ DAMPER STRATEGY: Zone {zone.Id}, Final W={damperAdj.finalWidth:F6}ft, H={damperAdj.finalHeight:F6}ft (placement point handled by DamperPlacementPointService)");
                    }
                    
                    tracker?.SetItemCount(1);
                    return (damperAdj.finalWidth, damperAdj.finalHeight, 0, false); // Dampers are always rectangular (false = not circular)
                }
                
                // ✅ CRITICAL FIX: Cable trays need special handling to read clearances directly from database
                // When batch parameter writing is enabled, we MUST read clearances from _conditions.ClearanceSettings (database)
                // NOT from any cached/stale values or parameter reads
                if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                {
                    // ✅ VALIDATE: Ensure _conditions has ClearanceSettings populated from database
                    if (_conditions?.ClearanceSettings == null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[NewSleevePlacer] ⚠️ _conditions.ClearanceSettings is NULL for cable tray zone {zone.Id} - using fallback clearances");
                    }
                    else
                    {
                        // ✅ DIAGNOSTIC: Log that we're using database clearances (critical for debugging batching issues)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ✅ Using DATABASE clearances: CableTrayTop={_conditions.ClearanceSettings.CableTrayTop}mm, CableTrayOther={_conditions.ClearanceSettings.CableTrayOther}mm");
                        }
                    }
                    
                    // ✅ CABLE TRAY: Use strategy's GetCableTrayPlacementAdjustment (reads from conditions.ClearanceSettings - database)
                    // This method prioritizes _conditions.ClearanceSettings (database) over UI settings
                    var ctRawWidth = zone.MepElementWidth;
                    var ctRawHeight = zone.MepElementHeight;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[NewSleevePlacer] CABLE TRAY STRATEGY: Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(ctRawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(ctRawHeight):F1}mm");
                    
                    // ✅ CRITICAL: Pass _conditions (with database ClearanceSettings) to strategy
                    // Strategy will read CableTrayTop and CableTrayOther from _conditions.ClearanceSettings (database)
                    var adj = cableTrayStrategy.GetCableTrayPlacementAdjustment(zone, _conditions, _clearanceSettings);
                    double ctFinalWidth = adj.finalWidth;
                    double ctFinalHeight = adj.finalHeight;
                    double ctFinalDiameter = ctFinalWidth; // Not used for cable trays (rectangular only)
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[NewSleevePlacer] CABLE TRAY STRATEGY: Final={RevitUnitConversionService.Instance.FromInternalMillimeters(ctFinalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(ctFinalHeight):F1}mm");
                    
                    tracker?.SetItemCount(1);
                    return (ctFinalWidth, ctFinalHeight, ctFinalDiameter, false); // Cable trays are always rectangular
                }
                
                // ✅ OOP METHOD: Use strategy to calculate clearance for other categories
                double rawWidth = zone.MepElementWidth;
                double rawHeight = zone.MepElementHeight;
                double rawDiameter = zone.MepElementOuterDiameter > 0 ? zone.MepElementOuterDiameter : 0;
                
                // ✅ COMPREHENSIVE LOGGING: Log raw dimensions BEFORE clearance calculation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-PRE-CLEARANCE] Zone {zone.Id}, Category={zone.MepElementCategory}, " +
                        $"Raw W={rawWidth:F6}ft ({rawWidth * 304.8:F1}mm), H={rawHeight:F6}ft ({rawHeight * 304.8:F1}mm), " +
                        $"D={rawDiameter:F6}ft ({rawDiameter * 304.8:F1}mm), IsInsulated={zone.IsInsulated}\n");
                }
                
                double clearance = _clearanceService.GetClearance(zone, _conditions, _clearanceSettings, _strategy);
                double clearanceMm = clearance * 304.8; // Convert feet (internal units) to mm (1ft = 304.8mm)
                
                // ✅ COMPREHENSIVE LOGGING: Log clearance value AFTER retrieval
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-CLEARANCE] Zone {zone.Id}, " +
                        $"Clearance={clearance:F6}ft ({clearanceMm:F1}mm)\n");
                }
            
                // ✅ OOP METHOD: Use insulation-aware sizing service for consistent calculation (SOLID principles)
                // Formula: RawSize + (2 × InsulationThickness) + (2 × Clearance)
                (double finalWidth, double finalHeight, double finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZone(
                    rawWidth, rawHeight, rawDiameter, zone, clearance);
                
                // ✅ COMPREHENSIVE LOGGING: Log final dimensions AFTER sizing calculation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-POST-SIZING] Zone {zone.Id}, " +
                        $"Final W={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), H={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm), " +
                        $"D={finalDiameter:F6}ft ({finalDiameter * 304.8:F1}mm), " +
                        $"ClearanceApplied W={finalWidth * 304.8 - rawWidth * 304.8:F1}mm, H={finalHeight * 304.8 - rawHeight * 304.8:F1}mm\n");
                }
            
                // ✅ GLOBAL SETTINGS: Determine opening type (circular vs rectangular) using global configuration rules
                // This matches the legacy UniversalSleevePlacerService.SelectUniversalFamily() logic
                bool isCircular = DetermineOpeningType(zone, rawDiameter, finalDiameter);

                tracker?.SetItemCount(1);
                return (finalWidth, finalHeight, finalDiameter, isCircular);
            }
        }

        /// <summary>
        /// ✅ GLOBAL SETTINGS: Determine opening type (circular vs rectangular) based on global configuration rules
        /// Matches legacy UniversalSleevePlacerService.SelectUniversalFamily() logic
        /// </summary>
        private bool DetermineOpeningType(ClashZone zone, double rawDiameter, double finalDiameter)
        {
            try
            {
                // ✅ Pipes: Use ConfigurationResolutionService to check global rules
                // Rules checked:
                // 1. RoundOpeningsBecomeRectangularIfDiameterGreaterThan (default 200mm) - if diameter > threshold, make rectangular
                // 2. Structural framing rule - pipes on structural framing are always circular
                if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // Create MepElementSize from ClashZone for configuration resolution
                    var mepSize = new MepElementSize
                    {
                        Width = zone.MepElementWidth,
                        Height = zone.MepElementHeight,
                        Diameter = rawDiameter > 0 ? rawDiameter : zone.MepElementOuterDiameter,
                        Shape = rawDiameter > 0 ? "Round" : "Rectangular",
                        IsInsulated = zone.MepElementSizeData?.IsInsulated ?? false,
                        InsulationThickness = zone.MepElementSizeData?.InsulationThickness ?? 0.0
                    };
                    
                    // Get UI preference from CONDITIONS XML (default to Circular)
                    var pipeType = _conditions?.OpeningTypePreferences?.Pipes ?? "Circular";
                    var hostType = zone.StructuralElementType ?? "Unknown";
                    
                    // Use PipePlacementStrategy to resolve opening type with global rules
                    if (_strategy is PipePlacementStrategy pipeStrategy)
                    {
                        var resolvedType = pipeStrategy.GetResolvedOpeningType(mepSize, pipeType, hostType);
                        bool isCircularResult = string.Equals(resolvedType, "Circular", StringComparison.OrdinalIgnoreCase);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] PIPE opening type resolved: Host={hostType}, UI='{pipeType}' → Global Rule='{resolvedType}' → isCircular={isCircularResult}");
                        }
                        
                        return isCircularResult;
                    }
                    else
                    {
                        // Fallback: Use raw diameter check if strategy not available
                        bool isCircularResult = rawDiameter > 0;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] PIPE opening type (fallback): rawDiameter={rawDiameter:F6}ft → isCircular={isCircularResult}");
                        }
                        return isCircularResult;
                    }
                }
                // ✅ Round Ducts: Check user preference from CONDITIONS XML
                else if (string.Equals(zone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                {
                    // Check if this is a round duct
                    bool isRoundDuct = string.Equals(zone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(zone.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase) ||
                                      (zone.MepElementSizeData != null && 
                                       (string.Equals(zone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(zone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)));
                    
                    if (isRoundDuct)
                    {
                        // Round ducts: use user preference from CONDITIONS XML
                        var roundDuctType = _conditions?.OpeningTypePreferences?.RoundDucts ?? "Circular";
                        bool isCircularResult = string.Equals(roundDuctType, "Circular", StringComparison.OrdinalIgnoreCase);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ROUND DUCT opening type from CONDITIONS XML: '{roundDuctType}' → isCircular={isCircularResult}");
                        }
                        
                        return isCircularResult;
                    }
                    else
                    {
                        // Rectangular ducts: always rectangular opening
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] RECTANGULAR DUCT → isCircular=false");
                        }
                        return false;
                    }
                }
                else
                {
                    // Other categories (cable trays, accessories): always rectangular
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] OTHER CATEGORY ({zone.MepElementCategory}) → isCircular=false");
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[NewSleevePlacer] Error determining opening type for zone {zone?.Id}: {ex.Message}");
                }
                // Fallback: Use raw diameter check
                return rawDiameter > 0;
            }
        }

        // ✅ SRP COMPLIANCE: GetClearance() method removed - now delegated to ClearanceCalculationService
        // This ensures NewSleevePlacerService focuses on placement logic, not clearance calculation

        private string GetSleeveFamilyName(ClashZone zone, bool isCircular)
        {
            // ✅ FIX: Use correct family names matching UniversalSleevePlacerService
            // Determine host type (check both singular and plural forms, plus Structural Framing)
            bool isWallOrFraming = zone.StructuralElementType == "Wall" || 
                                  zone.StructuralElementType == "Walls" ||
                                  zone.StructuralElementType == "Structural Framing" ||
                                  (zone.StructuralElementType != null && zone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase));
            
            // Select family using exact family names (matching UniversalSleevePlacerService.SelectUniversalFamily)
            if (isWallOrFraming)
            {
                return isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
            }
            else // Slab/Floor
            {
                return isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }
        }

        private FamilySymbol LoadFamilySymbol(string familyName)
        {
            // ✅ PERFORMANCE MONITORING: Track family symbol loading
            using (var tracker = _performanceMonitor?.TrackOperation("Load Family Symbol"))
            {
                // ✅ CACHE VALIDATION: Check cache first, but validate before returning
            if (_familySymbolCache.ContainsKey(familyName))
            {
                var cachedSymbol = _familySymbolCache[familyName];
                
                // ✅ CRITICAL SAFETY: Validate symbol before use (prevents stale reference errors)
                if (cachedSymbol != null && cachedSymbol.IsValidObject)
                {
                    try
                    {
                        // Test if we can access IsActive (will throw if symbol is stale)
                        var _ = cachedSymbol.IsActive;
                        return cachedSymbol; // Symbol is valid
                    }
                    catch (InvalidOperationException)
                    {
                        // Symbol is stale - remove from cache
                        _familySymbolCache.Remove(familyName);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] Stale symbol detected for '{familyName}', removed from cache");
                        }
                    }
                }
                else
                {
                    // Symbol is invalid - remove from cache
                    _familySymbolCache.Remove(familyName);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[NewSleevePlacer] Invalid symbol detected for '{familyName}', removed from cache");
                    }
                }
            }
            
            // Use FamilyManager if available
            FamilySymbol foundSymbol = null;
            if (_familyManager != null)
            {
                // Load family using FamilyManager (returns bool, not FamilySymbol)
                bool familyLoaded = _familyManager.LoadFamily(_doc, familyName);
                if (familyLoaded)
                {
                    // ✅ FIX: After loading, find the symbol using Family.Name (not FamilyName property)
                    foundSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                }
            }
            
            // ✅ FIX: Fallback to simple lookup if FamilyManager didn't find it
            // Use Family.Name (not FamilyName property) to match UniversalSleevePlacerService
            if (foundSymbol == null)
            {
                foundSymbol = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
            }
            
            // ✅ CACHE VALIDATION: Only cache if symbol is valid
            if (foundSymbol != null && foundSymbol.IsValidObject)
            {
                try
                {
                    // Test if symbol is active (will throw if stale)
                    var _ = foundSymbol.IsActive;
                    _familySymbolCache[familyName] = foundSymbol;
                }
                catch (InvalidOperationException)
                {
                    // Symbol is stale - don't cache
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[NewSleevePlacer] Stale symbol detected for '{familyName}', skipping cache");
                    }
                }
            }
            
                tracker?.SetItemCount(foundSymbol != null ? 1 : 0);
                return foundSymbol;
            }
        }

        private FamilyInstance PlaceSleeveInstance(FamilySymbol symbol, XYZ point, ClashZone zone, double rotation)
        {
            // ✅ PERFORMANCE MONITORING: Track family symbol placement
            using (var tracker = _performanceMonitor?.TrackOperation("Place Sleeve Instance"))
            {
                // ✅ CRITICAL SAFETY: Validate symbol before accessing IsActive
                if (symbol == null || !symbol.IsValidObject)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Invalid symbol for zone {zone.Id}");
                    }
                    return null;
                }
                
                try
                {
                    if (!symbol.IsActive) symbol.Activate();
                }
                catch (InvalidOperationException ex)
                {
                    // Symbol is stale - remove from cache
                    var familyName = symbol.Family?.Name ?? "Unknown";
                    _familySymbolCache.Remove(familyName);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Stale symbol '{familyName}' for zone {zone.Id}: {ex.Message}");
                    }
                    return null;
                }

            // ✅ CRITICAL FIX: Use absolute placement point (do NOT pass level to NewFamilyInstance)
            // When level is passed, Revit interprets Z coordinate as relative to level elevation
            // But our placement point is ABSOLUTE (already in host document coordinates)
            // Solution: Create instance without level, then set level parameter separately
            
            // Determine level for parameter setting (not for placement)
            Level level = _doc.ActiveView?.GenLevel;
            
            // Try to find level by name if available
            if (level == null && !string.IsNullOrEmpty(zone.MepElementLevelName))
            {
                level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault(l => l.Name.Equals(zone.MepElementLevelName, StringComparison.OrdinalIgnoreCase));
            }
            
            // Final fallback: get first level in document
            if (level == null)
            {
                level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault();
            }

                // ✅ SAFE TRANSACTION MANAGEMENT: Validate document is still modifiable
                if (OptimizationFlags.UseSafeTransactionManagement && !_doc.IsModifiable)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Document became read-only during placement for zone {zone.Id}");
                    }
                    return null;
                }
                
                // ✅ CRITICAL FIX: Place instance WITHOUT level parameter to preserve absolute coordinates
                // The placement point is ABSOLUTE (already transformed to host document coordinates)
                // Passing level to NewFamilyInstance causes Revit to interpret Z as relative to level elevation
                FamilyInstance instance = _doc.Create.NewFamilyInstance(point, symbol, StructuralType.NonStructural);
                
                // ✅ Set level parameter separately (for schedule consistency, not for placement)
                if (instance != null && level != null)
                {
                    try
                    {
                        Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                        if (levelParam != null && !levelParam.IsReadOnly)
                        {
                            levelParam.Set(level.Id);
                        }
                        else
                        {
                            var levelByName = instance.LookupParameter("Level");
                            if (levelByName != null && !levelByName.IsReadOnly)
                            {
                                levelByName.Set(level.Id);
                            }
                        }
                    }
                    catch (Exception levelEx)
                    {
                        // Level parameter setting is optional - log but don't fail
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⚠️ Failed to set level parameter for zone {zone.Id}: {levelEx.Message}\n");
                        }
                    }
                }
                
                // ✅ SAFE ELEMENT VALIDATION: Validate instance was created successfully
                if (OptimizationFlags.UseSafeElementValidation)
                {
                    if (instance == null || !instance.IsValidObject)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] Failed to create instance for zone {zone.Id}");
                        }
                        return null;
                    }
                }
                
                // Apply rotation
                if (Math.Abs(rotation) > 1e-6)
                {
                    Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                    ElementTransformUtils.RotateElement(_doc, instance.Id, axis, rotation);
                }

                tracker?.SetItemCount(1);
                return instance;
            }
        }

        // ✅ SRP COMPLIANCE: All parameter setting methods have been moved to SleeveParameterService
        // The service is injected via constructor and used throughout this class
        
        /// <summary>
        /// ✅ FAMILY SYMBOL CACHING: Pre-cache family symbols with validation.
        /// Validates symbols before caching to prevent stale references.
        /// </summary>
        private void PreCacheFamilySymbols(List<ClashZone> clashZones)
        {
            try
            {
                var uniqueFamilyNames = new HashSet<string>();
                
                foreach (var zone in clashZones)
                {
                    // Determine if circular based on diameter
                    bool isCircular = zone.MepElementOuterDiameter > 0;
                    string familyName = GetSleeveFamilyName(zone, isCircular);
                    uniqueFamilyNames.Add(familyName);
                }
                
                foreach (var familyName in uniqueFamilyNames)
                {
                    // Skip if already cached and valid
                    if (_familySymbolCache.ContainsKey(familyName))
                    {
                        var cached = _familySymbolCache[familyName];
                        if (cached != null && cached.IsValidObject)
                        {
                            try
                            {
                                var _ = cached.IsActive; // Test access
                                continue; // Already cached and valid
                            }
                            catch (InvalidOperationException)
                            {
                                // Stale - remove and reload
                                _familySymbolCache.Remove(familyName);
                            }
                        }
                    }
                    
                    // Load and validate symbol
                    var symbol = LoadFamilySymbol(familyName);
                    if (symbol != null && symbol.IsValidObject)
                    {
                        try
                        {
                            // Test if symbol is active (will throw if stale)
                            var _ = symbol.IsActive;
                            _familySymbolCache[familyName] = symbol;
                        }
                        catch (InvalidOperationException)
                        {
                            // Symbol is stale - don't cache
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[NewSleevePlacer] Stale symbol detected for '{familyName}', skipping cache");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[NewSleevePlacer] Error pre-caching family symbols: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Immediately updates SleeveInstanceId in database after placement.
        /// This ensures the database is updated even if batch persistence is skipped or fails.
        /// Required for cleanup service to identify individual sleeves correctly.
        /// </summary>
        private void UpdateSleeveInstanceIdImmediately(Guid zoneId, int sleeveInstanceId)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    repository.UpdateSleeveInstanceId(zoneId, sleeveInstanceId);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ IMMEDIATE DB UPDATE: Set SleeveInstanceId={sleeveInstanceId} for zone {zoneId}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[NewSleevePlacer] Failed to update SleeveInstanceId immediately: {ex.Message}");
                }
                throw; // Re-throw so caller can handle
            }
        }
        
        private void UpdateSleeveDataInDatabase(ClashZone zone, FamilyInstance sleeve)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    
                    // Update Instance ID
                    repository.UpdateSleeveInstanceId(zone.Id, sleeve.Id.IntegerValue);
                    
                    // ✅ CRITICAL FIX: Get actual placement point from sleeve instance
                    // Don't rely on zone.SleevePlacementPointX/Y/Z which might be (0,0,0)
                    XYZ actualPlacementPoint = null;
                    if (sleeve.Location is LocationPoint locationPoint)
                    {
                        actualPlacementPoint = locationPoint.Point;
                    }
                    else if (sleeve.Location is LocationCurve locationCurve)
                    {
                        var curve = locationCurve.Curve;
                        if (curve != null)
                        {
                            actualPlacementPoint = curve.GetEndPoint(0);
                        }
                    }
                    
                    // Fallback to zone placement point if we can't get it from instance
                    if (actualPlacementPoint == null)
                    {
                        actualPlacementPoint = zone.SleevePlacementPoint ?? new XYZ(
                            zone.SleevePlacementPointX,
                            zone.SleevePlacementPointY,
                            zone.SleevePlacementPointZ);
                    }
                    
                    // Update Placement Data
                    repository.UpdateSleevePlacement(
                        zone.Id,
                        sleeve.Id.IntegerValue,
                        zone.SleeveWidth,
                        zone.SleeveHeight,
                        zone.SleeveDiameter,
                        actualPlacementPoint.X, // Use actual placement point from instance
                        actualPlacementPoint.Y,
                        actualPlacementPoint.Z,
                        actualPlacementPoint.X, // Active doc coords (same as placement point)
                        actualPlacementPoint.Y,
                        actualPlacementPoint.Z,
                        zone.MepElementRotationAngle
                    );
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[NewSleevePlacer] Failed to update DB for zone {zone.Id}: {ex.Message}");
            }
        }
        
        private void ResetReadyForPlacementFlags(List<Guid> zoneIds)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    repository.BulkResetReadyForPlacementFlags(zoneIds);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[NewSleevePlacer] Failed to reset flags: {ex.Message}");
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL FIX: Calculate bounding box from Width/Height/Depth dimensions when batch writing is enabled.
        /// When UseBatchedParameterWrites=true, sleeve.get_BoundingBox() returns STALE values because parameters
        /// haven't been flushed yet. This method constructs the bounding box from the calculated dimensions.
        /// </summary>
        /// <param name="sleeve">The sleeve family instance</param>
        /// <param name="width">Width dimension (in internal units, feet)</param>
        /// <param name="height">Height dimension (in internal units, feet)</param>
        /// <param name="depth">Depth dimension (in internal units, feet)</param>
        /// <param name="zone">The clash zone (for placement point and orientation)</param>
        /// <returns>Bounding box calculated from dimensions, or null if calculation fails</returns>
        private BoundingBoxXYZ CalculateBoundingBoxFromDimensions(FamilyInstance sleeve, double width, double height, double depth, ClashZone zone)
        {
            try
            {
                if (sleeve == null || !sleeve.IsValidObject || zone == null)
                    return null;
                
                // Get placement point from sleeve location
                var location = sleeve.Location as LocationPoint;
                XYZ placementPoint;
                
                if (location != null)
                {
                    placementPoint = location.Point;
                }
                else
                {
                    // Fallback: Use intersection point from zone
                    placementPoint = new XYZ(
                        zone.IntersectionPointX,
                        zone.IntersectionPointY,
                        zone.IntersectionPointZ
                    );
                }
                
                // ✅ Calculate bounding box centered at placement point
                // For rectangular openings: width (X), depth (Y), height (Z)
                var bbox = new BoundingBoxXYZ
                {
                    Min = new XYZ(
                        placementPoint.X - width / 2.0,
                        placementPoint.Y - depth / 2.0,
                        placementPoint.Z - height / 2.0
                    ),
                    Max = new XYZ(
                        placementPoint.X + width / 2.0,
                        placementPoint.Y + depth / 2.0,
                        placementPoint.Z + height / 2.0
                    ),
                    Enabled = true
                };
                
                return bbox;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] [CALC-BBOX] Error calculating bbox from dimensions for sleeve {sleeve?.Id?.IntegerValue ?? -1}: {ex.Message}\n");
                }
                return null;
            }
        }
    }
}
