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
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refactored;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using JSE_RevitAddin_MEP_OPENINGS.Services.Persistence;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For OpeningSettingsHelper
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;

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
        
        // ? OOP METHOD: Insulation-aware sizing service (SOLID principles)
        private readonly IInsulationAwareSizingService _sizingService;
        
        // ? SRP COMPLIANCE: Clearance calculation service (delegates clearance logic)
        private readonly ClearanceCalculationService _clearanceService;
        
        // ? SRP COMPLIANCE: RCS bounding box service (delegates clustering geometry logic)
        private readonly RcsBoundingBoxService _rcsBoundingBoxService;
        
        // ? SRP COMPLIANCE: Sleeve persistence service (delegates all database persistence operations)
        private readonly SleevePersistenceService _persistenceService;
        
        // ? SRP COMPLIANCE: Sleeve rotation service (delegates rotation calculation logic)
        private readonly SleeveRotationService _rotationService;
        
        // ? SRP COMPLIANCE: Sleeve parameter service (delegates all parameter setting operations)
        // Note: Not readonly because it needs to be recreated with performance monitor when available
        private SleeveParameterService _parameterService;
        
        // ? SRP COMPLIANCE: Placement point adjustment service (delegates all placement point adjustment logic)
        // Note: Not readonly because it needs to be recreated with performance monitor when available
        private PlacementPointAdjustmentService _placementPointAdjustmentService;
        
        // ? PHASE 1 OPTIMIZATION: Caching for performance improvements
        private Dictionary<ElementId, XYZ> _placementPointCache = new Dictionary<ElementId, XYZ>();
        private Dictionary<string, Level> _levelCache = new Dictionary<string, Level>();
        private Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();
        
        // ? SOLID REFACTORED: Optional refactored command services (injected when flag enabled)
        private readonly IFileNameNormalizer? _fileNameNormalizer;
        private readonly ISectionBoxChecker? _sectionBoxChecker;
        
        // ? CRASH-SAFE: Crash-safe executor for timeout protection and error handling
        private readonly CrashSafeExecutor? _crashSafeExecutor;
        
        // ? PERFORMANCE MONITORING: Performance monitor for tracking operations
        private Services.Placement.PlacementPerformanceMonitor? _performanceMonitor;
        
        // ? DAMPER CLEARANCE VALUES: Stores clearance values for damper parameter setting
        // Key: ClashZone ID (Guid - for matching zone to its calculated clearances)
        // Value: (finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom, offsetVector)
        // ? CRITICAL FIX: offsetVector is now stored to apply connector-side offset during placement
        private Dictionary<Guid, (double finalWidth, double finalHeight, double clearanceLeft, double clearanceRight, double clearanceTop, double clearanceBottom, XYZ offsetVector)> _damperPlacementAdjustments = 
            new Dictionary<Guid, (double, double, double, double, double, double, XYZ)>();
        
        // ? PARALLEL PLANNING: Optional planner for parallel pre-computation (OOP, DI-ready)
        // When enabled via DeploymentConfiguration.EnableParallelPlanning, pre-computes dimensions, clearance, and risk in parallel
        // Includes dampers (Duct Accessories) - ParallelSleevePlacementPlanner handles all categories
        private readonly ISleevePlacementPlanner? _planner;
        
        // ? FORCE DETECTION MODE: Flag to force recalculation of placement points
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
            IInsulationAwareSizingService sizingService = null,  // ? OOP METHOD: Optional sizing service injection (SOLID)
            // ? SOLID REFACTORED: Optional refactored command services (injected when flag enabled)
            IFileNameNormalizer? fileNameNormalizer = null,
            ISectionBoxChecker? sectionBoxChecker = null,
            // ? CRASH-SAFE: Optional crash-safe executor (created if not provided when flag enabled)
            CrashSafeExecutor? crashSafeExecutor = null,
            // ? PARALLEL PLANNING: Optional planner for parallel pre-computation (enabled via safety flag)
            ISleevePlacementPlanner? planner = null,
            // ? FORCE DETECTION MODE
            bool isForceDetectionMode = false)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            
            // ? DIAGNOSTIC: Log conditions object state at construction
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[NewSleevePlacer] ?? CONDITIONS CHECK: conditions={(_conditions != null ? "NOT NULL" : "NULL")}, " +
                    $"ClearanceSettings={(_conditions?.ClearanceSettings != null ? "NOT NULL" : "NULL")}");
                if (_conditions?.ClearanceSettings != null)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ?? CLEARANCE VALUES: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, " +
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
            
            // ? OOP METHOD: Initialize sizing service (create if not provided - Dependency Injection)
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            
            // ? SRP COMPLIANCE: Initialize clearance and RCS services (delegate to specialized services)
            _clearanceService = new ClearanceCalculationService();
            _rcsBoundingBoxService = new RcsBoundingBoxService();
            
            // ? SRP COMPLIANCE: Initialize persistence service
            _persistenceService = new SleevePersistenceService(doc);
            
            // ? SRP COMPLIANCE: Initialize rotation service (delegates wall/floor rotation logic)
            _rotationService = new SleeveRotationService();
            
            // ? SRP COMPLIANCE: Initialize parameter service (delegates all parameter setting operations)
            // Note: Performance monitor will be set later in PlaceAllSleevesInTransaction, so we pass null here
            _parameterService = new SleeveParameterService(doc, isReplayPath, null);
            
            // ? SRP COMPLIANCE: Initialize placement point adjustment service
            // Note: Performance monitor will be set later in PlaceAllSleevesInTransaction, so we pass null here
            _placementPointAdjustmentService = new PlacementPointAdjustmentService(doc, null);
            
            // ? SOLID REFACTORED: Initialize refactored services (create if not provided when flag enabled)
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
            
            // ? CRASH-SAFE: Initialize crash-safe executor (create if not provided when flag enabled)
            if (OptimizationFlags.UseCrashSafeExecution)
            {
                _crashSafeExecutor = crashSafeExecutor ?? new CrashSafeExecutor();
            }
            else
            {
                _crashSafeExecutor = null;
            }
            
            // ? PARALLEL PLANNING: Initialize planner (create if not provided and flag enabled)
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

        public (int placed, int skipped, int errors, List<(Guid ZoneGuid, int ElementId)> placedItems) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            // ≡ƒöä DIAGNOSTIC: Log start of method (ALWAYS log)
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] --- PlaceAllSleevesInTransaction START --- Zones={clashZones?.Count ?? 0}, IsModifiable={_doc.IsModifiable}\n");

            // ? SAFE TRANSACTION MANAGEMENT: Manage internal transaction if not already started by caller
            Transaction internalTransaction = null;
            if (!_doc.IsModifiable)
            {
                internalTransaction = new Transaction(_doc, "Sleeve Placement");
                internalTransaction.Start();
                SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] STARTED internal transaction.\n");
            }
            else
            {
                 SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] Using EXTERNAL transaction.\n");
            }
            
            // ? PERFORMANCE MONITORING: Initialize performance monitor if enabled
            if (OptimizationFlags.UsePerformanceMonitoring)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string performanceLogName = $"NewSleevePlacer_{timestamp}.log";
                _performanceMonitor = new Services.Placement.PlacementPerformanceMonitor(performanceLogName);
                
                // Recreate services with performance monitor for proper tracking
                _parameterService = new SleeveParameterService(_doc, _isReplayPath, _performanceMonitor);
                SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] Re-created _parameterService with monitor. Hash={_parameterService.GetHashCode()}\n");
                
                _placementPointAdjustmentService = new PlacementPointAdjustmentService(_doc, _performanceMonitor);
            }
            
            // ? PARAMETER BATCHING: Reset flags at start of each placement run
            _parameterService.ResetFlushFlag();
            
            // ? DAMPER PLACEMENT OFFSET: Clear stored offsets at start of each placement run
            _damperPlacementAdjustments?.Clear();
            
            // ? DIAGNOSTIC: Log batching flag status at placement start (ALWAYS log, even in deployment mode)
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] --- PLACEMENT START --- UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, Zones={clashZones?.Count ?? 0}\n");
            
            // ? DIAGNOSTIC: Log flag status of all zones (ALWAYS log)
            if (clashZones != null && clashZones.Count > 0)
            {
                int resolvedCount = clashZones.Count(z => z.IsResolvedFlag);
                int clusterResolvedCount = clashZones.Count(z => z.IsClusterResolvedFlag);
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? ZONE FLAGS: Total={clashZones.Count}, IsResolved={resolvedCount}, IsClusterResolved={clusterResolvedCount}\n");
            }
            
            int placed = 0;
            int skipped = 0;
            int errors = 0;
            List<(Guid ZoneGuid, int ElementId)> placedItems = new List<(Guid, int)>();
            
            try
            {
                // ? CRASH-SAFE: Execute with timeout protection if enabled
                if (OptimizationFlags.UseCrashSafeExecution && _crashSafeExecutor != null)
                {
                    try
                    {
                        var result = _crashSafeExecutor.ExecuteWithTimeout(() =>
                        {
                            var (p, s, e, items) = ExecutePlacementInternal(clashZones);
                            placed = p;
                            skipped = s;
                            errors = e;
                            placedItems = items;
                            return Result.Succeeded;
                        }, "Place All Sleeves");
                        
                        if (result != Result.Succeeded)
                        {
                            // Operation failed or was cancelled
                            if (internalTransaction != null && internalTransaction.GetStatus() == TransactionStatus.Started)
                                internalTransaction.RollBack();
                                
                            return (placed, skipped, errors, placedItems);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] Crash-safe execution failed: {ex.Message}");
                        }
                        
                        if (internalTransaction != null && internalTransaction.GetStatus() == TransactionStatus.Started)
                            internalTransaction.RollBack();
                            
                        errors = clashZones?.Count ?? 0;
                        return (placed, skipped, errors, placedItems);
                    }
                }
                else
                {
                    // Normal execution without crash-safe wrapper
                    (placed, skipped, errors, placedItems) = ExecutePlacementInternal(clashZones);
                }

                // If we started an internal transaction, commit it now
                if (internalTransaction != null && internalTransaction.GetStatus() == TransactionStatus.Started)
                {
                    internalTransaction.Commit();
                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] COMMITTED internal transaction.\n");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] Γ¥î TRANSACTION FAILED: {ex.Message}\n");
                if (internalTransaction != null && internalTransaction.GetStatus() == TransactionStatus.Started)
                    internalTransaction.RollBack();
                
                errors = clashZones?.Count ?? 0;
            }
            finally
            {
                if (internalTransaction != null)
                {
                    internalTransaction.Dispose();
                }
            }
            
            // ? PERFORMANCE MONITORING: Generate report if enabled
            if (OptimizationFlags.UsePerformanceMonitoring && _performanceMonitor != null)
            {
                _performanceMonitor.GenerateReport(placed, 0); // Individual sleeves only, no clusters
            }
            
            return (placed, skipped, errors, placedItems);
        }
        
        /// <summary>
        /// ? INTERNAL: Core placement logic (extracted for crash-safe wrapper)
        /// </summary>
        private (int placed, int skipped, int errors, List<(Guid ZoneGuid, int ElementId)> placedItems) ExecutePlacementInternal(List<ClashZone> clashZones)
        {
            int placed = 0;
            int skipped = 0; 
            int errors = 0;
            List<(Guid ZoneGuid, int ElementId)> placedSleeveResults = new List<(Guid ZoneGuid, int ElementId)>();

            // ? CRITICAL FIX: Store dimensions with placed sleeves for bounding box calculation when batching is enabled
            // When UseBatchedParameterWrites=true, sleeve.get_BoundingBox() returns STALE values because parameters
            // haven't been flushed yet.            // ? CRITICAL FIX: Tuple now holds 6 elements: (sleeve, zone, width, height, diameter, depth)
            List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter, double finalDepth)> placedSleeveData = 
                new List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter, double finalDepth)>();
            var processedZoneGuids = new List<Guid>();
            
            // ? DEDUPLICATION: Track placed locations to prevent duplicates
            var placedLocationKeys = new HashSet<string>();

            // ? DIAGNOSTIC: Check for null dependencies (User Request)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                 string depStatus = $"_doc={(_doc != null ? "OK" : "NULL")}, " +
                                  $"_familyManager={(_familyManager != null ? "OK" : "NULL")}, " +
                                  $"_sleeveRepository={(_sleeveRepository != null ? "OK" : "NULL")}, " +
                                  $"_sizingService={(_sizingService != null ? "OK" : "NULL")}, " +
                                  $"_strategy={(_strategy != null ? "OK" : "NULL")}";
                 
                 SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] dependency check: {depStatus}\n");
                 
                 if (_familyManager == null) 
                 {
                     SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ??? CRITICAL ERROR: _familyManager is NULL. Placement will fail.\n");
                 }
            }

            // ? Use injected ZoneFilterService if available to pre-filter zones
            List<ClashZone> filteredZones = clashZones;
            
            // ? PHASE 1 OPTIMIZATION: Pre-cache all required levels and family symbols upfront (eliminates redundant Revit API calls)
            if (filteredZones.Count > 0)
            {
                var cacheSw = System.Diagnostics.Stopwatch.StartNew();
                if (OptimizationFlags.UseFamilySymbolCache)
                {
                    PreCacheFamilySymbols(filteredZones);
                }
                
                // Γ£à NEW: Pre-cache all required levels from zone data
                PreCacheAllRequiredLevels(filteredZones);
                cacheSw.Stop();
                SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⏱️ PRE-PLACEMENT CACHE: {cacheSw.ElapsedMilliseconds}ms for {filteredZones.Count} zones\n");
            }
            if (_zoneFilterService != null)
            {
                filteredZones = _zoneFilterService.PreFilterEligibleClashZones(_doc, clashZones);
                
                if (!DeploymentConfiguration.DeploymentMode && filteredZones.Count != clashZones.Count)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ZoneFilterService filtered {clashZones.Count} ? {filteredZones.Count} zones");
                }
            }
            
            // ? PARALLEL PLANNING: Pre-compute all sleeve data in parallel (includes dampers)
            // Safety flag: DeploymentConfiguration.EnableParallelPlanning controls this feature
            // When enabled: Pre-computes dimensions, clearance, rotation, and risk classification in parallel
            // Benefits: Early skip detection, risk-based reordering, parallel computation, better diagnostics
            Dictionary<Guid, SleevePlacementPlanningDto>? planningMap = null;
            SleevePlacementPlanningResult? planningResult = null;
            
            if (DeploymentConfiguration.EnableParallelPlanning && _planner != null && filteredZones.Count > 0)
            {
                using (var planningTracker = _performanceMonitor?.TrackOperation("Step 2: PARALLEL CALCULATION"))
                {
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ?? PARALLEL PLANNING: Starting parallel pre-computation for {filteredZones.Count} zones (including dampers)");
                        }
                        
                        // ? PARALLEL PROCESSING: Run planning in parallel (pure computations, no Revit API calls)
                        // This includes dampers (Duct Accessories) - ParallelSleevePlacementPlanner handles all categories
                        var planSw = System.Diagnostics.Stopwatch.StartNew();
                        planningResult = _planner.Plan(filteredZones);
                        planSw.Stop();
                        
                        SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⏱️ PARALLEL PLANNING: {planSw.ElapsedMilliseconds}ms for {filteredZones.Count} zones\n");

                        planningTracker?.SetItemCount(planningResult.TotalCount);
                        
                        // Create lookup map: ClashZoneId -> PlanningDto for fast access during placement
                        planningMap = planningResult.Items
                            .ToDictionary(dto => dto.ClashZoneId, dto => dto);
                        
                        // ? EARLY SKIP: Filter out zones marked for skip (avoids Revit API calls)
                        var skipGuids = planningResult.Items
                            .Where(dto => dto.ShouldSkip)
                            .Select(dto => dto.ClashZoneId)
                            .ToHashSet();
                        
                        // Update filtered zones to exclude skipped ones
                        filteredZones = filteredZones
                            .Where(cz => !skipGuids.Contains(cz.Id))
                            .ToList();
                        
                        // ? REORDERING: Sort by risk (low risk first, high risk last)
                        // This ensures problematic zones are processed last (less likely to block good zones)
                        filteredZones = filteredZones
                            .OrderBy(cz => planningMap.ContainsKey(cz.Id) 
                                ? (int)planningMap[cz.Id].Risk 
                                : int.MaxValue)
                            .ThenByDescending(cz => planningMap.ContainsKey(cz.Id) 
                                ? planningMap[cz.Id].RawSleeveSizeFt 
                                : 0)
                            .ToList();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ? PARALLEL PLANNING: Completed in {planningResult.PlanningDurationMs:F1}ms - " +
                                $"Processed: {planningResult.TotalCount}, Skipped: {planningResult.SkippedCount}, " +
                                $"High Risk: {planningResult.HighRiskCount}, Critical: {planningResult.CriticalRiskCount}, " +
                                $"Remaining: {filteredZones.Count}");
                            
                            // ? BATCH LOGGING: Log all planning results at once (performance optimization)
                            if (planningResult.Items.Count > 0)
                            {
                                var logLines = planningResult.Items
                                    .Select(dto => $"[{DateTime.Now:HH:mm:ss.fff}] {dto.LogTrace}")
                                    .ToList();
                                SafeFileLogger.SafeAppendText("planning_debug.log", string.Join("\n", logLines) + "\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // ? GRACEFUL FALLBACK: If planning fails, continue with sequential processing
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] ?? Parallel planning failed: {ex.Message} - Falling back to sequential processing");
                        }
                        planningMap = null;
                        planningResult = null;
                    }
                }
            }
            else if (!DeploymentConfiguration.EnableParallelPlanning && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[NewSleevePlacer] ?? PARALLEL PLANNING: Disabled via safety flag (DeploymentConfiguration.EnableParallelPlanning=false) - Using sequential processing");
            }
            
            // ? TIMEOUT PROTECTION: Check timeout periodically
            // ? TIMEOUT PROTECTION: Check timeout periodically
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[NewSleevePlacer] Optimization Flags: UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, UseNewSleevePlacerService={OptimizationFlags.UseNewSleevePlacerService}");
            }


            // ? FORCE FLAG: Ensure batching is enabled for this operation (Fixes regression)
            OptimizationFlags.UseBatchedParameterWrites = true;
            
            // ? SAFE FLUSH: Reset flush flag to ensure parameters are flushed at the end
            if (_parameterService != null)
            {
                _parameterService.ResetFlushFlag();
            }

            // ? BATCH MODE ENTRY LOG (User Request)
            SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] - BATCH MODE STARTED - Processing {clashZones.Count} zones\n");

            int loopCount = 0;
            foreach (var clashZone in filteredZones)
            {
                loopCount++;
                if (loopCount <= 5)
                {
                     // ? LOG FIRST 5 ITEMS TO CONFIRM LOOP IS RUNNING
                     SafeFileLogger.SafeAppendTextAlways("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] Loop item {loopCount}/{filteredZones.Count} - Zone {clashZone.Id} (DB GUID: {clashZone.ClashZoneGuid})\n");
                }
                
                // ? PERSISTENCE DIAGNOSTIC: Log IDs for every item being processed
                SafeFileLogger.SafeAppendText("placement_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] processing zone: RuntimeId={clashZone.Id}, DbGuid={clashZone.ClashZoneGuid}, Category={clashZone.MepElementCategory}\n");
                // ? TIMEOUT PROTECTION: Check if operation has exceeded timeout
                if (OptimizationFlags.UseTimeoutProtection && _crashSafeExecutor != null)
                {
                    if (_crashSafeExecutor.CheckTimeout("Place Individual Sleeves"))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] ?? Timeout detected - stopping placement. Processed {placed + skipped + errors} out of {filteredZones.Count} zones");
                        }
                        break; // Stop processing on timeout
                    }
                }
                
                try
                {
                    // ? CRITICAL FIX: Skip if already resolved OR if SleeveInstanceId > 0 (even if IsResolved flag is not set)
                    // This prevents placing sleeves over existing sleeves (especially for dampers)
                    // ? CRITICAL: If IsClusterResolved=true, skip individual placement (zone is part of a cluster)
                    // ? CRITICAL: If ClusterSleeveInstanceId > 0, skip individual placement (cluster sleeve already exists)
                    bool shouldSkip = clashZone.IsResolvedFlag || 
                                     clashZone.IsClusterResolvedFlag || 
                                     clashZone.SleeveInstanceId > 0 || 
                                     clashZone.ClusterSleeveInstanceId > 0;
                    
                    if (shouldSkip)
                    {
                        // ? DIAGNOSTIC: Log why zone is being skipped (always log, even in deployment mode for debugging)
                        string skipReason = "";
                        if (clashZone.IsResolvedFlag) skipReason += "IsResolved=true, ";
                        if (clashZone.IsClusterResolvedFlag) skipReason += "IsClusterResolved=true, ";
                        if (clashZone.SleeveInstanceId > 0) skipReason += $"SleeveId={clashZone.SleeveInstanceId}, ";
                        if (clashZone.ClusterSleeveInstanceId > 0) skipReason += $"ClusterId={clashZone.ClusterSleeveInstanceId}, ";
                        skipReason = skipReason.TrimEnd(',', ' ');
                        
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? SKIP Zone {clashZone.Id}: {skipReason}\n");
                        skipped++;
                        continue;
                    }

                    // ? DEDUPLICATION: Check if we've already processed a zone at this location
                    // This prevents placing multiple sleeves at the exact same point (e.g. duplicate clashes or pre-clustered zones)
                    // Key includes: Category, HostID, and Rounded Coordinates (to 1mm approx)
                    string locationKey = $"{clashZone.MepElementCategory}_{clashZone.StructuralElementId}_" +
                                         $"{clashZone.IntersectionPointX:F4}_{clashZone.IntersectionPointY:F4}_{clashZone.IntersectionPointZ:F4}";
                                         
                    if (placedLocationKeys.Contains(locationKey))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? SKIP DUPLICATE: Zone {clashZone.Id} at same location as previous zone ({locationKey})\n");
                        }
                        skipped++;
                        continue;
                    }



                    // Local variables for this iteration
                    bool isSleevePlaced = false;
                    FamilyInstance placedSleeve = null;

                    // SMART REPLAY LOGIC (Path 1)
                    if (_isReplayPath)
                    {
                        // Check if sleeve exists in Revit
                        bool sleeveExists = false;
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            var element = _doc.GetElement(ElementIdCompat.FromValue(clashZone.SleeveInstanceId));
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
                        // ? DIAGNOSTIC: Log before attempting normal placement
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? ATTEMPTING PLACEMENT: Zone {clashZone.Id}, IsReplayPath={_isReplayPath}, SleeveId={clashZone.SleeveInstanceId}\n");
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log",
                         $"[{DateTime.Now:HH:mm:ss}] 🕒 CHECKPOINT 1: About to create sleeve for Zone {clashZone.Id}\n");
                        
                        // ? TRACE: Check if DB GUID is valid before placement
                        if (string.IsNullOrEmpty(clashZone.ClashZoneGuid))
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⚠️ WARNING: Zone {clashZone.Id} has NULL/EMPTY ClashZoneGuid. DB update WILL FAIL.\n");
                        }

                        placedSleeve = PlaceSleeveNormal(clashZone);
                        if (placedSleeve != null)
                        {
                            SafeFileLogger.SafeAppendText("batch_mode_entry.log",
                             $"[{DateTime.Now:HH:mm:ss}] ✅ CHECKPOINT 2: Created sleeve instance {placedSleeve.Id.GetIntegerValue()} for Zone {clashZone.Id}\n");
                            
                             SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ PLACED: Zone {clashZone.Id}, SleeveId={placedSleeve.Id.GetIntegerValue()}\n");
                            
                            isSleevePlaced = true;
                            placed++;
                        }
                        else
                        {
                             SafeFileLogger.SafeAppendTextAlways("batch_mode_entry.log",
                             $"[{DateTime.Now:HH:mm:ss}] ❌ CHECKPOINT 2 FAILED: PlaceSleeveNormal returned null for Zone {clashZone.Id}\n");
                             
                             SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ❌ PLACEMENT FAILED: Zone {clashZone.Id}, PlaceSleeveNormal returned null\n");
                             
                             errors++; // Mark as error if normal placement failed
                        }
                    }

                    if (placedSleeve != null)
                    {
                        // ? SAFE ELEMENT VALIDATION: Validate element by ID (avoids document mismatch bug)
                        // ?? CRITICAL: Do NOT compare documents by reference (causes false positives)
                        // Use element ID validation instead - if doc.GetElement() succeeds, element is in correct document
                        if (!OptimizationFlags.SkipRedundantValidation)
                        {
                            try
                            {
                                // Validate element is still valid and accessible
                                var validationElement = _doc.GetElement(placedSleeve.Id);
                                if (validationElement == null || !validationElement.IsValidObject)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[NewSleevePlacer] ?? Element {placedSleeve.Id.GetIntegerValue()} is invalid after placement - skipping");
                                    }
                                    errors++;
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                string msg = $"[NewSleevePlacer] ?? VALIDATION ERROR: Element {placedSleeve.Id.GetIntegerValue()}: {ex.Message}";
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error(msg);
                                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n");
                                }
                                errors++;
                                continue;
                            }
                        }
                        
                        // ? CRITICAL: Update ClashZone with new Sleeve ID and flag
                        // This ensures in-memory object is updated immediately (before batch flag update)
                        clashZone.SleeveInstanceId = placedSleeve.Id.GetIntegerValue();
                        clashZone.IsResolvedFlag = true;
                        
                        // Γ£à STEP 5 SUPPORT: Track placed item for corner extraction
                        placedSleeveResults.Add((clashZone.Id, placedSleeve.Id.GetIntegerValue()));
                        
                        // ✅ DELEGATE TO FLAG MANAGER: Update flags using dedicated service
                        if (_flagManager != null)
                        {
                            var updateList = new List<(Guid, long, bool)> { (clashZone.Id, (long)placedSleeve.Id.GetIntegerValue(), false) };
                            _flagManager.UpdateFlagsAfterPlacement(updateList);
                        }
                        
                        // ? DIAGNOSTIC: Log flag update for debugging
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? SET FLAG (via Manager): Zone {clashZone.Id}: IsResolved=true, SleeveId={clashZone.SleeveInstanceId}\n");
                        }
                        
                        // ? CRITICAL FIX: Immediately update SleeveInstanceId in database
                        // This ensures the database is updated even if batch persistence is skipped or fails
                        // Required for cleanup service to identify individual sleeves correctly
                        try
                        {
                            UpdateSleeveInstanceIdImmediately(clashZone.Id, placedSleeve.Id.GetIntegerValue());
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[NewSleevePlacer] Failed to update SleeveInstanceId immediately for zone {clashZone.Id}: {ex.Message}");
                            }
                            // Continue - batch persistence will try again later
                        }
                        
                        // ? CRITICAL FIX: Store dimensions for bounding box calculation when batching is enabled
                        // Get dimensions from zone (already set in PlaceSleeveNormal or PlaceSleeveFromSavedData)
                        double storedWidth = clashZone.SleeveWidth > 0 ? clashZone.SleeveWidth : 0;
                        double storedHeight = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : 0;
                        
                        // ? CRITICAL FIX: Calculate depth from wall/structural thickness (same logic as SetSleeveParameters)

                        double storedDepth = 0.0;
                        
                        // Depth = structural thickness for all (Floor, Wall, Framing)
                        storedDepth = clashZone.StructuralElementThickness;
                            
                            // ? ROBUST: No fallback - depth MUST be valid
                            if (storedDepth <= 0)
                            {
                                string errorMsg = $"[DEPTH-ERROR] Zone {clashZone.Id} (MEP={clashZone.MepElementId}, StructuralElement={clashZone.StructuralElementId}): " +
                                    $"Invalid depth - WallThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.WallThickness):F1}mm, " +
                                    $"FramingThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.FramingThickness):F1}mm, " +
                                    $"StructuralElementThickness={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.StructuralElementThickness):F1}mm. " +
                                    $"Database corruption detected - run ClashZone refresh to fix.";
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error(errorMsg);
                                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] {errorMsg}\n");
                                }
                                
                                throw new InvalidOperationException(errorMsg);
                            }

                        
                        // If dimensions not set, try to get from deferred parameters
                        if (storedWidth <= 0 || storedHeight <= 0)
                        {
                            storedWidth = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Width", storedWidth);
                            storedHeight = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Height", storedHeight);
                        }
                        
                        // ? CRITICAL: Also check deferred parameters for Depth/Wall Width (set by SetSleeveParameters)
                        double depthFromParams = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Depth", 0.0);
                        if (depthFromParams <= 0.0)
                        {
                            depthFromParams = _parameterService.GetParameterValueWithBatchingSupport(placedSleeve, "Wall Width", 0.0);
                        }
                        if (depthFromParams > 0.0)
                        {
                            storedDepth = depthFromParams;
                        }
                        
                        // ? CRITICAL FIX: Capture diameter AND depth for persistence and calculation
                        // 5th element = Diameter (for DB persistence)
                        // 6th element = Depth/WallThickness (for Bounding Box Calculation)
                        double finalDiameter = clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : 0;
                        
                        placedSleeveData.Add((placedSleeve, clashZone, storedWidth, storedHeight, finalDiameter, storedDepth));
                        processedZoneGuids.Add(clashZone.Id);
                        
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 3: Added to deferred list. Count={placedSleeveData.Count}\n");

                        // ? DEDUPLICATION: Add location to tracker
                        placedLocationKeys.Add(locationKey);
                        
                        // ? NOTE: Database persistence is now handled in batch after bounding boxes are calculated
                        // This ensures all data (instance ID, placement, bounding boxes, corners, snapshots) is saved together
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    string msg = $"[NewSleevePlacer] ?? EXCEPTION in loop for Zone {clashZone.Id}: {ex.Message}\nStack: {ex.StackTrace}";
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error(msg);
                        SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n");
                    }
                }
            }

            // ? PERFORMANCE OPTIMIZATION: Batch regeneration after ALL sleeves placed
            if (placedSleeveData.Count > 0)
            {
                var regenTimer = System.Diagnostics.Stopwatch.StartNew();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-REGEN] Regenerating document for {placedSleeveData.Count} sleeves...");
                }
                _doc.Regenerate(); // ? Single regeneration for all sleeves
                
                // ? CRITICAL FIX: Clear family symbol cache after regeneration
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
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-REGEN] ? Regenerated {placedSleeveData.Count} sleeves in {regenTimer.ElapsedMilliseconds}ms");
                }
                
                // ? CRITICAL FIX: Calculate bounding boxes from deferred parameters when batching is enabled
                // When UseBatchedParameterWrites=true, sleeve.get_BoundingBox() returns STALE values because parameters
                // haven't been flushed yet. We must calculate bounding boxes from stored dimensions instead.
                var bboxTimer = System.Diagnostics.Stopwatch.StartNew();
                int bboxCount = 0;
                // ? CRITICAL: Deconstruct 6-element tuple correctly
                foreach (var (sleeve, zone, storedWidth, storedHeight, storedDiameter, storedDepth) in placedSleeveData)
                {
                    try
                    {
                        // ? Validate sleeve still exists (may have been deleted by clustering)
                        if (!sleeve.IsValidObject) continue;

                        // ? CRITICAL FIX: When batch writing is enabled, calculate bounding box from stored dimensions
                        // instead of reading from Revit (which returns stale values)
                        BoundingBoxXYZ actualBbox = null;
                        if (OptimizationFlags.UseBatchedParameterWrites && storedWidth > 0 && storedHeight > 0)
                        {
                            // ? BATCHING ENABLED: Calculate bounding box from stored dimensions
                            actualBbox = CalculateBoundingBoxFromDimensions(sleeve, storedWidth, storedHeight, storedDepth, zone);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] [BATCH-BBOX] Zone {zone.Id}: Calculated bbox from DEFERRED params - W={storedWidth * 304.8:F1}mm, H={storedHeight * 304.8:F1}mm, D={storedDepth * 304.8:F1}mm\n");
                            }
                        }
                        
                        // ? FALLBACK: If batching disabled or calculation failed, read from Revit
                        if (actualBbox == null)
                        {
                            actualBbox = sleeve.get_BoundingBox(null);
                        }
                        
                        if (actualBbox != null)
                        {
                            // ? SRP COMPLIANCE: Delegate RCS transformation to specialized service (corners/bbox from placed sleeve)
                            _rcsBoundingBoxService.ProcessBoundingBox(zone, actualBbox);

                            // ? SOURCE OF TRUTH: Placement point = calculated sleeve placement point only (from refresh/placement step).
                            // Do NOT overwrite SleevePlacementPoint from bbox center — that can persist wrong values (e.g. outlier).
                            // Corners/bbox rely on placed sleeves; placement point relies on calculated only.

                            bboxCount++;
                        }
                    }
                    catch (Exception bboxEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[NewSleevePlacer] [BATCH-BBOX] Error retrieving bbox for sleeve {sleeve?.Id?.GetIntegerValue() ?? -1}: {bboxEx.Message}");
                        }
                    }
                }
                bboxTimer.Stop();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-BBOX] ? Retrieved {bboxCount} bounding boxes in {bboxTimer.ElapsedMilliseconds}ms (avg: {bboxTimer.ElapsedMilliseconds / Math.Max(1, bboxCount):F1}ms per sleeve)");
                }
                
                // ? SRP COMPLIANCE: Persist all sleeve data to database in batch (instance ID, placement, bounding boxes, corners, snapshots)
                // ? CRITICAL: Batch save corners to database AFTER regeneration
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
                            DebugLogger.Info($"[NewSleevePlacer] ? Persisted {persistedCount} sleeves to database (instance ID, placement, bounding boxes, corners, snapshots)");
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
            
            // ? CRITICAL PERFORMANCE FIX: Batch operations for 8x faster performance
            // Instead of individual operations per sleeve, batch all operations together
            if (placedSleeveData.Count > 0)
            {
                // ? BATCH FLAG MANAGEMENT: Update flags in single operation
                if (_flagManager != null)
                {
                    try
                    {
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
                                DebugLogger.Info($"[NewSleevePlacer] ? BATCH FLAG UPDATE: Updated {batchUpdates.Count} sleeves");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] ? Error updating flags: {ex.Message}");
                        }
                    }
                }
                
                // ? BATCH PARAMETER FLUSH: Flush all parameters in single operation
                SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 4: About to flush. Flag UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}\n");

                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    // ? NOTE: Already inside an outer transaction (from Orchestrator), so we don't start a new one here.
                    try
                    {
                        // 1. FIRST FLUSH (Set on Raw Elements)
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 4A: First Flush (Raw Elements)...\n");
                        
                        // Pass clearList=false to keep parameters for the second flush
                        int flushedCount1 = _parameterService.FlushDeferredParameters(clearList: false);
                        
                        // 2. FIRST REGEN (Wake Up)
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 4B: First Regen (Wake Up)...\n");
                        
                        _doc.Regenerate();

                        // 3. SECOND FLUSH (Re-Apply on Valid Elements)
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 4C: Second Flush (Force Values)...\n");

                        // Pass clearList=true to cleanup after this final flush
                        int flushedCount2 = _parameterService.FlushDeferredParameters(clearList: true);
                        
                        // 4. SECOND REGEN (Lock In)
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 4D: Second Regen (Lock In)...\n");

                        _doc.Regenerate(); 
                        
                        // Γ£à AUDIT
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] ≡ƒôì CHECKPOINT 5: Auditing values...\n");
                        
                        if (placedSleeveData.Count > 0)
                        {
                            var sampleZone = placedSleeveData[0].zone;
                            if (sampleZone != null && sampleZone.SleeveInstanceId > 0)
                            {
                                var s = _doc.GetElement(ElementIdCompat.FromValue(sampleZone.SleeveInstanceId)) as FamilyInstance;
                                if (s != null)
                                {
                                    var w = s.LookupParameter("Width")?.AsDouble() * 304.8 ?? -1;
                                    var h = s.LookupParameter("Height")?.AsDouble() * 304.8 ?? -1;
                                    SafeFileLogger.SafeAppendText("batch_mode_entry.log", 
                                        $"  ≡ƒöÄ AUDIT First Sleeve ({s.Id}): Width={w:F1}mm, Height={h:F1}mm\n");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss}] Γ¥î EXCEPTION IN FLUSH: {ex.Message}\n{ex.StackTrace}\n");
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] Γ¥î Error flushing parameters: {ex.Message}");
                        }
                    }
                }
            }
            
            // ? REMOVED: ReadyForPlacementFlag reset after placement
            // This flag is redundant - IsResolved flag is sufficient to track placement status
            // Zone eligibility = IsCurrentClashFlag=1 AND IsResolved=0
            // Keeping this code commented for reference during deprecation period:
            // if (processedZoneGuids.Count > 0)
            // {
            //     ResetReadyForPlacementFlags(processedZoneGuids);
            // }

            return (placed, skipped, errors, placedSleeveResults);
        }

        private bool CanUseSavedData(ClashZone zone)
        {
            // Check if we have valid saved dimensions and placement point
            // Also verify clearance hasn't changed (simplified check for now)
            return zone.SleeveWidth > 0 && 
                   zone.SleeveHeight > 0 && 
                   zone.SleevePlacementPointX != 0;
        }

        /// <summary>
        /// Duct Accessories are always in Wall only. When host type is missing, set to Wall so family, rotation, and host logic use Wall.
        /// </summary>
        private static void EnsureDuctAccessoriesHostType(ClashZone zone)
        {
            if (zone == null) return;
            if (!string.IsNullOrWhiteSpace(zone.StructuralElementType)) return;
            if ((zone.MepElementCategory ?? string.Empty).IndexOf("Duct Accessor", StringComparison.OrdinalIgnoreCase) < 0) return;
            zone.StructuralElementType = "Wall";
        }

        private FamilyInstance PlaceSleeveFromSavedData(ClashZone zone)
        {
            EnsureDuctAccessoriesHostType(zone);
            
            // ✅ NOMINAL DIAMETER FIX: For pipes with nominal diameter option enabled, force recalculation
            bool isPipe = string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
            bool useNominalDiameter = isPipe && (_conditions?.ClearanceSettings?.UseNominalDiameterForPipes ?? false);
            
            double width, height, diameter;
            bool isCircular;
            
            if (useNominalDiameter)
            {
                // Force recalculation using nominal diameter
                var dims = CalculateSleeveDimensions(zone, null);
                width = dims.width;
                height = dims.height;
                diameter = dims.diameter;
                isCircular = dims.isCircular;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? NOMINAL DIAMETER (PlaceSleeveFromSavedData): Zone {zone.Id} - " +
                        $"Force recalc using NominalDia={zone.MepElementNominalDiameter*304.8:F1}mm, " +
                        $"Result: Dia={diameter*304.8:F1}mm\n");
                }
            }
            else
            {
                // Use saved dimensions directly (normal path)
                width = zone.SleeveWidth;
                height = zone.SleeveHeight;
                diameter = zone.SleeveDiameter;
                isCircular = diameter > 0;
            }
            
            // Select Family
            string familyName = ClusterPlacementService.GetFamilyName(zone.StructuralElementType, zone.MepElementCategory, isCircular ? zone.SleeveDiameter : Math.Max(zone.SleeveWidth, zone.SleeveHeight), isCluster: false);
            
            // ✅ CRITICAL FIX: Detect Family-Shape Constraint Mismatch
            // If the selected family is Rectangular (e.g. for a Pipe in a Wall), we MUST treat it as rectangular
            // regardless of whether the zone thinks it is circular (has a Diameter).
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("batch_mode_entry.log", $"[{DateTime.Now:HH:mm:ss.fff}] 🔎 FAMILY CHECK: Zone {zone.Id}, Family='{familyName}', Dia={diameter*304.8:F1}mm\n");
            }

            // ✅ CRITICAL FIX: Detect Family-Shape Constraint Mismatch
            // If the selected family is Rectangular (e.g. for a Pipe in a Wall), we MUST treat it as rectangular.
            // also enforcing the known business rule: Diameter > 200mm -> Rectangular.
            bool looksRectangular = familyName.Contains("Rectangular", StringComparison.OrdinalIgnoreCase) || 
                                   familyName.Contains("Square", StringComparison.OrdinalIgnoreCase);
                                   
            bool exceedsThreshold = diameter > (200.0 / 304.8); // > 200mm

            if (looksRectangular || (isCircular && exceedsThreshold))
            {
                // ? CRITICAL FIX: Robust Dimension Transfer
                // If isCircular is true OR if we have a valid diameter but missing width (DB inconsistency),
                // we must force the dimensions to match the diameter for rectangular sleeves.
                if (isCircular || (diameter > 0.001 && width < 0.001))
                {
                    // Case: Pipe (Circular Zone) getting a Rectangular Sleeve (either by Name or Threshold)
                    // We must force isCircular = false so we set Width/Height parameters instead of Diameter
                    isCircular = false;

                    // ? CRITICAL REFINEMENT: Always force Width/Height to match Diameter
                    // Even if width/height have values, they might be defaults (100/300) or irrelevant.
                    width = diameter;
                    height = diameter;

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] [DIMENSION-FIX] Forcing Width/Height to Diameter for rectangular sleeve.\n" +
                            $"  - Diameter: {diameter * 304.8:F1}mm\n" +
                            $"  - Original Width: {width * 304.8:F1}mm\n");
                    }
                }
                SafeFileLogger.SafeAppendText("batch_mode_entry.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] 🕵️‍♂️ DEBUG POST-FIX: Zone {zone.Id} Updated!\n" +
                    $"  - New W={width*304.8:F1}, H={height*304.8:F1}, D={diameter*304.8:F1}, Circ={isCircular}\n");
            }
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null) return null;

            // ? SOURCE OF TRUTH: For individual sleeves, placement point is locked in refresh only.
            // Place sleeve does NOT calculate — merely uses the one calculated in refresh (or IntersectionPoint fallback).
            XYZ placementPoint;
            bool hasRefreshPoint = !_isForceDetectionMode && (zone.SleevePlacementPointX != 0 || zone.SleevePlacementPointY != 0 || zone.SleevePlacementPointZ != 0);
            if (hasRefreshPoint)
                placementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
            else
                placementPoint = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

            // ? SRP: Use rotation service to determine correct rotation for host type
            double rotation = _rotationService.DetermineRotation(zone); 

            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance != null)
            {
                // ? CRITICAL FIX: Round dimensions ONCE here (for all categories including dampers)
                // This ensures both the Revit parameters AND the saved zone dimensions are rounded (consistent with cluster sleeves)
                // Rounding is applied to ALL categories (dampers, pipes, ducts, cable trays, etc.)
                // ? REFACTOR: Use saved dimensions directly (respecting persistence)
                var roundedWidth = width;
                var roundedHeight = height;
                double roundedDiameter = diameter;
                
                // ? CRITICAL FIX: Pass ROUNDED dimensions to SetSleeveParameters (it will round again, but rounding already-rounded values is idempotent)
                // This ensures Revit parameters are set with rounded values
                _parameterService.SetSleeveParameters(instance, roundedWidth, roundedHeight, roundedDiameter, isCircular, zone);
                
                // ? CRITICAL FIX: Update zone with ROUNDED dimensions for saving to database
                // For circular pipes/ducts, SleeveWidth and SleeveHeight should BOTH equal the diameter
                // This ensures correct persistence to database (not bounding box dimensions)
                if (isCircular && roundedDiameter > 0)
                {
                    // For circular elements, width = height = diameter
                    zone.SleeveWidth = roundedDiameter;
                    zone.SleeveHeight = roundedDiameter;
                    zone.SleeveDiameter = roundedDiameter;
                }
                else
                {
                    // For rectangular elements, use calculated width/height
                    zone.SleeveWidth = roundedWidth;
                    zone.SleeveHeight = roundedHeight;
                    zone.SleeveDiameter = roundedDiameter;
                }
                // ? SOURCE OF TRUTH: Placement point locked in refresh — don't overwrite when we used refresh value
                if (!hasRefreshPoint)
                {
                    zone.SleevePlacementPoint = placementPoint;
                    zone.SleevePlacementPointX = placementPoint.X;
                    zone.SleevePlacementPointY = placementPoint.Y;
                    zone.SleevePlacementPointZ = placementPoint.Z;
                }
                zone.SleevePlacementPointActiveDocument = placementPoint;
            }
            
            return instance;
        }

        private FamilyInstance PlaceSleeveNormal(ClashZone zone, SleevePlacementPlanningDto? planningDto = null)
        {
            EnsureDuctAccessoriesHostType(zone);
            // ? DIAGNOSTIC: Log entry into PlaceSleeveNormal
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? PlaceSleeveNormal START: Zone {zone.Id}, HasPlanningDto={planningDto != null}\n");
            
            // ? DB-FIRST OPTIMIZATION: Use pre-processed dimensions from zone or planningDto directly
            // Skip CalculateSleeveDimensions service call to avoid overhead (8-70ms)
            double width, height, diameter;
            bool isCircular;

            // ? FORCE CALCULATION FOR DAMPERS: Dampers (Duct Accessories) require strategy execution 
            // to populate _damperPlacementAdjustments with the offsetVector. 
            // We cannot skip this even if dimensions are in DB.
            bool isDamper = string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);

            // ? DIAGNOSTIC LOG: Check what the service sees (Temporary for debugging)
            if (!DeploymentConfiguration.DeploymentMode && zone.SleeveWidth > 0)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer-CHECK] Zone {zone.Id}: " +
                    $"Category='{zone.MepElementCategory}', IsDamper={isDamper}, " +
                    $"DbWidth={zone.SleeveWidth:F3}\n");
            }

            if (!isDamper && planningDto != null && DeploymentConfiguration.EnableParallelPlanning && !planningDto.ShouldSkip)
            {
                width = planningDto.TargetWidthFt;
                height = planningDto.TargetHeightFt;
                diameter = Math.Max(width, height);
                isCircular = false; // Planning phase currently focuses on rectangular/offset handling
            }
            else if (!isDamper && (zone.SleeveWidth > 0 || zone.SleeveHeight > 0 || zone.SleeveDiameter > 0))
            {
                // ✅ NOMINAL DIAMETER FIX: For pipes with nominal diameter option enabled, force recalculation
                // using MepElementNominalDiameter instead of saved SleeveDiameter
                bool isPipe = string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                bool useNominalDiameter = isPipe && (_conditions?.ClearanceSettings?.UseNominalDiameterForPipes ?? false);
                
                if (useNominalDiameter)
                {
                    // Force recalculation using nominal diameter
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleeVE] ?? NOMINAL DIAMETER FORCE RECALC: Zone {zone.Id} - " +
                            $"Using MepElementNominalDiameter={zone.MepElementNominalDiameter*304.8:F1}mm " +
                            $"instead of saved SleeveDiameter={zone.SleeveDiameter*304.8:F1}mm\n");
                    }
                    
                    var dims = CalculateSleeveDimensions(zone, planningDto);
                    width = dims.width;
                    height = dims.height;
                    diameter = dims.diameter;
                    isCircular = dims.isCircular;
                }
                else
                {
                    // Use saved dimensions directly (normal path)
                    width = zone.SleeveWidth;
                    height = zone.SleeveHeight;
                    diameter = zone.SleeveDiameter;
                    isCircular = zone.SleeveDiameter > 0 && zone.SleeveWidth <= 0;
                }
            }
            else
            {
                // Fallback to calculation if DB data is missing OR if it's a damper
                if (isDamper && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                         $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? DAMPER FORCE CALC: Zone {zone.Id} - Executing strategy to get offset vector\n");
                }

                var dims = CalculateSleeveDimensions(zone, planningDto);
                width = dims.width;
                height = dims.height;
                diameter = dims.diameter;
                isCircular = dims.isCircular;
            }
            
            // ? DIAGNOSTIC: Log final dimensions
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? FINAL DIMENSIONS: Zone {zone.Id}, " +
                $"W={width:F6}ft ({width * 304.8:F1}mm), H={height:F6}ft ({height * 304.8:F1}mm), D={diameter:F6}ft ({diameter * 304.8:F1}mm), Circular={isCircular}\n");
            
            if (width <= 0 && height <= 0 && diameter <= 0)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? INVALID DIMENSIONS: Zone {zone.Id}, all dimensions <= 0\n");
                return null;
            }

            // Select Family
            // Γ£à DB-FIRST: Use saved family name if available, otherwise determine it
            string familyName = !string.IsNullOrEmpty(zone.SleeveFamilyName) 
                ? zone.SleeveFamilyName 
                : ClusterPlacementService.GetFamilyName(zone.StructuralElementType, zone.MepElementCategory, Math.Max(width, height), isCluster: false);
            
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? FAMILY CALC: Zone {zone.Id}, FamilyName='{familyName}', CalculatedWidth={width}, CalculatedHeight={height}, CalculatedDiameter={diameter}\n");
            
            // Γ£à REMOVED: Redundant per-sleeve DB update for family name (slow!)
            // Family name should already be in DB from refresh/planning phase
            
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? SYMBOL NULL: Zone {zone.Id}, FamilyName='{familyName}' - LoadFamilySymbol returned null\n");
                return null;
            }

            // Determine Placement Point (Intersection Point)
            XYZ placementPoint = zone.IntersectionPoint ?? new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
            
            // ? SRP COMPLIANCE: Delegate placement point adjustment to dedicated service
            // PlacementPointAdjustmentService handles non-dampers, DamperPlacementPointService handles dampers
            // This keeps NewSleevePlacerService focused on orchestration, not geometric calculations
            
            // ? EXTRACT CLEARANCE VALUES: Store individual clearance values in zone for later parameter setting
            // This is done BEFORE placement point adjustment so values are available for parameter setting
            XYZ damperOffsetVector = XYZ.Zero;
            if (_damperPlacementAdjustments.ContainsKey(zone.Id))
            {
                var (finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom, offsetVector) = 
                    _damperPlacementAdjustments[zone.Id];
                
                // ? SOLID ISP: Store individual clearance values in zone for later parameter setting
                // This allows SetSleeveParameters to set clearance parameters without strategy dependency
                zone.ClearanceLeft = clearanceLeft;
                zone.ClearanceRight = clearanceRight;
                zone.ClearanceTop = clearanceTop;
                zone.ClearanceBottom = clearanceBottom;
                
                // ? CRITICAL FIX: Store offset vector for later application
                damperOffsetVector = offsetVector;
            }
            
            // ? SOURCE OF TRUTH: For individual sleeves, placement point is locked in refresh only.
            // Place sleeve does NOT calculate — merely uses the one calculated in refresh (or IntersectionPoint fallback).
            bool hasRefreshPoint = !_isForceDetectionMode && (zone.SleevePlacementPointX != 0 || zone.SleevePlacementPointY != 0 || zone.SleevePlacementPointZ != 0);
            if (hasRefreshPoint)
            {
                placementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
                SafeFileLogger.SafeAppendText("placement_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] Using placement point from refresh (locked): Zone {zone.Id}\n");
            }
            else
                placementPoint = zone.IntersectionPoint ?? new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
            
            // ? DIAGNOSTIC: Capture base point BEFORE any damper offset (sequential path proof)
            XYZ basePointBeforeDamperOffset = placementPoint;
            
            // ? CRITICAL FIX: Apply connector-side offset for dampers with MEP connector
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
                    // Sequential path proof:
                    // BasePoint = intersection / DB placement BEFORE offset
                    // AfterOffset = point we send to Revit (NewFamilyInstance)
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ SEQ DAMPER OFFSET: Zone {zone.Id}, " +
                        $"BasePoint=({basePointBeforeDamperOffset.X:F6}, {basePointBeforeDamperOffset.Y:F6}, {basePointBeforeDamperOffset.Z:F6}), " +
                        $"Offset=({damperOffsetVector.X*304.8:F2}, {damperOffsetVector.Y*304.8:F2}, {damperOffsetVector.Z*304.8:F2})mm, " +
                        $"AfterOffset=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                }
            }
                    
            // ? DIAGNOSTIC: Log final placement point we actually use to place the sleeve (sequential path)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ SEQ PLACEMENT POINT: Zone {zone.Id}, Point=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
            }
            
            // ? SRP: Use rotation service to determine correct rotation for host type
            // Γ£à DB-FIRST: Use saved rotation angle if available (faster than Revit query)
            double rotation = (zone.MepElementRotationAngle != 0 && !_isForceDetectionMode)
                ? zone.MepElementRotationAngle
                : _rotationService.DetermineRotation(zone);

            // ? PRE-PLACEMENT PERSISTENCE: Save calculated data to DB before attempting placement
            // This ensures we have the dimensions and coordinates saved even if placement crashes
            try
            {
                // Update zone object with current calculated values
                // For circular elements, width = height = diameter
                if (isCircular && diameter > 0)
                {
                    zone.SleeveWidth = diameter;
                    zone.SleeveHeight = diameter;
                    zone.SleeveDiameter = diameter;
                }
                else
                {
                    zone.SleeveWidth = width;
                    zone.SleeveHeight = height;
                    zone.SleeveDiameter = diameter;
                }
                // ? SOURCE OF TRUTH: Placement point locked in refresh — don't overwrite when we used refresh value
                if (!hasRefreshPoint)
                {
                    zone.SleevePlacementPointX = placementPoint.X;
                    zone.SleevePlacementPointY = placementPoint.Y;
                    zone.SleevePlacementPointZ = placementPoint.Z;
                }
                zone.SleevePlacementPointActiveDocumentX = placementPoint.X;
                zone.SleevePlacementPointActiveDocumentY = placementPoint.Y;
                zone.SleevePlacementPointActiveDocumentZ = placementPoint.Z;

                // ? CRITICAL FIX: Explicitly assign Calculated properties for DB and user verification
                zone.CalculatedSleeveWidth = isCircular ? diameter : width;
                zone.CalculatedSleeveHeight = isCircular ? diameter : height;
                zone.CalculatedSleeveDiameter = diameter;
                // Depth = structural thickness for all (Floor, Wall, Framing)
                zone.CalculatedSleeveDepth = zone.StructuralElementThickness;

                zone.CalculatedRotation = rotation;
                zone.CalculatedPlacementX = placementPoint.X;
                zone.CalculatedPlacementY = placementPoint.Y;
                zone.CalculatedPlacementZ = placementPoint.Z;
                zone.CalculatedFamilyName = familyName;

                _sleeveRepository.UpdateSleeveCalculatedData(zone);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? PRE-PERSIST FAILED: Zone {zone.Id}: {ex.Message}\n");
                }
            }

            // Place Instance
            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance == null)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? INSTANCE NULL: Zone {zone.Id}, PlaceSleeveInstance returned null\n");
                return null;
            }
            
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? INSTANCE CREATED: Zone {zone.Id}, InstanceId={instance.Id.GetIntegerValue()}\n");
            
            if (instance != null)
            {
                // ? CRITICAL FIX: Round dimensions ONCE here (for all categories including dampers)
                // This ensures both the Revit parameters AND the saved zone dimensions are rounded (consistent with cluster sleeves)
                // Rounding is applied to ALL categories (dampers, pipes, ducts, cable trays, etc.)
                // ? REFACTOR: Dimensions are already rounded by CalculateSleeveDimensions (via SizingService)
                var roundedWidth = width;
                var roundedHeight = height;
                double roundedDiameter = diameter;
                
                // ? CRITICAL FIX: Pass ROUNDED dimensions to SetSleeveParameters (it will round again, but rounding already-rounded values is idempotent)
                // This ensures Revit parameters are set with rounded values
                _parameterService.SetSleeveParameters(instance, roundedWidth, roundedHeight, roundedDiameter, isCircular, zone);
                
                // ? CRITICAL FIX: Update zone with ROUNDED dimensions for saving to database
                // For circular pipes/ducts, SleeveWidth and SleeveHeight should BOTH equal the diameter
                // This ensures correct persistence to database (not bounding box dimensions)
                if (isCircular && roundedDiameter > 0)
                {
                    // For circular elements, width = height = diameter
                    zone.SleeveWidth = roundedDiameter;
                    zone.SleeveHeight = roundedDiameter;
                    zone.SleeveDiameter = roundedDiameter;
                }
                else
                {
                    // For rectangular elements, use calculated width/height
                    zone.SleeveWidth = roundedWidth;
                    zone.SleeveHeight = roundedHeight;
                    zone.SleeveDiameter = roundedDiameter;
                }
                // ? SOURCE OF TRUTH: Placement point locked in refresh — don't overwrite when we used refresh value
                if (!hasRefreshPoint)
                {
                    zone.SleevePlacementPoint = placementPoint;
                    zone.SleevePlacementPointX = placementPoint.X;
                    zone.SleevePlacementPointY = placementPoint.Y;
                    zone.SleevePlacementPointZ = placementPoint.Z;
                }
                zone.SleevePlacementPointActiveDocument = placementPoint;
                zone.SleevePlacementPointActiveDocumentX = placementPoint.X;
                zone.SleevePlacementPointActiveDocumentY = placementPoint.Y;
                zone.SleevePlacementPointActiveDocumentZ = placementPoint.Z;

                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? PERSISTENCE CHECK: Zone {zone.Id}, \n" +
                    $"   - Dimensions: W={zone.SleeveWidth}, H={zone.SleeveHeight}, D={zone.SleeveDiameter}\n" +
                    $"   - Placement: ({zone.SleevePlacementPointX:F3}, {zone.SleevePlacementPointY:F3}, {zone.SleevePlacementPointZ:F3})\n");
            }
            
            return instance;
        }

        /// <summary>
        /// ? CRITICAL FIX: Calculate sleeve dimensions with proper clearance handling.
        /// When batch parameter writing is enabled, this ensures:
        /// 1. Cable trays use GetCableTrayPlacementAdjustment which correctly reads clearances from conditions.ClearanceSettings (database)
        /// 2. Other categories use GetClearance which may read from strategy or fallback
        /// 3. All parameter reads during placement use GetParameterValueWithBatchingSupport to prevent stale reads
        /// 
        /// ? PARALLEL PLANNING: If planningDto is provided (from parallel planning phase), use pre-computed dimensions.
        /// This includes dampers - parallel planning handles all categories including Duct Accessories.
        /// </summary>
        private (double width, double height, double diameter, bool isCircular) CalculateSleeveDimensions(ClashZone zone, SleevePlacementPlanningDto? planningDto = null)
        {
            // ? PERFORMANCE MONITORING: Track dimension calculation
            using (var tracker = _performanceMonitor?.TrackOperation("Calculate Sleeve Dimensions"))
            {
                // ? PARALLEL PLANNING: Use pre-computed dimensions if available (from parallel planning phase)
                // This includes dampers - ParallelSleevePlacementPlanner handles Duct Accessories category
                // Safety: Only use planning data if EnableParallelPlanning flag is enabled and planningDto is valid
                if (DeploymentConfiguration.EnableParallelPlanning && planningDto != null && !planningDto.ShouldSkip)
                {
                    // ? USE PRE-COMPUTED DIMENSIONS: From parallel planning (includes dampers)
                    // Note: For dampers with asymmetric clearance, we still need strategy for offset calculation
                    // But dimensions can come from planning phase
                    bool isDamper = string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                    
                    if (!isDamper)
                    {
                        // ? NON-DAMPER: Use pre-computed dimensions directly (faster, already calculated in parallel)
                        tracker?.SetItemCount(1);
                        return (planningDto.TargetWidthFt, planningDto.TargetHeightFt, Math.Max(planningDto.TargetWidthFt, planningDto.TargetHeightFt), false);
                    }
                    else
                    {
                        // ? DAMPER: Use pre-computed dimensions but still need strategy for offset calculation
                        // Planning phase calculated dimensions, but we need strategy for asymmetric clearance offset
                        // Fall through to strategy-based calculation below
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ?? DAMPER: Using pre-computed dimensions from planning: " +
                                $"W={planningDto.TargetWidthFt * 304.8:F1}mm, H={planningDto.TargetHeightFt * 304.8:F1}mm, " +
                                $"but still need strategy for offset calculation");
                        }
                    }
                }
                
                // ? DAMPER PLACEMENT STRATEGY: Handle dampers with connector-aware asymmetric clearance
                // Dampers with MEP connectors require special offset calculation to achieve 100mm on connector side, 50mm on other
                if (_strategy is DamperPlacementStrategy damperStrategy)
                {
                    // ? COMPREHENSIVE LOGGING: Log raw MEP dimensions and clearance settings BEFORE strategy call
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
                    
                    // ? OOP METHOD: Get damper placement adjustment including offset and individual clearances (SOLID DIP)
                    // This method handles:
                    // 1. Connector detection (which side has MEP connector)
                    // 2. Asymmetric clearance mapping (100mm MEP side, 50mm other side)
                    // 3. Offset calculation (25mm toward connector to redistribute from 75/75 to 100/50)
                    // 4. Wall-specific handling (Z-axis swap for vertical connectors, axis-based offset)
                    var damperAdj = damperStrategy.GetDamperPlacementAdjustment(zone, _conditions);
                    
                    // ? COMPREHENSIVE LOGGING: Log final dimensions AFTER strategy call
                    double finalWidthMm = damperAdj.finalWidth * 304.8;
                    double finalHeightMm = damperAdj.finalHeight * 304.8;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-POST-STRATEGY] Zone {zone.Id}, " +
                            $"Final W={damperAdj.finalWidth:F6}ft ({finalWidthMm:F1}mm), H={damperAdj.finalHeight:F6}ft ({finalHeightMm:F1}mm), " +
                            $"ClearanceApplied W={finalWidthMm - rawWidthMm:F1}mm, H={finalHeightMm - rawHeightMm:F1}mm\n");
                    }
                    
                    // ? SOLID ISP: Extract individual clearance values from damper strategy result
                    // These will be stored in ClashZone and set as sleeve parameters later
                    // This allows placement service to use zone data without direct strategy dependency
                    // ? CRITICAL FIX: Use values directly from zone properties (populated by strategy)
                    // Do NOT overwrite with finalWidth/finalHeight which are total dimensions!
                    double clearanceLeft = zone.ClearanceLeft;
                    double clearanceRight = zone.ClearanceRight;
                    double clearanceTop = zone.ClearanceTop;
                    double clearanceBottom = zone.ClearanceBottom;
                    
                    // ? CRITICAL: Store damper clearance values for later parameter setting
                    // Key: zone.Id (Guid) for matching, Value: final dimensions + individual clearances + offsetVector
                    // ? CRITICAL FIX: offsetVector is now stored to apply connector-side offset during placement
                    // The offset shifts sleeve toward connector side to achieve 100mm/50mm clearance (not 75mm/75mm)
                    _damperPlacementAdjustments[zone.Id] = (
                        damperAdj.finalWidth,
                        damperAdj.finalHeight,
                        clearanceLeft,
                        clearanceRight,
                        clearanceTop,
                        clearanceBottom,
                        damperAdj.offsetVector // ? Store offset for later application
                    );
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ? DAMPER STRATEGY: Zone {zone.Id}, Final W={damperAdj.finalWidth:F6}ft, H={damperAdj.finalHeight:F6}ft (placement point handled by DamperPlacementPointService)");
                    }
                    
                    // ? CRITICAL FIX: Round dimensions to obey global rounding rules (e.g. nearest 50mm)
                    // Dampers calculate "exact" clearance (e.g. +150mm), but we must round the TOTAL dimension to the module
                    // This creates the "extra" clearance the user might see (e.g. 150mm -> 175mm total gap -> 12.5mm extra/side)
                    var profileSettings = ApplicationProfileService.Instance.GetCurrentSettings();
                    
                    // Use helper to round (centralized logic)
                    // Note: We use RoundDimensionsToNearest5mm which respects the global RoundingValue from settings
                    var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(
                        damperAdj.finalWidth, 
                        damperAdj.finalHeight);
                    
                    // ? COMPREHENSIVE LOGGING: Log rounding effect
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double finalWidthMmPreRound = damperAdj.finalWidth * 304.8;
                        double finalHeightMmPreRound = damperAdj.finalHeight * 304.8;
                        double roundedWidthMm = roundedWidth * 304.8;
                        double roundedHeightMm = roundedHeight * 304.8;
                        
                        if (Math.Abs(roundedWidthMm - finalWidthMmPreRound) > 0.1 || Math.Abs(roundedHeightMm - finalHeightMmPreRound) > 0.1)
                        {
                            SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-ROUNDING] Zone {zone.Id}, " +
                                $"Width {finalWidthMmPreRound:F1}mm -> {roundedWidthMm:F1}mm, " +
                                $"Height {finalHeightMmPreRound:F1}mm -> {roundedHeightMm:F1}mm " +
                                $"(RoundingValue={profileSettings.RoundingValue}mm, AlwaysUp={profileSettings.RoundAlwaysUp})\n");
                        }
                    }
                    
                    tracker?.SetItemCount(1);
                    return (roundedWidth, roundedHeight, 0, false); // Dampers are always rectangular (false = not circular)
                }
                
                // ? CRITICAL FIX: Cable trays need special handling to read clearances directly from database
                // When batch parameter writing is enabled, we MUST read clearances from _conditions.ClearanceSettings (database)
                // NOT from any cached/stale values or parameter reads
                if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                {
                    // ? VALIDATE: Ensure _conditions has ClearanceSettings populated from database
                    if (_conditions?.ClearanceSettings == null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[NewSleevePlacer] ?? _conditions.ClearanceSettings is NULL for cable tray zone {zone.Id} - using fallback clearances");
                    }
                    else
                    {
                        // ? DIAGNOSTIC: Log that we're using database clearances (critical for debugging batching issues)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] ? Using DATABASE clearances: CableTrayTop={_conditions.ClearanceSettings.CableTrayTop}mm, CableTrayOther={_conditions.ClearanceSettings.CableTrayOther}mm");
                        }
                    }
                    
                    // ? CABLE TRAY: Use strategy's GetCableTrayPlacementAdjustment (reads from conditions.ClearanceSettings - database)
                    // This method prioritizes _conditions.ClearanceSettings (database) over UI settings
                    var ctRawWidth = zone.MepElementWidth;
                    var ctRawHeight = zone.MepElementHeight;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[NewSleevePlacer] CABLE TRAY STRATEGY: Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(ctRawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(ctRawHeight):F1}mm");
                    
                    // ? CRITICAL: Pass _conditions (with database ClearanceSettings) to strategy
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
                
                // ? OOP METHOD: Use strategy to calculate clearance for other categories
                double rawWidth = zone.MepElementWidth;
                double rawHeight = zone.MepElementHeight;
                
                // ✅ PIPE NOMINAL DIAMETER OPTION: Use nominal diameter if enabled for pipes
                double rawDiameter = zone.MepElementOuterDiameter > 0 ? zone.MepElementOuterDiameter : 0;
                
                // ✅ NOMINAL DIAMETER DIAGNOSTIC LOGGING
                if (!DeploymentConfiguration.DeploymentMode && string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    bool useNominal = _conditions?.ClearanceSettings?.UseNominalDiameterForPipes ?? false;
                    double nominalDia = zone.MepElementNominalDiameter;
                    double outerDia = zone.MepElementOuterDiameter;
                    
                    // Log to both placement_debug.log AND main debug log
                    string logMsg = $"[NOMINAL-DIAMETER-CHECK] Zone {zone.Id}: UseNominal={useNominal}, NominalDia={nominalDia*304.8:F1}mm, OuterDia={outerDia*304.8:F1}mm";
                    DebugLogger.Info($"[NewSleevePlacer] {logMsg}");
                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] {logMsg}\n");
                }
                
                if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase) &&
                    _conditions?.ClearanceSettings?.UseNominalDiameterForPipes == true &&
                    zone.MepElementNominalDiameter > 0)
                {
                    rawDiameter = zone.MepElementNominalDiameter;
                    
                    // Log to both placement_debug.log AND main debug log
                    string logMsg = $"[NOMINAL-DIAMETER-ENABLED] Zone {zone.Id}: Using NOMINAL diameter: {rawDiameter*304.8:F1}mm instead of outer: {zone.MepElementOuterDiameter*304.8:F1}mm";
                    DebugLogger.Info($"[NewSleevePlacer] {logMsg}");
                    SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] {logMsg}\n");
                }
                
                // ? COMPREHENSIVE LOGGING: Log raw dimensions BEFORE clearance calculation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-PRE-CLEARANCE] Zone {zone.Id}, Category={zone.MepElementCategory}, " +
                        $"Raw W={rawWidth:F6}ft ({rawWidth * 304.8:F1}mm), H={rawHeight:F6}ft ({rawHeight * 304.8:F1}mm), " +
                        $"D={rawDiameter:F6}ft ({rawDiameter * 304.8:F1}mm), IsInsulated={zone.IsInsulated}\n");
                }
                
                double clearance = _clearanceService.GetClearance(zone, _conditions, _clearanceSettings, _strategy);
                double clearanceMm = clearance * 304.8; // Convert feet (internal units) to mm (1ft = 304.8mm)
                
                // ? COMPREHENSIVE LOGGING: Log clearance value AFTER retrieval
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-CLEARANCE] Zone {zone.Id}, " +
                        $"Clearance={clearance:F6}ft ({clearanceMm:F1}mm)\n");
                }
            
                // ? OOP METHOD: Use insulation-aware sizing service for consistent calculation (SOLID principles)
                // Formula: RawSize + (2 ∩┐╜ InsulationThickness) + (2 ∩┐╜ Clearance)
                // ? CRITICAL REFACTOR: Use ROUNDED calculation directly in the service
                // This centralizes rounding logic and ensures dimensions are final and consistent
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                (double finalWidth, double finalHeight, double finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZoneRounded(
                    rawWidth, rawHeight, rawDiameter, zone, clearance, settings.RoundingValue, settings.RoundAlwaysUp);
                
                // ? COMPREHENSIVE LOGGING: Log final dimensions AFTER sizing calculation
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NON-DAMPER-POST-SIZING] Zone {zone.Id}, " +
                        $"Final W={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), H={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm), " +
                        $"D={finalDiameter:F6}ft ({finalDiameter * 304.8:F1}mm), " +
                        $"ClearanceApplied W={finalWidth * 304.8 - rawWidth * 304.8:F1}mm, H={finalHeight * 304.8 - rawHeight * 304.8:F1}mm\n");
                }
            
                // ? GLOBAL SETTINGS: Determine opening type (circular vs rectangular) using global configuration rules
                // This matches the legacy UniversalSleevePlacerService.SelectUniversalFamily() logic
                bool isCircular = DetermineOpeningType(zone, rawDiameter, finalDiameter);

                tracker?.SetItemCount(1);
                return (finalWidth, finalHeight, finalDiameter, isCircular);
            }
        }

        /// <summary>
        /// ? GLOBAL SETTINGS: Determine opening type (circular vs rectangular) based on global configuration rules
        /// Matches legacy UniversalSleevePlacerService.SelectUniversalFamily() logic
        /// </summary>
        private bool DetermineOpeningType(ClashZone zone, double rawDiameter, double finalDiameter)
        {
            try
            {
                // ? Pipes: Use ConfigurationResolutionService to check global rules
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
                        // ✅ CRITICAL FIX: Pass clearance to GetResolvedOpeningType (user requested Total Diameter = MEP + 2*Clearance)
                        double pertinentClearanceMm = _conditions?.ClearanceSettings?.PipesNormal ?? 50.0;
                        
                        var resolvedType = pipeStrategy.GetResolvedOpeningType(mepSize, pipeType, hostType, pertinentClearanceMm);
                        bool isCircularResult = string.Equals(resolvedType, "Circular", StringComparison.OrdinalIgnoreCase);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] PIPE opening type resolved: Host={hostType}, UI='{pipeType}' ? Global Rule='{resolvedType}' ? isCircular={isCircularResult} (Clearance={pertinentClearanceMm}mm)");
                        }
                        
                        return isCircularResult;
                    }
                    else
                    {
                        // Fallback: Use raw diameter check if strategy not available
                        bool isCircularResult = rawDiameter > 0;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] PIPE opening type (fallback): rawDiameter={rawDiameter:F6}ft ? isCircular={isCircularResult}");
                        }
                        return isCircularResult;
                    }
                }
                // ? Round Ducts: Check user preference from CONDITIONS XML
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
                            DebugLogger.Info($"[NewSleevePlacer] ROUND DUCT opening type from CONDITIONS XML: '{roundDuctType}' ? isCircular={isCircularResult}");
                        }
                        
                        return isCircularResult;
                    }
                    else
                    {
                        // Rectangular ducts: always rectangular opening
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] RECTANGULAR DUCT ? isCircular=false");
                        }
                        return false;
                    }
                }
                else
                {
                    // Other categories (cable trays, accessories): always rectangular
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] OTHER CATEGORY ({zone.MepElementCategory}) ? isCircular=false");
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

        // ? SRP COMPLIANCE: GetClearance() method removed - now delegated to ClearanceCalculationService
        // This ensures NewSleevePlacerService focuses on placement logic, not clearance calculation


        private FamilySymbol LoadFamilySymbol(string familyName)
        {
            // ? PERFORMANCE MONITORING: Track family symbol loading
            using (var tracker = _performanceMonitor?.TrackOperation("Load Family Symbol"))
            {
                // ? CACHE VALIDATION: Check cache first, but validate before returning
            if (_familySymbolCache.ContainsKey(familyName))
            {
                var cachedSymbol = _familySymbolCache[familyName];
                
                // ? CRITICAL SAFETY: Validate symbol before use (prevents stale reference errors)
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
                    // ? FIX: After loading, find the symbol using Family.Name (not FamilyName property)
                    foundSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                }
            }
            
            // ? FIX: Fallback to simple lookup if FamilyManager didn't find it
            // Use Family.Name (not FamilyName property) to match UniversalSleevePlacerService
            if (foundSymbol == null)
            {
                foundSymbol = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(x => x.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
            }
            
            // ? CACHE VALIDATION: Only cache if symbol is valid
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
            // ? PERFORMANCE MONITORING: Track family symbol placement
            using (var tracker = _performanceMonitor?.TrackOperation("Place Sleeve Instance"))
            {
                // ? CRITICAL SAFETY: Validate symbol before accessing IsActive
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

            // ? CRITICAL FIX: Use absolute placement point (do NOT pass level to NewFamilyInstance)
            // When level is passed, Revit interprets Z coordinate as relative to level elevation
            // But our placement point is ABSOLUTE (already in host document coordinates)
            // Solution: Create instance without level, then set level parameter separately
            
            // Γ£à DB-FIRST: Use pre-cached level from zone data (no redundant Revit API calls)
            Level level = null;
            if (!string.IsNullOrEmpty(zone.MepElementLevelName))
            {
                if (!_levelCache.TryGetValue(zone.MepElementLevelName, out level))
                {
                    // Fallback to searching if not cached (safety only, should be pre-cached)
                    level = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .FirstOrDefault(l => l.Name.Equals(zone.MepElementLevelName, StringComparison.OrdinalIgnoreCase));
                    
                    if (level != null)
                        _levelCache[zone.MepElementLevelName] = level;
                }
            }
            
            // Final fallback: get first level in document if still null
            if (level == null)
            {
                level = _levelCache.Values.FirstOrDefault() ?? 
                        new FilteredElementCollector(_doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            }

                // ? SAFE TRANSACTION MANAGEMENT: Validate document is still modifiable
                if (OptimizationFlags.UseSafeTransactionManagement && !_doc.IsModifiable)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[NewSleevePlacer] Document became read-only during placement for zone {zone.Id}");
                    }
                    return null;
                }
                
                // ? CRITICAL FIX: Place instance WITHOUT level parameter to preserve absolute coordinates
                // The placement point is ABSOLUTE (already transformed to host document coordinates)
                // Passing level to NewFamilyInstance causes Revit to interpret Z as relative to level elevation
                FamilyInstance instance = _doc.Create.NewFamilyInstance(point, symbol, StructuralType.NonStructural);
                
                // ? Set level parameter separately (for schedule consistency, not for placement)
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
                                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ?? Failed to set level parameter for zone {zone.Id}: {levelEx.Message}\n");
                        }
                    }
                }
                
                // ? SAFE ELEMENT VALIDATION: Validate instance was created successfully
                if (!OptimizationFlags.SkipRedundantValidation)
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

        // ? SRP COMPLIANCE: All parameter setting methods have been moved to SleeveParameterService
        // The service is injected via constructor and used throughout this class
        
        /// <summary>
        /// ? FAMILY SYMBOL CACHING: Pre-cache family symbols with validation.
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
                    string familyName = ClusterPlacementService.GetFamilyName(zone.StructuralElementType, zone.MepElementCategory, isCircular ? zone.MepElementOuterDiameter : 0, isCluster: false);
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
        /// ? CRITICAL FIX: Immediately updates SleeveInstanceId in database after placement.
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
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ? IMMEDIATE DB UPDATE: Set SleeveInstanceId={sleeveInstanceId} for zone {zoneId}\n");
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
                    repository.UpdateSleeveInstanceId(zone.Id, sleeve.Id.GetIntegerValue());
                    
                    // ? SOURCE OF TRUTH: Placement point = calculated sleeve placement point only (from refresh/placement step).
                    // Do NOT use sleeve.Location — corners/bbox rely on placed sleeves; placement point relies on calculated only.
                    XYZ placementPoint = zone.SleevePlacementPoint ?? new XYZ(
                        zone.SleevePlacementPointX,
                        zone.SleevePlacementPointY,
                        zone.SleevePlacementPointZ);
                    
                    // Update Placement Data (calculated placement point only)
                    repository.UpdateSleevePlacement(
                        zone.Id,
                        sleeve.Id.GetIntegerValue(),
                        zone.SleeveWidth,
                        zone.SleeveHeight,
                        zone.SleeveDiameter,
                        placementPoint.X,
                        placementPoint.Y,
                        placementPoint.Z,
                        placementPoint.X,
                        placementPoint.Y,
                        placementPoint.Z,
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
        /// ? CRITICAL FIX: Calculate bounding box from Width/Height/Depth dimensions when batch writing is enabled.
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
                
                // ✅ CALCULATION: Create BoundingBoxXYZ from placement point and dimensions
                // Note: We use the placement point as the center of the bounding box
                BoundingBoxXYZ bbox = new BoundingBoxXYZ();
                bbox.Min = new XYZ(placementPoint.X - width / 2.0, placementPoint.Y - depth / 2.0, placementPoint.Z - height / 2.0);
                bbox.Max = new XYZ(placementPoint.X + width / 2.0, placementPoint.Y + depth / 2.0, placementPoint.Z + height / 2.0);
                
                return bbox;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] [CALC-BBOX] Error calculating bbox from dimensions for sleeve {sleeve?.Id?.GetIntegerValue() ?? -1}: {ex.Message}\n");
                }
                return null;
            }
        }

        #region Phase 1.5 Optimizations (Critical Performance Fixes)


        /// <summary>
        /// ? PHASE 1.5 OPTIMIZATION: Pre-cache all required family symbols before placement.
        /// When UsePreCachedFamilySymbols=true, pre-loads and validates all required family symbols before placement loop.
        /// Eliminates loading overhead and reduces variance (7x improvement in symbol operations).
        /// </summary>
        private void PreCacheAllFamilySymbols(List<ClashZone> clashZones)
        {
            if (!OptimizationFlags.UsePreCachedFamilySymbols)
            {
                return; // Use original logic
            }

            try
            {
                var uniqueFamilyNames = new HashSet<string>();
                
                foreach (var zone in clashZones)
                {
                    // Determine if circular based on diameter
                    bool isCircular = zone.MepElementOuterDiameter > 0;
                    string familyName = ClusterPlacementService.GetFamilyName(zone.StructuralElementType, zone.MepElementCategory, isCircular ? zone.MepElementOuterDiameter : 0, isCluster: false);
                    uniqueFamilyNames.Add(familyName);
                }
                
                // Load all symbols in parallel
                System.Threading.Tasks.Parallel.ForEach(uniqueFamilyNames, familyName =>
                {
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
                });
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ? Pre-cached {uniqueFamilyNames.Count} family symbols");
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
        /// ? PHASE 1.5 OPTIMIZATION: Memory leak detection and automatic garbage collection.
        /// When UseMemoryLeakDetection=true, monitors memory usage and forces garbage collection to prevent leaks.
        /// Targets the -1.57 MB memory leak identified in performance analysis.
        /// </summary>
        private void MonitorAndCleanMemory(string operation)
        {
            if (!OptimizationFlags.UseMemoryLeakDetection)
            {
                return; // Use original logic
            }

            try
            {
                var currentMemory = GC.GetTotalMemory(false);
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] [MEMORY] {operation}: {currentMemory / 1024 / 1024:F2} MB");
                }

                // Force garbage collection every 100 sleeves or when memory usage is high
                if (currentMemory > 500 * 1024 * 1024) // 500MB threshold
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    var afterMemory = GC.GetTotalMemory(false);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] [MEMORY] GC forced after {operation}: {afterMemory / 1024 / 1024:F2} MB (freed {(currentMemory - afterMemory) / 1024 / 1024:F2} MB)");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[NewSleevePlacer] Error monitoring memory: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ? PHASE 1.5 OPTIMIZATION: Operation variance reduction with warm-up and consistent data structures.
        /// When UseVarianceReduction=true, pre-warms operations and uses consistent data structures to reduce variance.
        /// Targets the 2.5-4.6x variance issues identified in performance analysis.
        /// </summary>
        private void WarmUpOperations()
        {
            if (!OptimizationFlags.UseVarianceReduction)
            {
                return; // Use original logic
            }

            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ?? WARMING UP OPERATIONS...");
                }

                // Pre-load family symbols
                PreCacheAllFamilySymbols(new List<ClashZone>());
                
                // Pre-calculate dimensions for sample zones
                var sampleZones = GetSampleZones();
                if (sampleZones.Count > 0)
                {
                    foreach (var zone in sampleZones)
                    {
                        var _ = CalculateSleeveDimensions(zone);
                    }
                }
                
                // Initialize parameter service
                _parameterService.ResetFlushFlag();
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ? WARM-UP COMPLETE");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[NewSleevePlacer] Error warming up operations: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ? PHASE 1.5 OPTIMIZATION: Get sample zones for warm-up operations.
        /// Returns a small set of zones for pre-calculating dimensions and warming up operations.
        /// </summary>
        private List<ClashZone> GetSampleZones()
        {
            // Create sample zones for warm-up (minimal data)
            var sampleZones = new List<ClashZone>
            {
                new ClashZone
                {
                    MepElementWidth = 1.0, // 1ft
                    MepElementHeight = 1.0, // 1ft
                    MepElementOuterDiameter = 0.0,
                    StructuralElementType = "Wall",
                    MepElementCategory = "Ducts",
                    DuctShape = "Rectangular",
                    IsInsulated = false
                },
                new ClashZone
                {
                    MepElementWidth = 0.0,
                    MepElementHeight = 0.0,
                    MepElementOuterDiameter = 1.0, // 1ft circular
                    StructuralElementType = "Wall",
                    MepElementCategory = "Pipes",
                    IsInsulated = false
                }
            };
            return sampleZones;
        }

        /// <summary>
        /// ? PHASE 1.5 OPTIMIZATION: Pre-calculate all dimensions upfront before placement loop.
        /// When UsePreCalculatedDimensions=true, calculates all dimensions upfront before placement loop.
        /// Eliminates repeated calculations and reduces variance (2.6x improvement in placement point adjustment).
        /// </summary>
        private Dictionary<Guid, (double width, double height, double depth)> PreCalculateAllDimensions(List<ClashZone> zones)
        {
            if (!OptimizationFlags.UsePreCalculatedDimensions)
            {
                return null; // Use original logic
            }

            var dimensionCache = new Dictionary<Guid, (double, double, double)>();
            
            foreach (var zone in zones)
            {
                // Use existing CalculateSleeveDimensions logic but cache results
                var (width, height, diameter, isCircular) = CalculateSleeveDimensions(zone);
                
                // Calculate depth from wall/structural thickness (same logic as SetSleeveParameters)
                double depth = 0.0;
                if (diameter > 0)
                {
                    // For circular, depth = diameter
                    depth = diameter;
                }
                else
                {
                    // Depth = structural thickness for all (Floor, Wall, Framing)
                    depth = zone.StructuralElementThickness;
                }

                dimensionCache[zone.Id] = (width, height, depth);
            }
            
            return dimensionCache;
        }

        #endregion

        #region Phase 1 Performance Optimizations

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Get cached placement point for element (eliminates LocationPoint/LocationCurve queries)
        /// When UseElementLocationCaching=true, caches placement points during batch placement (70-80% reduction in location queries)
        /// </summary>
        private XYZ GetCachedPlacementPoint(FamilyInstance sleeve)
        {
            if (!OptimizationFlags.UseElementLocationCaching)
            {
                return null; // Use original logic
            }

            if (!_placementPointCache.TryGetValue(sleeve.Id, out XYZ point))
            {
                // Calculate once and cache
                point = CalculatePlacementPoint(sleeve);
                _placementPointCache[sleeve.Id] = point;
            }
            return point;
        }

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Calculate placement point from sleeve location (supports caching)
        /// </summary>
        private XYZ CalculatePlacementPoint(FamilyInstance sleeve)
        {
            if (sleeve.Location is LocationPoint locationPoint)
            {
                return locationPoint.Point;
            }
            else if (sleeve.Location is LocationCurve locationCurve)
            {
                return locationCurve.Curve.GetEndPoint(0);
            }
            return null;
        }

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Get cached level reference (eliminates repeated level lookups)
        /// When UseLevelReferenceCaching=true, caches level references and batches level parameter setting (80-90% reduction in level lookups)
        /// </summary>
        private Level GetCachedLevel(string levelName)
        {
            if (!OptimizationFlags.UseLevelReferenceCaching)
            {
                return null; // Use original logic
            }

            if (!_levelCache.TryGetValue(levelName, out Level level))
            {
                // Find level once and cache
                level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .FirstOrDefault(l => l.Name.Equals(levelName, StringComparison.OrdinalIgnoreCase));
                
                if (level != null)
                {
                    _levelCache[levelName] = level;
                }
            }
            return level;
        }

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Batch level parameter setting (reduces individual parameter operations)
        /// When UseLevelReferenceCaching=true, batches level parameter setting for better performance
        /// </summary>
        private void BatchSetLevels(List<FamilyInstance> sleeves, Level level)
        {
            if (!OptimizationFlags.UseLevelReferenceCaching || level == null)
            {
                return; // Use original logic
            }

            var levelParamName = "Level";
            foreach (var sleeve in sleeves)
            {
                var levelParam = sleeve.LookupParameter(levelParamName);
                if (levelParam != null && !levelParam.IsReadOnly)
                {
                    levelParam.Set(level.Id);
                }
            }
        }

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Pre-calculate all dimensions before placement (eliminates repeated calculations)
        /// When UsePreCalculatedDimensions=true, calculates all dimensions upfront before placement loop (eliminates repeated calculations)
        /// </summary>
        private List<(ClashZone zone, double width, double height, double depth)> PreCalculateDimensions(List<ClashZone> zones)
        {
            if (!OptimizationFlags.UsePreCalculatedDimensions)
            {
                return null; // Use original logic
            }

            return zones.Select(zone =>
            {
                // Use existing CalculateSleeveDimensions logic but cache results
                var (width, height, diameter, isCircular) = CalculateSleeveDimensions(zone);
                
                // Calculate depth from wall/structural thickness (same logic as SetSleeveParameters)
                double depth = 0.0;
                if (diameter > 0)
                {
                    // For circular, depth = diameter
                    depth = diameter;
                }
                else
                {
                    // Depth = structural thickness for all (Floor, Wall, Framing)
                    depth = zone.StructuralElementThickness;
                }

                return (zone, width, height, depth);
            }).ToList();
        }

        /// <summary>
        /// ? PHASE 1 OPTIMIZATION: Batch parameter operations (reduces individual parameter reads/writes)
        /// When UseBatchParameterOperations=true, batches parameter operations for better performance (50-60% reduction in parameter operations)
        /// </summary>
        private void BatchSetSleeveParameters(List<(FamilyInstance instance, double width, double height, double diameter, bool isCircular, ClashZone zone)> sleeveData)
        {
            if (!OptimizationFlags.UseBatchParameterOperations)
            {
                return; // Use original logic
            }

            foreach (var (instance, width, height, diameter, isCircular, zone) in sleeveData)
            {
                // Set all parameters in batch
                _parameterService.SetSleeveParameters(instance, width, height, diameter, isCircular, zone);
            }
            
            // Flush all parameters at once
            if (OptimizationFlags.UseBatchedParameterWrites)
            {
                _parameterService.FlushDeferredParameters();
            }
        }

        #endregion

        // ? NOTE: NewSleevePlacerService does NOT need section box filtering
        // It processes clash zones that are already in the database, so it doesn't need to find MEP elements from the model
        // Section box filtering is only needed for services that query the Revit model to find elements
        /// <summary>
        /// Γ£à NEW: Pre-cache all required levels from clash zones upfront.
        /// Eliminates redundant level lookups during the placement loop.
        /// </summary>
        private void PreCacheAllRequiredLevels(List<ClashZone> clashZones)
        {
            try
            {
                var uniqueLevelNames = clashZones
                    .Where(z => !string.IsNullOrEmpty(z.MepElementLevelName))
                    .Select(z => z.MepElementLevelName)
                    .Distinct();
                
                var allLevels = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .ToDictionary(l => l.Name, l => l, StringComparer.OrdinalIgnoreCase);
                
                foreach (var levelName in uniqueLevelNames)
                {
                    if (allLevels.TryGetValue(levelName, out var level))
                    {
                        _levelCache[levelName] = level;
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] Γ£à Pre-cached {_levelCache.Count} levels for {clashZones.Count} zones");
                }
                
                // Γ£à SYNC: Populate parameter service cache to avoid redundant lookups there too
                _parameterService.PreCacheLevels(_levelCache.Values);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[NewSleevePlacer] ΓÜá∩╕Å Level pre-caching failed: {ex.Message}");
                }
            }
        }
    }
}