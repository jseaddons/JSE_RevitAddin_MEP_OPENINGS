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
        
        // ✅ SOLID REFACTORED: Optional refactored command services (injected when flag enabled)
        private readonly IFileNameNormalizer? _fileNameNormalizer;
        private readonly ISectionBoxChecker? _sectionBoxChecker;
        
        // ✅ CRASH-SAFE: Crash-safe executor for timeout protection and error handling
        private readonly CrashSafeExecutor? _crashSafeExecutor;
        
        // ✅ PERFORMANCE MONITORING: Performance monitor for tracking operations
        private Services.Placement.PlacementPerformanceMonitor? _performanceMonitor;
        
        // Family symbol cache for performance
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();
        
        // ✅ PARAMETER BATCHING: Deferred parameter writes (4-6× faster placement)
        // Accumulates parameter values during placement loop, writes all after single regeneration
        // Key: ElementId of sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private Dictionary<ElementId, Dictionary<string, object>> _deferredParameters = new Dictionary<ElementId, Dictionary<string, object>>();
        
        // ✅ SAFETY FLAG: Prevents multiple flushes (critical for performance)
        private bool _hasFlushedParameters = false;
        
        // ✅ DAMPER PLACEMENT OFFSET: Stores offset for asymmetric clearance placement
        // Key: ClashZone ID (Guid - for matching zone to its calculated offset)
        // Value: (offsetVector, finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom)
        private Dictionary<Guid, (XYZ offsetVector, double finalWidth, double finalHeight, double clearanceLeft, double clearanceRight, double clearanceTop, double clearanceBottom)> _damperPlacementAdjustments = 
            new Dictionary<Guid, (XYZ, double, double, double, double, double, double)>();

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
            CrashSafeExecutor? crashSafeExecutor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _sleeveRepository = sleeveRepository ?? throw new ArgumentNullException(nameof(sleeveRepository));
            _zoneFilterService = zoneFilterService; // Can be null for now
            _familyManager = familyManager; // Can be null for now
            _flagManager = flagManager; // Passed as IFlagManager? (can be null)
            _isReplayPath = isReplayPath;
            _filterName = filterName;
            
            // ✅ OOP METHOD: Initialize sizing service (create if not provided - Dependency Injection)
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            
            // ✅ SRP COMPLIANCE: Initialize clearance and RCS services (delegate to specialized services)
            _clearanceService = new ClearanceCalculationService();
            _rcsBoundingBoxService = new RcsBoundingBoxService();
            
            // ✅ SRP COMPLIANCE: Initialize persistence service
            _persistenceService = new SleevePersistenceService(doc);
            
            // ✅ SRP COMPLIANCE: Initialize rotation service (delegates wall/floor rotation logic)
            _rotationService = new SleeveRotationService();
            
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
            }
            
            // ✅ PARAMETER BATCHING: Reset flags at start of each placement run
            _hasFlushedParameters = false;
            _deferredParameters?.Clear();
            
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
                    // Skip if already resolved (unless we are forcing update, but typically we skip)
                    if (clashZone.IsResolved || clashZone.IsClusterResolved)
                    {
                        // ✅ DIAGNOSTIC: Log why zone is being skipped (always log, even in deployment mode for debugging)
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ⏭️ SKIP Zone {clashZone.Id}: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, SleeveId={clashZone.SleeveInstanceId}, ClusterId={clashZone.ClusterSleeveInstanceId}\n");
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
                        
                        // Update ClashZone with new Sleeve ID
                        clashZone.SleeveInstanceId = placedSleeve.Id.IntegerValue;
                        clashZone.IsResolved = true;
                        
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
                            if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null)
                            {
                                var sleeveId = placedSleeve.Id;
                                if (_deferredParameters.ContainsKey(sleeveId))
                                {
                                    var paramsDict = _deferredParameters[sleeveId];
                                    if (paramsDict.ContainsKey("Width") && paramsDict["Width"] is double w) storedWidth = w;
                                    if (paramsDict.ContainsKey("Height") && paramsDict["Height"] is double h) storedHeight = h;
                                }
                            }
                        }
                        
                        // ✅ CRITICAL: Also check deferred parameters for Depth/Wall Width (set by SetSleeveParameters)
                        if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null)
                        {
                            var sleeveId = placedSleeve.Id;
                            if (_deferredParameters.ContainsKey(sleeveId))
                            {
                                var paramsDict = _deferredParameters[sleeveId];
                                if (paramsDict.ContainsKey("Depth") && paramsDict["Depth"] is double d && d > 0) storedDepth = d;
                                else if (paramsDict.ContainsKey("Wall Width") && paramsDict["Wall Width"] is double ww && ww > 0) storedDepth = ww;
                            }
                        }
                        
                        placedSleeveData.Add((placedSleeve, clashZone, storedWidth, storedHeight, storedDepth));
                        processedZoneGuids.Add(clashZone.Id);
                        
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
            
            // Batch update flags at the end
            if (placedSleeveData.Count > 0)
            {
                try
                {
                    var batchUpdates = placedSleeveData
                        .Where(x => x.zone.SleeveInstanceId > 0)
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
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[NewSleevePlacer] Error updating flags: {ex.Message}");
                }
            }
            
            // ✅ PARAMETER BATCHING: Flush deferred parameters AFTER bounding box calculation
            // Bounding boxes must be calculated BEFORE flushing, so they use deferred parameter values
            if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null && _deferredParameters.Count > 0)
            {
                try
                {
                    int flushedCount = FlushDeferredParameters();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ✅ Flushed {flushedCount} deferred parameters for {_deferredParameters.Count} sleeves");
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
                SetSleeveParameters(instance, width, height, diameter, isCircular, zone);
                
                // ✅ CRITICAL FIX: Update zone with calculated dimensions for bounding box calculation
                // This ensures dimensions are available when batching is enabled
                zone.SleeveWidth = width;
                zone.SleeveHeight = height;
                zone.SleeveDiameter = diameter;
                zone.SleevePlacementPoint = placementPoint;
                zone.SleevePlacementPointX = placementPoint.X;
                zone.SleevePlacementPointY = placementPoint.Y;
                zone.SleevePlacementPointZ = placementPoint.Z;
            }
            
            return instance;
        }

        private FamilyInstance PlaceSleeveNormal(ClashZone zone)
        {
            // ✅ DIAGNOSTIC: Log entry into PlaceSleeveNormal
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 🔨 PlaceSleeveNormal START: Zone {zone.Id}\n");
            
            // Calculate dimensions based on MEP element and clearance using strategy
            var (width, height, diameter, isCircular) = CalculateSleeveDimensions(zone);
            
            // ✅ DIAGNOSTIC: Log calculated dimensions
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] 📐 DIMENSIONS: Zone {zone.Id}, W={width:F6}ft, H={height:F6}ft, D={diameter:F6}ft, Circular={isCircular}\n");
            
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
            
            // ✅ DAMPER OFFSET: Apply offset for asymmetric clearance placement
            // If damper placement adjustment was calculated, apply the offset to achieve correct clearance distribution
            if (_damperPlacementAdjustments.ContainsKey(zone.Id))
            {
                var (offsetVector, finalWidth, finalHeight, clearanceLeft, clearanceRight, clearanceTop, clearanceBottom) = 
                    _damperPlacementAdjustments[zone.Id];
                
                // ✅ CRITICAL: Apply offset to placement point to move sleeve toward connector side
                placementPoint = placementPoint.Add(offsetVector);
                
                // ✅ SOLID ISP: Store individual clearance values in zone for later parameter setting
                // This allows SetSleeveParameters to set clearance parameters without strategy dependency
                zone.ClearanceLeft = clearanceLeft;
                zone.ClearanceRight = clearanceRight;
                zone.ClearanceTop = clearanceTop;
                zone.ClearanceBottom = clearanceBottom;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ✅ DAMPER OFFSET APPLIED: Zone {zone.Id}, Offset=({offsetVector.X:F6}, {offsetVector.Y:F6}, {offsetVector.Z:F6}), New Placement=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3})");
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [NewSleevePlacer] ✅ DAMPER OFFSET: Zone {zone.Id}, Offset Vector=({offsetVector.X:F6}, {offsetVector.Y:F6}, {offsetVector.Z:F6}), Applied to Point ({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3})\n");
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
                SetSleeveParameters(instance, width, height, diameter, isCircular, zone);
                
                // Update zone with calculated dimensions for saving
                zone.SleeveWidth = width;
                zone.SleeveHeight = height;
                zone.SleeveDiameter = diameter;
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
        /// </summary>
        private (double width, double height, double diameter, bool isCircular) CalculateSleeveDimensions(ClashZone zone)
        {
            // ✅ PERFORMANCE MONITORING: Track dimension calculation
            using (var tracker = _performanceMonitor?.TrackOperation("Calculate Sleeve Dimensions"))
            {
                // ✅ DAMPER PLACEMENT STRATEGY: Handle dampers with connector-aware asymmetric clearance
                // Dampers with MEP connectors require special offset calculation to achieve 100mm on connector side, 50mm on other
                if (_strategy is DamperPlacementStrategy damperStrategy)
                {
                    // ✅ OOP METHOD: Get damper placement adjustment including offset and individual clearances (SOLID DIP)
                    // This method handles:
                    // 1. Connector detection (which side has MEP connector)
                    // 2. Asymmetric clearance mapping (100mm MEP side, 50mm other side)
                    // 3. Offset calculation (25mm toward connector to redistribute from 75/75 to 100/50)
                    // 4. Wall-specific handling (Z-axis swap for vertical connectors, axis-based offset)
                    var damperAdj = damperStrategy.GetDamperPlacementAdjustment(zone, _conditions);
                    
                    // ✅ SOLID ISP: Extract individual clearance values from damper strategy result
                    // These will be stored in ClashZone and set as sleeve parameters later
                    // This allows placement service to use zone data without direct strategy dependency
                    double clearanceLeft = damperAdj.finalWidth > 0 ? damperAdj.finalWidth : zone.ClearanceLeft;
                    double clearanceRight = damperAdj.finalHeight > 0 ? damperAdj.finalHeight : zone.ClearanceRight;
                    double clearanceTop = damperAdj.finalWidth > 0 ? damperAdj.finalHeight : zone.ClearanceTop;
                    double clearanceBottom = damperAdj.finalHeight > 0 ? damperAdj.finalWidth : zone.ClearanceBottom;
                    
                    // ✅ CRITICAL: Store damper placement adjustment for later use in PlaceSleeveNormal
                    // Key: zone.Id (Guid) for matching, Value: offset vector + final dimensions + individual clearances
                    _damperPlacementAdjustments[zone.Id] = (
                        damperAdj.offsetVector,
                        damperAdj.finalWidth,
                        damperAdj.finalHeight,
                        clearanceLeft,
                        clearanceRight,
                        clearanceTop,
                        clearanceBottom
                    );
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ✅ DAMPER STRATEGY: Zone {zone.Id}, Offset=({damperAdj.offsetVector.X:F6}, {damperAdj.offsetVector.Y:F6}, {damperAdj.offsetVector.Z:F6}), Final W={damperAdj.finalWidth:F6}ft, H={damperAdj.finalHeight:F6}ft");
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
                double clearance = _clearanceService.GetClearance(zone, _conditions, _clearanceSettings, _strategy);
            
                double rawWidth = zone.MepElementWidth;
                double rawHeight = zone.MepElementHeight;
                double rawDiameter = zone.MepElementOuterDiameter > 0 ? zone.MepElementOuterDiameter : 0;
            
                // ✅ OOP METHOD: Use insulation-aware sizing service for consistent calculation (SOLID principles)
                // Formula: RawSize + (2 × InsulationThickness) + (2 × Clearance)
                (double finalWidth, double finalHeight, double finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZone(
                    rawWidth, rawHeight, rawDiameter, zone, clearance);
            
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

            // Determine level - use active view level as fallback
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
                
                // Place instance
                FamilyInstance instance = _doc.Create.NewFamilyInstance(point, symbol, level, StructuralType.NonStructural);
                
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

        private void SetSleeveParameters(FamilyInstance instance, double width, double height, double diameter, bool isCircular, ClashZone zone)
        {
            // ✅ PERFORMANCE MONITORING: Track parameter setting
            using (var tracker = _performanceMonitor?.TrackOperation("Set Sleeve Parameters"))
            {
                if (instance == null) return;

                // ✅ SAFE ELEMENT VALIDATION: Validate instance is still valid (avoids document mismatch bug)
                if (OptimizationFlags.UseSafeElementValidation)
                {
                    try
                    {
                        // ⚠️ CRITICAL: Do NOT compare documents by reference (causes false positives like in parameter transfer)
                        // Instead, validate by element ID - if doc.GetElement() succeeds, element is in correct document
                        var validationElement = _doc.GetElement(instance.Id);
                        if (validationElement == null || !validationElement.IsValidObject)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[NewSleevePlacer] ⚠️ Instance {instance.Id.IntegerValue} is invalid - skipping parameter setting");
                            }
                            return;
                        }

                        // Additional validation: ensure element IDs match (not just document reference)
                        if (validationElement.Id != instance.Id)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[NewSleevePlacer] ⚠️ Element ID mismatch for instance {instance.Id.IntegerValue}");
                            }
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[NewSleevePlacer] Error validating instance {instance.Id.IntegerValue}: {ex.Message}");
                        }
                        return;
                    }
                }

                var currentSleeveId = instance.Id;

                // ✅ GLOBAL SETTINGS: Apply rounding based on global configuration (RoundingValue and RoundAlwaysUp)
                // This matches the legacy UniversalSleevePlacerService behavior
                // Dampers are excluded from rounding (preserve exact calculated dimensions)
                bool isDamper = zone?.MepElementCategory != null && 
                               (zone.MepElementCategory.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                zone.MepElementCategory.IndexOf("Duct Accessories", StringComparison.OrdinalIgnoreCase) >= 0);
                
                double roundedWidth = width;
                double roundedHeight = height;
                double roundedDiameter = diameter;
                
                if (!isDamper)
                {
                    // ✅ Apply rounding for non-damper elements using global settings
                    // OpeningSettingsHelper reads RoundingValue and RoundAlwaysUp from ApplicationProfileService
                    (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                    roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(diameter);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // Log rounding if values changed
                        if (Math.Abs(width - roundedWidth) > 1e-6 || Math.Abs(height - roundedHeight) > 1e-6 || Math.Abs(diameter - roundedDiameter) > 1e-6)
                        {
                            DebugLogger.Info($"[NewSleevePlacer] [ROUNDING] Zone {zone?.Id}: " +
                                $"Width {RevitUnitConversionService.Instance.FromInternalMillimeters(width):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, " +
                                $"Height {RevitUnitConversionService.Instance.FromInternalMillimeters(height):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm, " +
                                $"Diameter {RevitUnitConversionService.Instance.FromInternalMillimeters(diameter):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedDiameter):F1}mm");
                        }
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] [ROUNDING] Zone {zone?.Id}: DAMPER - No rounding applied, using exact calculated dimensions");
                    }
                }

                if (isCircular)
                {
                    var param = instance.LookupParameter("Diameter") ?? instance.LookupParameter("Sleeve Diameter");
                    if (param != null && !param.IsReadOnly)
                    {
                        TimedSetDouble(param, roundedDiameter, "Diameter", currentSleeveId);
                    }
                }
                else
                {
                    var wParam = instance.LookupParameter("Width") ?? instance.LookupParameter("Sleeve Width");
                    var hParam = instance.LookupParameter("Height") ?? instance.LookupParameter("Sleeve Height");
                    if (wParam != null && !wParam.IsReadOnly)
                    {
                        TimedSetDouble(wParam, roundedWidth, "Width", currentSleeveId);
                    }
                    if (hParam != null && !hParam.IsReadOnly)
                    {
                        TimedSetDouble(hParam, roundedHeight, "Height", currentSleeveId);
                    }

                    tracker?.SetItemCount(1);
                }

                // ✅ SRP COMPLIANCE: Delegate depth parameter setting to dedicated method
                if (zone != null)
                {
                    SetDepthParameter(instance, zone, currentSleeveId);
                }

                // ✅ CLUSTERING SUPPORT: Set MEP_ElementId and MEP_Category for clustering and parameter transfer
                if (zone != null)
                {
                    // Set MEP_Category - Required for clustering
                    var mepCategoryParam = instance.LookupParameter("MEP_Category");
                    if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["MEP_Category"] = zone.MepElementCategory;
                        }
                        else
                        {
                            mepCategoryParam.Set(zone.MepElementCategory);
                        }
                    }

                    // ✅ CRITICAL: Set MEP_ElementId IMMEDIATELY (not deferred) - Required for clustering and corner retrieval
                    // Without MEP_ElementId, cluster sizing cannot find corners in database
                    var mepElementIdParam = instance.LookupParameter("MEP_ElementId");
                    if (mepElementIdParam != null && !mepElementIdParam.IsReadOnly)
                    {
                        // ✅ ALWAYS set immediately, regardless of batching flag - this is critical for clustering
                        mepElementIdParam.Set(zone.MepElementId.IntegerValue);
                        
                        // If batching is enabled, also add to deferred for consistency, but immediate set is critical
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["MEP_ElementId"] = zone.MepElementId.IntegerValue;
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] ✅ IMMEDIATE: Set 'MEP_ElementId'={zone.MepElementId.IntegerValue} for sleeve {currentSleeveId}\n");
                        }
                    }
                    else
                    {
                        // ⚠️ CRITICAL WARNING: MEP_ElementId parameter not found or read-only
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️⚠️⚠️ CRITICAL: Cannot set 'MEP_ElementId' for sleeve {currentSleeveId} - parameter not found or read-only!\n" +
                                $"  This will prevent cluster sizing from finding corners in database!\n");
                        }
                    }
                    
                    // ✅ DAMPER ASYMMETRIC CLEARANCE: Set individual clearance parameters from zone (populated by DamperPlacementStrategy)
                    // These parameters are used by the sleeve family to display asymmetric clearance visualization
                    // SOLID ISP: Each clearance parameter is independent, allowing selective setting
                    
                    // Set Clearance_Left (100mm on connector side, 50mm on other side, mapped by wall orientation)
                    var clearanceLeftParam = instance.LookupParameter("Clearance_Left");
                    if (clearanceLeftParam != null && !clearanceLeftParam.IsReadOnly && zone.ClearanceLeft > 0)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["Clearance_Left"] = zone.ClearanceLeft;
                        }
                        else
                        {
                            clearanceLeftParam.Set(zone.ClearanceLeft);
                        }
                    }
                    
                    // Set Clearance_Right
                    var clearanceRightParam = instance.LookupParameter("Clearance_Right");
                    if (clearanceRightParam != null && !clearanceRightParam.IsReadOnly && zone.ClearanceRight > 0)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["Clearance_Right"] = zone.ClearanceRight;
                        }
                        else
                        {
                            clearanceRightParam.Set(zone.ClearanceRight);
                        }
                    }
                    
                    // Set Clearance_Top
                    var clearanceTopParam = instance.LookupParameter("Clearance_Top");
                    if (clearanceTopParam != null && !clearanceTopParam.IsReadOnly && zone.ClearanceTop > 0)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["Clearance_Top"] = zone.ClearanceTop;
                        }
                        else
                        {
                            clearanceTopParam.Set(zone.ClearanceTop);
                        }
                    }
                    
                    // Set Clearance_Bottom
                    var clearanceBottomParam = instance.LookupParameter("Clearance_Bottom");
                    if (clearanceBottomParam != null && !clearanceBottomParam.IsReadOnly && zone.ClearanceBottom > 0)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(currentSleeveId))
                                _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                            _deferredParameters[currentSleeveId]["Clearance_Bottom"] = zone.ClearanceBottom;
                        }
                        else
                        {
                            clearanceBottomParam.Set(zone.ClearanceBottom);
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode && (zone.ClearanceLeft > 0 || zone.ClearanceRight > 0 || zone.ClearanceTop > 0 || zone.ClearanceBottom > 0))
                    {
                        DebugLogger.Info($"[NewSleevePlacer] ✅ DAMPER CLEARANCES SET: Zone {zone.Id}, L={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceLeft):F1}mm, R={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceRight):F1}mm, T={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceTop):F1}mm, B={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceBottom):F1}mm");
                    }
                }
            }
        }
        
        /// <summary>
        /// ✅ SRP COMPLIANCE: Dedicated method for setting Depth/Wall Width parameter based on host type.
        /// Single Responsibility: Calculate and set structural thickness parameter only.
        /// Maintains all optimization features: batching, performance monitoring, safe validation, diagnostic logging.
        /// </summary>
        private void SetDepthParameter(FamilyInstance instance, ClashZone zone, ElementId currentSleeveId)
        {
            if (instance == null || zone == null) return;
            
            bool isWallHost = zone.StructuralElementType == "Wall" || zone.StructuralElementType == "Walls";
            bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            // Get the correct thickness based on host type
            double thickness = 0.0;
            if (isWallHost)
            {
                thickness = zone.WallThickness > 0 ? zone.WallThickness : zone.StructuralElementThickness;
            }
            else if (isFramingHost)
            {
                thickness = zone.FramingThickness > 0 ? zone.FramingThickness : zone.StructuralElementThickness;
            }
            else
            {
                thickness = zone.StructuralElementThickness;
            }
            
            // ✅ CRITICAL FIX: For PATH 3 (Non-Fresh), if thickness is 0, retrieve from linked file structural element
            // This is a fallback - thickness should have been saved during refresh, but if missing, retrieve it now
            if (!_isReplayPath && thickness <= 0.0 && zone.StructuralElementIdValue > 0)
            {
                try
                {
                    // Try to get structural element from linked files
                    Element structuralElement = null;
                    var linkInstances = new FilteredElementCollector(_doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null)
                        {
                            try
                            {
                                structuralElement = linkDoc.GetElement(new ElementId(zone.StructuralElementIdValue));
                                if (structuralElement != null)
                                {
                                    // Retrieve thickness from linked file element
                                    if (structuralElement is Wall wall)
                                    {
                                        thickness = wall.Width;
                                    }
                                    else if (structuralElement is Floor floor)
                                    {
                                        thickness = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.0;
                                    }
                                    else if (structuralElement?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                                    {
                                        // Get framing thickness from type parameter 'b'
                                        var typeId = (structuralElement as FamilyInstance)?.GetTypeId() ?? structuralElement.GetTypeId();
                                        var typeElem = linkDoc.GetElement(typeId);
                                        if (typeElem != null)
                                        {
                                            var param = typeElem.LookupParameter("b") ?? typeElem.LookupParameter("B");
                                            if (param != null) thickness = param.AsDouble();
                                        }
                                    }
                                    
                                    if (thickness > 0.0)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[NewSleevePlacer] [DEPTH-FALLBACK] Retrieved thickness={thickness * 304.8:F1}mm from linked file for structural element {zone.StructuralElementIdValue} (Zone={zone.Id})");
                                        break;
                                    }
                                }
                            }
                            catch { continue; }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[NewSleevePlacer] [DEPTH-FALLBACK] Error retrieving thickness from linked file for Zone={zone.Id}: {ex.Message}");
                }
            }
            
            // Set Depth or Wall Width parameter
            bool depthSetSuccess = false;
            if (isWallHost)
            {
                var wallWidthParam = instance.LookupParameter("Wall Width");
                if (wallWidthParam != null && !wallWidthParam.IsReadOnly)
                {
                    // ✅ OPTIMIZATION: Uses TimedSetDouble which supports parameter batching
                    TimedSetDouble(wallWidthParam, thickness, "Wall Width", currentSleeveId);
                    depthSetSuccess = true;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] [DEPTH-SET] Zone={zone.Id}, Sleeve={instance.Id}: Set Wall Width={thickness * 304.8:F1}mm");
                    }
                }
            }
            
            if (!depthSetSuccess)
            {
                var depthParam = instance.LookupParameter("Depth");
                if (depthParam != null && !depthParam.IsReadOnly)
                {
                    // ✅ OPTIMIZATION: Uses TimedSetDouble which supports parameter batching
                    TimedSetDouble(depthParam, thickness, "Depth", currentSleeveId);
                    depthSetSuccess = true;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[NewSleevePlacer] [DEPTH-SET] Zone={zone.Id}, Sleeve={instance.Id}: Set Depth={thickness * 304.8:F1}mm");
                    }
                }
            }
            
            if (!depthSetSuccess && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[NewSleevePlacer] [DEPTH-SET] ❌ Zone={zone.Id}, Sleeve={instance.Id}: Could not set Depth or Wall Width parameter");
            }
        }
        
        /// <summary>
        /// ✅ PARAMETER BATCHING: Set parameter value with batching support and diagnostic logging.
        /// Deferred when batching enabled, immediate when disabled.
        /// </summary>
        private void TimedSetDouble(Parameter param, double value, string logicalName, ElementId currentSleeveId)
        {
            if (param == null || param.IsReadOnly) return;
            
            if (OptimizationFlags.UseBatchedParameterWrites)
            {
                if (!_deferredParameters.ContainsKey(currentSleeveId))
                    _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                
                // ✅ CRITICAL DIAGNOSTIC: Log if parameter is being overwritten
                bool isOverwrite = _deferredParameters[currentSleeveId].ContainsKey(logicalName);
                if (isOverwrite && (logicalName == "Width" || logicalName == "Height" || 
                                   logicalName == "Depth" || logicalName == "Wall Width"))
                {
                    var oldValue = _deferredParameters[currentSleeveId][logicalName];
                    SafeFileLogger.SafeAppendText("parameter_overwrite_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ⚠️ PARAMETER OVERWRITE: Sleeve {currentSleeveId.IntegerValue}, " +
                        $"Parameter='{logicalName}', OldValue={oldValue}, NewValue={value}\n");
                }
                
                _deferredParameters[currentSleeveId][logicalName] = value;
            }
            else
            {
                param.Set(value);
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL FIX: Read parameter from deferred cache first, then fallback to Revit element.
        /// This prevents stale reads during corner placement calculations when batching is enabled.
        /// </summary>
        private double GetParameterValueWithBatchingSupport(
            FamilyInstance sleeve, 
            string parameterName, 
            double fallbackValue)
        {
            if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null)
            {
                var sleeveId = sleeve.Id;
                if (_deferredParameters.ContainsKey(sleeveId) && 
                    _deferredParameters[sleeveId].ContainsKey(parameterName))
                {
                    var cachedValue = _deferredParameters[sleeveId][parameterName];
                    if (cachedValue is double doubleVal)
                    {
                        return doubleVal; // Return cached value
                    }
                }
            }
            
            // Fallback to Revit element
            var param = sleeve.LookupParameter(parameterName);
            return param?.AsDouble() ?? fallbackValue;
        }
        
        /// <summary>
        /// ✅ PARAMETER BATCHING: Flush all deferred parameters to Revit elements.
        /// Applies all accumulated parameter values after regeneration.
        /// </summary>
        public int FlushDeferredParameters()
        {
            // ✅ SAFETY FLAG: Prevent multiple flushes (critical for performance)
            if (_hasFlushedParameters)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var stackTrace = new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false);
                    var caller = stackTrace.GetFrame(0)?.GetMethod()?.Name ?? "Unknown";
                    DebugLogger.Warning($"[NewSleevePlacer] [BATCH-PARAMS] ⚠️ SAFETY: FlushDeferredParameters called AGAIN from {caller} - IGNORING (already flushed once). This indicates a bug - parameters should only flush once at the end!");
                }
                return 0; // ✅ CRITICAL: Exit early to prevent duplicate flushes
            }
            
            if (_deferredParameters == null || _deferredParameters.Count == 0)
            {
                _hasFlushedParameters = true; // Mark as flushed even if empty
                return 0;
            }
            
            int successCount = 0;
            int failCount = 0;
            var errorLog = new System.Text.StringBuilder();
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                DebugLogger.Info($"[NewSleevePlacer] [BATCH-PARAMS] 🔄 Flushing {_deferredParameters.Count} individual sleeves with {totalParams} total parameters...");
            }
            
            try
            {
                foreach (var kvp in _deferredParameters)
                {
                    var sleeveId = kvp.Key;
                    var paramValues = kvp.Value;
                    
                    var sleeve = _doc.GetElement(sleeveId) as FamilyInstance;
                    if (sleeve == null) continue;
                    
                    foreach (var paramKvp in paramValues)
                    {
                        var param = sleeve.LookupParameter(paramKvp.Key);
                        if (param != null && !param.IsReadOnly)
                        {
                            try
                            {
                                if (paramKvp.Value is double dVal)
                                    param.Set(dVal);
                                else if (paramKvp.Value is string sVal)
                                    param.Set(sVal);
                                
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                failCount++;
                                errorLog.AppendLine($"[NewSleevePlacer] Failed to set parameter '{paramKvp.Key}' on sleeve {sleeveId.IntegerValue}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[NewSleevePlacer] [BATCH-PARAMS] Error during flush: {ex.Message}");
                }
            }
            finally
            {
                _hasFlushedParameters = true;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                    DebugLogger.Info($"[NewSleevePlacer] [BATCH-PARAMS] ✅ Flushed {successCount} parameters for {_deferredParameters.Count} sleeves, {failCount} failed");
                    if (errorLog.Length > 0)
                    {
                        SafeFileLogger.SafeAppendText("parameter_batching_errors.log", errorLog.ToString());
                    }
                }
                
                // Clear deferred parameters after flush (ready for next placement batch)
                _deferredParameters.Clear();
            }
            
            return successCount;
        }
        
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

        private void UpdateSleeveDataInDatabase(ClashZone zone, FamilyInstance sleeve)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    
                    // Update Instance ID
                    repository.UpdateSleeveInstanceId(zone.Id, sleeve.Id.IntegerValue);
                    
                    // Update Placement Data
                    repository.UpdateSleevePlacement(
                        zone.Id,
                        sleeve.Id.IntegerValue,
                        zone.SleeveWidth,
                        zone.SleeveHeight,
                        zone.SleeveDiameter,
                        zone.SleevePlacementPointX,
                        zone.SleevePlacementPointY,
                        zone.SleevePlacementPointZ,
                        zone.SleevePlacementPointX, // Active doc coords (same for now)
                        zone.SleevePlacementPointY,
                        zone.SleevePlacementPointZ,
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
