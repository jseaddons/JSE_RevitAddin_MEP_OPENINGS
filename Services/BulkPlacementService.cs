using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class BulkPlacementResult
    {
        public bool OverallSuccess { get; set; }
        public int PlacedCount { get; set; }
        public int FailedCount { get; set; }
        // Name second item 'ElementId' to match existing call sites
        public List<(ClashZone Zone, ElementId ElementId)> PlacedItems { get; } = new();
        public List<(Guid ZoneId, string Reason)> Failures { get; } = new();
        public TimeSpan ElapsedTime { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// Performs bulk placement of individual sleeves from pre-planned data.
    /// This is the service used by OpeningCommandOrchestrator Operation 2.
    /// </summary>
    public class BulkPlacementService
    {
        private readonly Document _doc;
        private readonly Action<string> _logger;
        private readonly IPerformanceMonitor _performanceMonitor;
        private readonly SleeveParameterService _parameterService;
        private readonly SleeveRotationService _rotationService;
        private readonly Func<SleeveDbContext> _contextFactory;
        private readonly IParameterSnapshotTransferService _snapshotTransferService;

        public BulkPlacementService(
            Document doc,
            Func<SleeveDbContext> contextFactory = null,
            Action<string> logger = null,
            IPerformanceMonitor performanceMonitor = null,
            SleeveParameterService parameterService = null,
            IParameterSnapshotTransferService snapshotTransferService = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _contextFactory = contextFactory;
            _logger = logger ?? (_ => { });
            _performanceMonitor = performanceMonitor;
            // SleeveParameterService expects a PlacementPerformanceMonitor; cast when available
            var placementMonitor = performanceMonitor as PlacementPerformanceMonitor;
            _parameterService = parameterService ?? new SleeveParameterService(doc, isReplayPath: false, placementMonitor);
            _rotationService = new SleeveRotationService();
            _snapshotTransferService = snapshotTransferService;
        }

        /// <summary>
        /// ✅ NEW: Autonomous entry point for bulk placement.
        /// Loads zones from DB that match the current session context (Ready=1, Current=1),
        /// plans them, pre-saves data, and executes placement.
        /// </summary>
        public BulkPlacementResult ExecuteContextPlacement(Document doc, OpeningFilter filter = null, IOperationTracker parentTracker = null)
        {
            var result = new BulkPlacementResult();
            var start = DateTime.Now;

            try
            {
                _logger($"[CONTEXT-PLACEMENT] 🚀 Starting autonomous context placement for filter: {filter?.Name ?? "Unknown"}...");
                var ctxSw = System.Diagnostics.Stopwatch.StartNew();

                List<ClashZone> zones;
                using (parentTracker?.TrackSubOperation("1. Loading Zones from DB"))
                {
                    using (var dbContext = _contextFactory())
                    {
                        var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext, _logger, _performanceMonitor as PerformanceMonitor);
                        zones = repo.GetReadyZonesInContext();
                    }
                }
                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 1 (DB Load): {ctxSw.ElapsedMilliseconds}ms — {zones.Count} zones");

                if (zones.Count == 0)
                {
                    _logger("[CONTEXT-PLACEMENT] ℹ️ No zones found with ReadyForPlacement=1 and IsCurrentClash=1.");
                    result.OverallSuccess = true;
                    return result;
                }

                _logger($"[CONTEXT-PLACEMENT] 📂 Loaded {zones.Count} zones from database.");
                
                // ✅ CRASH SAFETY: Limit zones to prevent Revit hangs
                const int MAX_ZONES_PER_BATCH = 1000;
                if (zones.Count > MAX_ZONES_PER_BATCH)
                {
                    _logger($"[CONTEXT-PLACEMENT] ⚠️ Too many zones ({zones.Count}). Limiting to {MAX_ZONES_PER_BATCH} to prevent Revit hang.");
                    zones = zones.Take(MAX_ZONES_PER_BATCH).ToList();
                }

                // 1. Planning Phase
                var allBulkTaskItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                var planningTimer = System.Diagnostics.Stopwatch.StartNew();
                
                using (parentTracker?.TrackSubOperation("2. Planning & Dimensioning"))
                {
                    // ✅ PERF: Resolve levels once for the entire batch
                    var levelMap = new Dictionary<string, ElementId>(StringComparer.OrdinalIgnoreCase);
                    var levelNames = zones.Select(z => z.MepElementLevelName).Where(ln => !string.IsNullOrEmpty(ln)).Distinct();
                    if (levelNames.Any())
                    {
                        var levels = new FilteredElementCollector(_doc)
                            .OfClass(typeof(Level))
                            .Cast<Level>();
                        
                        foreach (var name in levelNames)
                        {
                            var level = levels.FirstOrDefault(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
                            if (level != null) levelMap[name] = level.Id;
                        }
                        _logger($"[CONTEXT-PLACEMENT] 🔋 Pre-cached {levelMap.Count} levels for placement optimization.");
                    }

                    var categoryGroups = zones.GroupBy(z => z.MepElementCategory);
                    var conditionsService = new ConditionsService(_doc);

                    foreach (var group in categoryGroups)
                    {
                        var category = group.Key;
                        var categoryZones = group.ToList();
                        // Load conditions using category name as filter key
                        var categoryConditions = conditionsService.LoadConditions(category);
                        
                        // ✅ PERF: Pass levelMap to planner
                        var planner = new ParallelSleevePlacementPlanner(categoryConditions, levelMap: levelMap);
                        var planningResult = planner.Plan(categoryZones);

                        _logger($"[CONTEXT-PLACEMENT] [PLANNING] Category '{category}': {planningResult.Items.Count} items planned in {planningResult.PlanningDurationMs}ms");

                        var plannedMap = planningResult.Items.ToDictionary(i => i.ClashZoneId);
                        foreach (var zone in categoryZones)
                        {
                            if (plannedMap.TryGetValue(zone.Id, out var plan))
                            {
                                if (plan.ShouldSkip) continue;

                                // Map calculated values for pre-saving
                                zone.SleeveWidth = plan.TargetWidthFt;
                                zone.SleeveHeight = plan.TargetHeightFt;
                                zone.SleeveDiameter = plan.TargetDiameterFt;
                                zone.SleeveDepth = plan.TargetDepthFt;
                                zone.SleeveFamilyName = plan.SleeveFamilyName;
                                zone.SleevePlacementPointX = plan.PlacementPoint.X;
                                zone.SleevePlacementPointY = plan.PlacementPoint.Y;
                                zone.SleevePlacementPointZ = plan.PlacementPoint.Z;
                                zone.MepElementRotationAngle = plan.RotationRadians;

                                // ✅ NEW: Map to Calculated* properties for DB persistence (BatchUpdateCalculatedData uses these)
                                zone.CalculatedSleeveWidth = plan.TargetWidthFt;
                                zone.CalculatedSleeveHeight = plan.TargetHeightFt;
                                zone.CalculatedSleeveDiameter = plan.TargetDiameterFt;
                                zone.CalculatedSleeveDepth = plan.TargetDepthFt;
                                zone.CalculatedPlacementX = plan.PlacementPoint.X;
                                zone.CalculatedPlacementY = plan.PlacementPoint.Y;
                                zone.CalculatedPlacementZ = plan.PlacementPoint.Z;
                                zone.CalculatedRotation = plan.RotationRadians;
                                zone.CalculatedFamilyName = plan.SleeveFamilyName;

                                // ✅ NEW: Pre-calculate conservative Bounding Box for clustering (Fixes "zero BBox" log and enables persistence)
                                double maxDim = Math.Max(plan.TargetWidthFt, Math.Max(plan.TargetHeightFt, Math.Max(plan.TargetDiameterFt, plan.TargetDepthFt)));
                                double halfMax = maxDim / 2.0;

                                zone.SleeveBoundingBoxMinX = plan.PlacementPoint.X - halfMax;
                                zone.SleeveBoundingBoxMinY = plan.PlacementPoint.Y - halfMax;
                                zone.SleeveBoundingBoxMinZ = plan.PlacementPoint.Z - halfMax;
                                zone.SleeveBoundingBoxMaxX = plan.PlacementPoint.X + halfMax;
                                zone.SleeveBoundingBoxMaxY = plan.PlacementPoint.Y + halfMax;
                                zone.SleeveBoundingBoxMaxZ = plan.PlacementPoint.Z + halfMax;

                                allBulkTaskItems.Add((zone, plan));
                            }
                        }
                    }
                }
                planningTimer.Stop();

                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 2 (Planning Total): {planningTimer.ElapsedMilliseconds}ms — {allBulkTaskItems.Count} items total");

                if (allBulkTaskItems.Count == 0)
                {
                    _logger("[CONTEXT-PLACEMENT] ℹ️ No zones remaining after planning/skipping.");
                    result.OverallSuccess = true;
                    return result;
                }

                // 2. Pre-Save Phase
                using (parentTracker?.TrackSubOperation("3. Pre-Saving Calculated Data"))
                {
                    using (var dbContext = _contextFactory())
                    {
                        var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext, _logger);
                        repo.BatchUpdateCalculatedData(allBulkTaskItems.Select(x => x.Zone).ToList());
                    }
                }
                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 3 (Pre-Save DB): {ctxSw.ElapsedMilliseconds}ms");

                // 3. Deduplication (Same location)
                List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> itemsToPlace;
                const double locationToleranceFt = 0.0328; // ~10mm
                Func<double, double> round = v => Math.Round(v / locationToleranceFt) * locationToleranceFt;
                
                itemsToPlace = allBulkTaskItems
                    .GroupBy(x =>
                    {
                        var pt = x.Plan.PlacementPoint;
                        return ($"{x.Plan.SleeveFamilyName}", round(pt.X), round(pt.Y), round(pt.Z));
                    })
                    .Select(g => g.First())
                    .ToList();

                if (itemsToPlace.Count < allBulkTaskItems.Count)
                {
                    _logger($"[CONTEXT-PLACEMENT] ✂️ Deduped by location: {allBulkTaskItems.Count} -> {itemsToPlace.Count}");
                }

                // 4. Revit Placement Phase
                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 4 START (Revit Placement): {ctxSw.ElapsedMilliseconds}ms");
                using (var t = new Transaction(doc, "Bulk Placement (Autonomous)"))
                {
                    var individualWarningHandler = new SwallowWarningsPreprocessor();
                    var failOpts = t.GetFailureHandlingOptions();
                    failOpts.SetFailuresPreprocessor(individualWarningHandler);
                    t.SetFailureHandlingOptions(failOpts);
                    t.Start();
                    
                    using (var tracker = parentTracker?.TrackSubOperation("4. Revit AI Placement"))
                    {
                        var placementResult = ExecuteBulkPlacement(doc, itemsToPlace, skipSpatialFiltering: false);
                        tracker?.SetItemCount(placementResult.PlacedCount);
                        
                        if (placementResult.OverallSuccess && placementResult.PlacedCount > 0)
                        {
                            // ✅ SNAPSHOT PARAMETER TRANSFER: Transfer MEP parameters from snapshots to placed sleeves
                            if (OptimizationFlags.EnableSnapshotParameterTransfer && _snapshotTransferService != null && _contextFactory != null)
                            {
                                try
                                {
                                    using (var dbContext = _contextFactory())
                                    {
                                        var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext, _logger);
                                        var sleeveElementIds = placementResult.PlacedItems.Select(p => p.ElementId).ToList();
                                        
                                        _logger($"[BULK-PLACEMENT] [SNAPSHOT-PARAMS] 🚀 Starting snapshot transfer for {sleeveElementIds.Count} placed sleeves");
                                        int deferredCount = _snapshotTransferService.TransferSnapshotParameters(doc, sleeveElementIds, repo, _parameterService);
                                        _logger($"[BULK-PLACEMENT] [SNAPSHOT-PARAMS] 📥 Deferred {deferredCount} parameters from snapshots");
                                    }
                                }
                                catch (Exception snapEx)
                                {
                                    _logger($"[BULK-PLACEMENT] [SNAPSHOT-PARAMS] ⚠️ Snapshot transfer failed: {snapEx.Message}");
                                }
                            }
                            
                            _parameterService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
                        }
                        
                        result.PlacedCount = placementResult.PlacedCount;
                        result.FailedCount = placementResult.FailedCount;
                        result.PlacedItems.AddRange(placementResult.PlacedItems);
                        result.Failures.AddRange(placementResult.Failures);
                        result.OverallSuccess = placementResult.OverallSuccess;
                    }
                    
                    using (_performanceMonitor?.TrackOperation("Transaction Commit (Individual)"))
                    {
                        t.Commit();
                    }
                    _performanceMonitor?.LogMetric("INDIVIDUAL TX WARNINGS",
                        $"Deleted={individualWarningHandler.WarningsDeleted}, Errors={individualWarningHandler.ErrorsFound}");
                }
                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 4 END (Revit Placement): {ctxSw.ElapsedMilliseconds}ms — placed {result.PlacedCount}");

                // 5. Post-Placement Update & DB Persist
                if (result.PlacedCount > 0)
                {
                    using (parentTracker?.TrackSubOperation("5. Updating Database Flags"))
                    {
                        UpdateZonesFromElements(doc, result.PlacedItems);

                        // ✅ STEP-BY-STEP DIAGNOSTIC LOG
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"\n[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════════════════════\n");
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}] STEP 1: INDIVIDUAL PLACEMENT COMPLETED\n");
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}]   Placed {result.PlacedCount} individual sleeves\n");

                        using (var dbContext = _contextFactory())
                        {
                            var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext, _logger);

                            // Log before update
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 2: UPDATING DATABASE FLAGS (BULK)\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Setting IsResolvedFlag=1 for {result.PlacedCount} zones\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   RESETTING MarkedForClusterProcess=FALSE (proximity check will re-evaluate)\n");

                            // ✅ PERFORMANCE OPTIMIZATION: Use high-performance bulk update for post-placement data
                            var zonesToUpdate = result.PlacedItems.Select(p => p.Zone).ToList();
                            foreach (var zone in zonesToUpdate)
                            {
                                zone.IsResolved = true;
                                zone.PlacementStatus = "Placed";
                                zone.IsClusteredFlag = false; // Individual placement resets this
                            }

                            var bulkUpdateTimer = System.Diagnostics.Stopwatch.StartNew();
                            repo.BatchUpdatePostPlacement(zonesToUpdate);
                            bulkUpdateTimer.Stop();

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   ✅ BatchUpdatePostPlacement (TEMP TABLE) completed in {bulkUpdateTimer.ElapsedMilliseconds}ms\n");

                            // DIAGNOSTIC: Verify that IsResolvedFlag is actually set in database
                            var verifyZones = repo.GetZonesReadyForProximityCheck();
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 3: VERIFICATION\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Zones ready for proximity check: {verifyZones.Count}\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   (Query: IsResolvedFlag=1 AND MarkedForClusterProcess IS NULL)\n");

                            if (verifyZones.Count > 0)
                            {
                                var sample = verifyZones.First();
                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   Sample zone: {sample.Id.ToString().Substring(0, 8)}...\n");
                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     IsResolvedFlag={sample.IsResolved}\n");
                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     MarkedForClusterProcess={sample.MarkedForClusterProcess?.ToString() ?? "NULL"}\n");
                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]     BBox: ({sample.SleeveBoundingBoxMinX:F3},{sample.SleeveBoundingBoxMinY:F3},{sample.SleeveBoundingBoxMinZ:F3})\n");

                                _logger($"[BULK-PLACEMENT] [DIAGNOSTIC] ✅ IsResolvedFlag correctly set to 1 for placed sleeves");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   ⚠️ WARNING: No zones ready for proximity check!\n");
                            }

                            _logger($"[BULK-PLACEMENT] [DIAGNOSTIC] Zones ready for clustering after flag update: {verifyZones.Count}");
                        }
                    }
                }


                _logger($"[CONTEXT-PLACEMENT] ⏱️ STEP 5 END (Post-DB): {ctxSw.ElapsedMilliseconds}ms");
                _logger($"[CONTEXT-PLACEMENT] ✅ Execution complete in {ctxSw.ElapsedMilliseconds}ms. Placed: {result.PlacedCount}, Failed: {result.FailedCount}");
            }

            catch (Exception ex)
            {
                result.OverallSuccess = false;
                result.Error = ex.Message;
                _logger($"[CONTEXT-PLACEMENT] ❌ CRITICAL ERROR: {ex.Message}\n{ex.StackTrace}");
                SafeFileLogger.SafeAppendText("placement_error.log", $"[{DateTime.Now:HH:mm:ss}] ExecuteContextPlacement Failed: {ex.Message}\n{ex.StackTrace}\n");
            }
            finally
            {
                result.ElapsedTime = DateTime.Now - start;
            }

            return result;
        }

        /// <summary>
        /// Execute bulk placement for the given planned sleeves.
        /// This method must be called inside an open Revit transaction.
        /// </summary>
        public BulkPlacementResult ExecuteBulkPlacement(
            Document doc,
            List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> items,
            bool skipSpatialFiltering = false)
        {
            var result = new BulkPlacementResult();
            var start = DateTime.Now;
            
            // ---------------------------------------------------------
            // 1. SESSION CONTEXT FILTERING (Section Box + File Combos)
            // ---------------------------------------------------------
            BoundingBoxXYZ sectionBox = null;
            if (_contextFactory != null)
            {
                try 
                {
                    using (var ctx = _contextFactory())
                    {
                        // A. Get Section Box from DB
                        var sbService = new JSE_RevitAddin_MEP_OPENINGS.Services.SectionBoxService();
                        sectionBox = sbService.GetSectionBoxBounds(ctx.Connection);
                        
                        // Log section box status
                        if (sectionBox != null && !Services.OptimizationFlags.DisableVerboseLogging)
                        {
                            _logger($"[SESSION-CONTEXT] 📦 Filtering by active Section Box: Min({sectionBox.Min}), Max({sectionBox.Max})");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger($"[SESSION-CONTEXT] ⚠️ Error retrieving session context: {ex.Message}");
                }
            }
            
            // B. Apply Filtering (Spatial + Context Flag)
            var filteredItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
            int skippedByBox = 0;
            int skippedByContext = 0;
            
            foreach (var item in items)
            {
                var zone = item.Zone;
                var pt = item.Plan.PlacementPoint;
                
                // Check 1: Session Context Flag (File/Filter/Category)
                // IsCurrentClashFlag is set by Refresh command based on active filters
                // If flag is FALSE, this zone is not part of the current user session context
                if (!zone.IsCurrentClash)
                {
                    skippedByContext++;
                    continue;
                }

                // Check 2: 3D Section Box (Spatial)
                if (sectionBox != null && !skipSpatialFiltering)
                {
                    // Strict containment check
                    if (pt.X < sectionBox.Min.X || pt.X > sectionBox.Max.X ||
                        pt.Y < sectionBox.Min.Y || pt.Y > sectionBox.Max.Y ||
                        pt.Z < sectionBox.Min.Z || pt.Z > sectionBox.Max.Z)
                    {
                        skippedByBox++;
                        continue;
                    }
                }
                
                filteredItems.Add(item);
            }
            
            if (skippedByBox > 0 || skippedByContext > 0)
            {
                 _logger($"[SESSION-CONTEXT] 🛑 Skipped {items.Count - filteredItems.Count} items out of {items.Count}: {skippedByContext} by Context Flag, {skippedByBox} by Section Box (skipSpatialFiltering={skipSpatialFiltering})");
                 
                 if (sectionBox != null && items.Count > 0 && filteredItems.Count == 0 && !skipSpatialFiltering)
                 {
                     _logger($"[SESSION-CONTEXT] ⚠️ ALL items were filtered out by Section Box. Current Box: Min({sectionBox.Min}), Max({sectionBox.Max}). First item pt: {items[0].Plan.PlacementPoint}");
                 }
            }
            else
            {
                 if (!Services.OptimizationFlags.DisableVerboseLogging)
                    _logger($"[SESSION-CONTEXT] ✅ All {items.Count} items passed context checks");
            }
             
            // Proceed with filtered list
            items = filteredItems;


            if (items == null || items.Count == 0)
            {
                result.OverallSuccess = true;
                return result;
            }

            try
            {
                // Step 1: Pre-activate symbols
                var symbolCache = PreActivateSymbols(doc, items.Select(x => x.Plan).ToList());
                
                // ✅ CRASH SAFETY: Limit batch size to prevent Revit hangs
                const int MAX_BATCH_SIZE = 500;
                if (items.Count > MAX_BATCH_SIZE)
                {
                    _logger($"[BULK-PLACEMENT] ⚠️ Large placement detected: {items.Count} items. Limiting to {MAX_BATCH_SIZE} per batch.");
                    items = items.Take(MAX_BATCH_SIZE).ToList();
                }
                
                // ✅ CRASH SAFETY: Initialize timeout tracking
                var placementStartTime = DateTime.Now;
                var maxPlacementDuration = TimeSpan.FromMinutes(5); // 5 minute limit for placement

                // Step 2: Create instances using BATCH API (NewFamilyInstances2)
                var idList = new List<ElementId>();
                var itemMap = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                var elementCache = new Dictionary<ElementId, FamilyInstance>();

                // ✅ Build FamilyInstanceCreationData list for batch placement
                var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
                var placementDataMap = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan, XYZ Location, string FamilyName)>();

                int placementIndex = 0;
                foreach (var (zone, plan) in items)
                {
                    placementIndex++;
                    try
                    {
                        string famName = plan.SleeveFamilyName ?? zone.SleeveFamilyName;
                        if (string.IsNullOrEmpty(famName) || !symbolCache.TryGetValue(famName, out var symbol))
                        {
                            result.Failures.Add((zone.Id, $"Family symbol '{famName}' not found"));
                            result.FailedCount++;
                            continue;
                        }

                        XYZ location = plan.PlacementPoint ??
                                       new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

                        // ✅ SANITY CHECK: Detect wild outliers
                        if (Math.Abs(location.X) > 3000 || Math.Abs(location.Y) > 3000)
                        {
                            _logger($"[BulkPlacement] ⚠️ WARNING: Placing sleeve at extreme coordinates ({location.X:F1}, {location.Y:F1}, {location.Z:F1}) for Zone {zone.Id}. This may be an outlier.");
                        }

                        // Create FamilyInstanceCreationData for batch placement
                        var creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(
                            location,
                            symbol,
                            StructuralType.NonStructural);

                        creationDataList.Add(creationData);
                        placementDataMap.Add((zone, plan, location, famName));
                    }
                    catch (Exception ex)
                    {
                        result.Failures.Add((zone.Id, $"Preparation failed: {ex.Message}"));
                        result.FailedCount++;
                    }
                }

                // ✅ CRASH SAFETY: Check timeout before batch placement
                if (DateTime.Now - placementStartTime > maxPlacementDuration)
                {
                    _logger($"[BULK-PLACEMENT] ⏱️ TIMEOUT: Placement exceeded 5 minutes. Stopping after preparation.");
                    result.OverallSuccess = false;
                    result.Error = "Placement timeout - exceeded 5 minute limit";
                    return result;
                }
                
                // ✅ CRASH SAFETY: Yield to Revit before heavy operation
                System.Windows.Forms.Application.DoEvents();

                // ✅ BATCH PLACEMENT using NewFamilyInstances2
                ICollection<ElementId> placedIds;
                using (var op = _performanceMonitor?.TrackOperation("Revit NewFamilyInstances2 (BATCH)"))
                {
                    var swBatch = System.Diagnostics.Stopwatch.StartNew();
                    placedIds = doc.Create.NewFamilyInstances2(creationDataList);
                    swBatch.Stop();
                    op?.SetItemCount(placedIds.Count);
                    _logger($"[BULK-PLACEMENT] ⏱️ STEP 2 (Revit API: NewFamilyInstances2): {swBatch.ElapsedMilliseconds}ms for {creationDataList.Count} items");
                }

                // ✅ PERF: Separate Regeneration from Placement call to isolate timings
                // Regeneration is needed so placement points refresh (avoiding 0.0 coordinates)
                using (var parentTracker = _performanceMonitor?.TrackOperation("Revit API: Global Refresh")) // Assuming _performanceMonitor is available and parentTracker is a suitable name
                {
                    var swRegen = System.Diagnostics.Stopwatch.StartNew();
                    doc.Regenerate();
                    swRegen.Stop();
                    _logger($"[BULK-PLACEMENT] ⏱️ STEP 2.5 (Revit API: Global Refresh): {swRegen.ElapsedMilliseconds}ms");
                }

                // ✅ CRITICAL: Verify 1-to-1 mapping
                var placedIdList = placedIds.ToList();
                if (placedIdList.Count != creationDataList.Count)
                {
                    var errorMsg = $"CRITICAL: Batch placement mismatch! Requested {creationDataList.Count} but got {placedIdList.Count} IDs";
                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PLACE] ❌ {errorMsg}\n");
                    throw new InvalidOperationException(errorMsg);
                }

                // ✅ FORCE locations and populate caches
                for (int i = 0; i < placedIdList.Count; i++)
                {
                    // ✅ CRASH SAFETY: Check timeout during post-processing
                    if (DateTime.Now - placementStartTime > maxPlacementDuration)
                    {
                        _logger($"[BULK-PLACEMENT] ⏱️ TIMEOUT: Post-processing exceeded 5 minutes. Stopping at {i}/{placedIdList.Count}.");
                        break;
                    }
                    
                    // ✅ CRASH SAFETY: Yield every 200 items to prevent UI freeze (was 50 — too frequent)
                    if (i % 200 == 0 && i > 0)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    
                    var elementId = placedIdList[i];
                    var (zone, plan, expectedLocation, famName) = placementDataMap[i];

                    var inst = doc.GetElement(elementId) as FamilyInstance;
                    if (inst == null)
                    {
                        result.Failures.Add((zone.Id, "Placed element not found"));
                        result.FailedCount++;
                        continue;
                    }

                    idList.Add(elementId);
                    itemMap.Add((zone, plan));
                    elementCache[elementId] = inst;
                }

                // Step 3: Apply rotation & parameters (deferred via SleeveParameterService batching)
                using (var tracker = _performanceMonitor?.TrackOperation("Apply Rotation & Parameters"))
                {
                    // PERF DIAGNOSTIC: Sub-operation timing (tick-precision to catch sub-ms operations)
                    long rotationTotalTicks = 0, paramsTotalTicks = 0;
                    int rotationCount = 0, preRotatedCount = 0;
                    var subSw = new System.Diagnostics.Stopwatch();
                    var loopSw = System.Diagnostics.Stopwatch.StartNew();
                    double ticksPerMs = System.Diagnostics.Stopwatch.Frequency / 1000.0;

                    for (int i = 0; i < idList.Count; i++)
                    {
                        var elementId = idList[i];
                        var (zone, plan) = itemMap[i];

                        if (!elementCache.TryGetValue(elementId, out var instance) || instance == null)
                        {
                            result.Failures.Add((zone.Id, "Element not found in cache"));
                            result.FailedCount++;
                            continue;
                        }

                        // ✅ PERF: Check if family is pre-rotated (_X variant) — skip rotation entirely
                        string plannedFamilyName = plan.SleeveFamilyName ?? zone.SleeveFamilyName ?? "";
                        bool isPreRotatedFamily = plannedFamilyName.EndsWith("_X", StringComparison.OrdinalIgnoreCase);

                        subSw.Restart();
                        if (!isPreRotatedFamily)
                        {
                            bool isWallOrFraming = IsWallOrFraming(zone);

                            double rotationRad = isWallOrFraming
                                ? _rotationService.DetermineRotation(zone)
                                : plan.RotationRadians;

                            bool isXWall = string.Equals(zone.HostOrientation, "X", StringComparison.OrdinalIgnoreCase);
                            bool isFloor = string.Equals(zone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase);

                            if (Math.Abs(rotationRad) > 0.001 && (isXWall || isFloor))
                            {
                                XYZ expectedLocation = plan.PlacementPoint ??
                                    new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

                                ApplyRotation(doc, instance, rotationRad, expectedLocation);
                                rotationCount++;
                            }
                        }
                        rotationTotalTicks += subSw.ElapsedTicks;
                        if (isPreRotatedFamily) preRotatedCount++;

                        subSw.Restart();
                        _parameterService.SetSleeveParameters(
                            instance,
                            plan.TargetWidthFt,
                            plan.TargetHeightFt,
                            plan.TargetDiameterFt,
                            plan.IsCircular,
                            zone,
                            plan.RequiredDepthFt,
                            skipValidation: true);
                        paramsTotalTicks += subSw.ElapsedTicks;

                        zone.SleeveInstanceId = elementId.GetIntegerValue();

                        result.PlacedItems.Add((zone, elementId));
                        result.PlacedCount++;
                    }

                    loopSw.Stop();
                    double totalLoopMs = loopSw.Elapsed.TotalMilliseconds;
                    double rotMs = rotationTotalTicks / ticksPerMs;
                    double paramMs = paramsTotalTicks / ticksPerMs;
                    double unaccountedMs = Math.Max(0, totalLoopMs - rotMs - paramMs);

                    // ✅ RECORD SUB-OPERATIONS: This makes them appear in the performance report table
                    if (tracker != null)
                    {
                        tracker.RecordSubOperation("Sub: Rotation", (long)rotMs, 0, rotationCount);
                        tracker.RecordSubOperation("Sub: SetParams", (long)paramMs, 0, idList.Count);
                        tracker.RecordSubOperation("Sub: Loop Overhead", (long)unaccountedMs, 0, 0);
                        tracker.SetItemCount(idList.Count);
                    }

                    _logger($"[BulkPlacement] SUB-TIMING: Total={totalLoopMs:F1}ms, Rotation={rotMs:F1}ms ({rotationCount} rotated), SetParams={paramMs:F1}ms, Unaccounted~{unaccountedMs:F0}ms");
                    
                    // ✅ ALWAYS WRITE to performance log (DeploymentMode should not suppress this specific breakdown)
                    SafeFileLogger.SafeAppendTextAlways("placement_performance.log",
                        $"\n[{DateTime.Now:HH:mm:ss}] SUB-TIMING BREAKDOWN ({idList.Count} sleeves):\n" +
                        $"  Total Loop:     {totalLoopMs:F1}ms\n" +
                        $"  Rotation:       {rotMs:F1}ms ({rotationCount} actually rotated, {preRotatedCount} _X pre-rotated, {idList.Count - rotationCount - preRotatedCount} Y-wall/no-rotation)\n" +
                        $"  SetParams:      {paramMs:F1}ms ({paramMs / idList.Count:F2}ms per sleeve)\n" +
                        $"  Unaccounted:    ~{unaccountedMs:F0}ms (loop overhead, cache lookups)\n");
                }

                _logger($"[BulkPlacement] ✅ Set parameters for all {result.PlacedCount} instances");
                result.OverallSuccess = true;
            }
            catch (Exception ex)
            {
                result.OverallSuccess = false;
                result.Error = ex.Message;
                _logger($"[BulkPlacement] CRITICAL ERROR: {ex.Message}");
            }
            finally
            {
                result.ElapsedTime = DateTime.Now - start;
            }

            return result;
        }

        private Dictionary<string, FamilySymbol> PreActivateSymbols(Document doc, List<SleevePlacementPlanningDto> plans)
        {
            var cache = new Dictionary<string, FamilySymbol>(StringComparer.OrdinalIgnoreCase);
            var familyNames = plans
                .Select(p => p.SleeveFamilyName)
                .Where(f => !string.IsNullOrEmpty(f))
                .Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (var familyName in familyNames)
            {
                var symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.Family.Name == familyName || s.Name == familyName);

                if (symbol != null)
                {
                    if (!symbol.IsActive) symbol.Activate();
                    cache[familyName] = symbol;
                }
            }

            _logger($"[BulkPlacement] Activated {cache.Count} unique symbols");
            return cache;
        }

        private static bool IsWallOrFraming(ClashZone zone)
        {
            if (zone?.StructuralElementType == null) return false;
            var t = zone.StructuralElementType.Trim();
            return t.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ApplyRotation(Document doc, FamilyInstance instance, double rotationRad, XYZ expectedPlacementPoint)
        {
            if (Math.Abs(rotationRad) < 0.001) return;

            try
            {
                var locPt = instance.Location as LocationPoint;
                XYZ axisPoint1 = locPt?.Point;

                if (axisPoint1 == null) return;

                // ✅ CRITICAL FIX: If axis point is (0,0,0), MOVE the instance to expected location!
                bool isAtOrigin = Math.Abs(axisPoint1.X) < 0.0001 && Math.Abs(axisPoint1.Y) < 0.0001 && Math.Abs(axisPoint1.Z) < 0.0001;
                if (isAtOrigin)
                {
                    XYZ moveVector = expectedPlacementPoint - axisPoint1;
                    Autodesk.Revit.DB.ElementTransformUtils.MoveElement(doc, instance.Id, moveVector);

                    // Re-read after move
                    axisPoint1 = (instance.Location as LocationPoint)?.Point;
                    if (axisPoint1 == null || (Math.Abs(axisPoint1.X) < 0.0001 && Math.Abs(axisPoint1.Y) < 0.0001))
                    {
                        locPt.Point = expectedPlacementPoint;
                        axisPoint1 = expectedPlacementPoint;
                    }
                }

                XYZ axisPoint2 = axisPoint1 + XYZ.BasisZ;
                Line axis = Line.CreateBound(axisPoint1, axisPoint2);
                instance.Location.Rotate(axis, rotationRad);
            }
            catch (Exception ex)
            {
                _logger($"[BulkPlacement] Rotation failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates ClashZone objects with actual placement point from placed elements.
        /// ✅ PERF: Only reads Location from Revit (1 API call per sleeve).
        /// BoundingBox and dimensions are already set during planning (Step 2) — no need to re-read.
        /// Previously did ~12 Revit API calls per sleeve (GetElement + get_BoundingBox + 6x LookupParameter + FamilyName).
        /// </summary>
        public void UpdateZonesFromElements(Document doc, List<(ClashZone Zone, ElementId ElementId)> items)
        {
            if (items == null || items.Count == 0) return;

            foreach (var (zone, ElementId) in items)
            {
                try
                {
                    var element = doc.GetElement(ElementId) as FamilyInstance;
                    if (element == null) continue;

                    // ✅ Only read-back actual placement point (Revit may have adjusted during commit)
                    var locPt = element.Location as LocationPoint;
                    if (locPt != null)
                    {
                        zone.SleevePlacementPointX = locPt.Point.X;
                        zone.SleevePlacementPointY = locPt.Point.Y;
                        zone.SleevePlacementPointZ = locPt.Point.Z;

                        zone.SleevePlacementActiveX = locPt.Point.X;
                        zone.SleevePlacementActiveY = locPt.Point.Y;
                        zone.SleevePlacementActiveZ = locPt.Point.Z;

                        zone.SleeveInstanceId = element.Id.Value;
                        zone.PlacementStatus = "Placed";
                        zone.IsResolvedFlag = true;
                        zone.IsClusteredFlag = false;
                        zone.SleeveState = Models.SleeveStateType.IndividualPlaced;
                    }

                    // ✅ PERF: BoundingBox, Width, Height, Diameter, Depth, FamilyName
                    // are already set on the zone from planning (Step 2, lines 127-146).
                    // No need to re-read from Revit — saves ~10 API calls per sleeve.
                }
                catch (Exception ex)
                {
                    _logger($"[BulkPlacement] Failed to update zone {zone?.Id} from element: {ex.Message}");
                }
            }
        }
    }
}