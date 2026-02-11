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
        private readonly IPerformanceMonitor? _performanceMonitor;
        private readonly SleeveParameterService _parameterService;
        private readonly SleeveRotationService _rotationService;
        private readonly Func<SleeveDbContext> _contextFactory;

        public BulkPlacementService(
            Document doc,
            Func<SleeveDbContext> contextFactory = null,
            Action<string>? logger = null,
            IPerformanceMonitor? performanceMonitor = null,
            SleeveParameterService? parameterService = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _contextFactory = contextFactory;
            _logger = logger ?? (_ => { });
            _performanceMonitor = performanceMonitor;
            // SleeveParameterService expects a PlacementPerformanceMonitor; cast when available
            var placementMonitor = performanceMonitor as PlacementPerformanceMonitor;
            _parameterService = parameterService ?? new SleeveParameterService(doc, isReplayPath: false, placementMonitor);
            _rotationService = new SleeveRotationService();
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

                List<ClashZone> zones;
                using (parentTracker?.TrackSubOperation("1. Loading Zones from DB"))
                {
                    using (var dbContext = _contextFactory())
                    {
                        var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext, _logger, _performanceMonitor as PerformanceMonitor);
                        zones = repo.GetReadyZonesInContext();
                    }
                }

                if (zones.Count == 0)
                {
                    _logger("[CONTEXT-PLACEMENT] ℹ️ No zones found with ReadyForPlacement=1 and IsCurrentClash=1.");
                    result.OverallSuccess = true;
                    return result;
                }

                _logger($"[CONTEXT-PLACEMENT] 📂 Loaded {zones.Count} zones from database.");

                // 1. Planning Phase
                var allBulkTaskItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                
                using (parentTracker?.TrackSubOperation("2. Planning & Dimensioning"))
                {
                    var categoryGroups = zones.GroupBy(z => z.MepElementCategory);
                    var conditionsService = new ConditionsService(_doc);

                    foreach (var group in categoryGroups)
                    {
                        var category = group.Key;
                        var categoryZones = group.ToList();
                        // Load conditions using category name as filter key
                        var categoryConditions = conditionsService.LoadConditions(category);
                        
                        var planner = new ParallelSleevePlacementPlanner(categoryConditions);
                        var planningResult = planner.Plan(categoryZones);

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
                        _logger($"[CONTEXT-PLACEMENT] ✅ Pre-saved {allBulkTaskItems.Count} zones to DB.");
                    }
                }

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
                using (var t = new Transaction(doc, "Bulk Placement (Autonomous)"))
                {
                    t.Start();
                    
                    using (parentTracker?.TrackSubOperation("4. Revit AI Placement"))
                    {
                        var placementResult = ExecuteBulkPlacement(doc, itemsToPlace);
                        
                        if (placementResult.OverallSuccess && placementResult.PlacedCount > 0)
                        {
                            _parameterService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
                        }
                        
                        result.PlacedCount = placementResult.PlacedCount;
                        result.FailedCount = placementResult.FailedCount;
                        result.PlacedItems.AddRange(placementResult.PlacedItems);
                        result.Failures.AddRange(placementResult.Failures);
                        result.OverallSuccess = placementResult.OverallSuccess;
                    }
                    
                    t.Commit();
                }

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
                                $"[{DateTime.Now:HH:mm:ss}] STEP 2: UPDATING DATABASE FLAGS\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Setting IsResolvedFlag=1 for {result.PlacedCount} zones\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   RESETTING MarkedForClusterProcess=FALSE (proximity check will re-evaluate)\n");

                            // Sample GUIDs being updated
                            var sampleGuids = string.Join(", ", result.PlacedItems.Take(3).Select(p => p.Zone.Id.ToString().Substring(0, 8)));
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Sample GUIDs: {sampleGuids}...\n");

                            var updates = result.PlacedItems.Select(p =>
                            (
                                p.Zone.Id,
                                true, // IsResolved
                                (bool?)false, // IsClusterResolved - Reset for individual placement
                                (bool?)null, // IsCombinedResolved (Preserve)
                                p.ElementId.IntegerValue,
                                -1, // ClusterID
                                (bool?)false, // IsClusteredFlag - Reset for individual placement
                                (bool?)false, // MarkedForCluster - RESET to false (will be re-evaluated by proximity check)
                                -1, // AfterClusterID
                                false, // IsClustered (deprecated field logic)
                                p.Zone.SleeveWidth,
                                p.Zone.SleeveHeight,
                                p.Zone.SleeveDiameter,
                                p.Zone.SleeveDepth,
                                p.Zone.SleeveFamilyName,
                                (double?)p.Zone.SleevePlacementActiveX, // ActivePlacementX
                                (double?)p.Zone.SleevePlacementActiveY, // ActivePlacementY
                                (double?)p.Zone.SleevePlacementActiveZ, // ActivePlacementZ
                                (double?)p.Zone.SleeveBoundingBoxMinX,
                                (double?)p.Zone.SleeveBoundingBoxMinY,
                                (double?)p.Zone.SleeveBoundingBoxMinZ,
                                (double?)p.Zone.SleeveBoundingBoxMaxX,
                                (double?)p.Zone.SleeveBoundingBoxMaxY,
                                (double?)p.Zone.SleeveBoundingBoxMaxZ
                            )).ToList();

                            repo.BatchUpdateFlags(new List<(Guid ClashZoneId, bool IsResolved, bool? IsClusterResolved, bool? IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool? IsClusteredFlag, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool IsClustered, double SleeveWidth, double SleeveHeight, double SleeveDiameter, double SleeveDepth, string SleeveFamilyName, double? ActivePlacementX, double? ActivePlacementY, double? ActivePlacementZ, double? BBoxMinX, double? BBoxMinY, double? BBoxMinZ, double? BBoxMaxX, double? BBoxMaxY, double? BBoxMaxZ)>(updates));

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   ✅ BatchUpdateFlags completed successfully\n");

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

                _logger($"[CONTEXT-PLACEMENT] ✅ Execution complete. Placed: {result.PlacedCount}, Failed: {result.FailedCount}");
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
            List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> items)
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
                if (sectionBox != null)
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
                 _logger($"[SESSION-CONTEXT] 🛑 Skipped {items.Count - filteredItems.Count} items: {skippedByContext} by Context Flag, {skippedByBox} by Section Box");
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

                        // ⚡ LOG BEFORE PLACEMENT
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PREP] [{placementIndex}/{items.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                            $"Location=({location.X:F6}, {location.Y:F6}, {location.Z:F6}), Family={famName}\n");

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

                        // ⚡ LOG: What location did we pass to FamilyInstanceCreationData?
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [CREATION-DATA] [{placementIndex}/{items.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                            $"Passed Location=({location.X:F6}, {location.Y:F6}, {location.Z:F6}) to FamilyInstanceCreationData\n");

                        creationDataList.Add(creationData);
                        placementDataMap.Add((zone, plan, location, famName));
                    }
                    catch (Exception ex)
                    {
                        result.Failures.Add((zone.Id, $"Preparation failed: {ex.Message}"));
                        result.FailedCount++;
                    }
                }

                // ✅ BATCH PLACEMENT using NewFamilyInstances2
                ICollection<ElementId> placedIds;
                using (_performanceMonitor?.TrackOperation("Revit NewFamilyInstances2 (BATCH)"))
                {
                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PLACE] Calling NewFamilyInstances2 with {creationDataList.Count} items\n");

                    placedIds = doc.Create.NewFamilyInstances2(creationDataList);
                    
                    // ✅ CRITICAL: Force Revit to calculate the location of newly batched instances
                    // Without this, the Location property may return (0,0,0) immediately after batch creation,
                    // which causes our "Origin Safety Check" in ApplyRotation to double-move the element.
                    doc.Regenerate();

                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PLACE] NewFamilyInstances2 returned {placedIds.Count} IDs and Regenerated\n");
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
                    var elementId = placedIdList[i];
                    var (zone, plan, expectedLocation, famName) = placementDataMap[i];

                    var inst = doc.GetElement(elementId) as FamilyInstance;
                    if (inst == null)
                    {
                        result.Failures.Add((zone.Id, "Placed element not found"));
                        result.FailedCount++;
                        continue;
                    }

                    // ⚡ LOG ACTUAL LOCATION AFTER BATCH PLACEMENT (don't force yet - Revit doubles it!)
                    var actualLoc = (inst.Location as LocationPoint)?.Point;
                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-VERIFY] [{i+1}/{placedIdList.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                        $"Expected=({expectedLocation.X:F6}, {expectedLocation.Y:F6}, {expectedLocation.Z:F6}), " +
                        $"Actual=({actualLoc?.X:F6}, {actualLoc?.Y:F6}, {actualLoc?.Z:F6})\n");

                    // ⚠️ DON'T FORCE LOCATION HERE! Revit doubles coordinates when setting from origin.
                    // Instead, we'll fix it in ApplyRotation when we detect origin.

                    idList.Add(elementId);
                    itemMap.Add((zone, plan));
                    elementCache[elementId] = inst;
                }

                // Step 3: Apply rotation & parameters (deferred via SleeveParameterService batching)
                const int batchLogInterval = 20;
                using (_performanceMonitor?.TrackOperation("Apply Rotation & Parameters"))
                {
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

                        // ⚡ DEBUG: What location does the cached instance have RIGHT NOW before rotation?
                        var cachedLoc = (instance.Location as LocationPoint)?.Point;
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [CACHE-RETRIEVE] [{i+1}/{idList.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                            $"InstId={elementId.IntegerValue}, CachedLocation=({cachedLoc?.X:F6}, {cachedLoc?.Y:F6}, {cachedLoc?.Z:F6})\n");

                        bool isWallOrFraming = IsWallOrFraming(zone);

                        // Geometric rotation: X wall or floor with rotated MEP
                        double rotationRad = isWallOrFraming
                            ? _rotationService.DetermineRotation(zone)
                            : plan.RotationRadians;

                        bool isXWall = string.Equals(zone.HostOrientation, "X", StringComparison.OrdinalIgnoreCase);
                        bool isFloor = string.Equals(zone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase);

                        // ⚡ DEBUG: Log rotation decision
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [ROTATION-CHECK] [{i+1}/{idList.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                            $"RotationRad={rotationRad:F6}, isXWall={isXWall}, isFloor={isFloor}, WillRotate={Math.Abs(rotationRad) > 0.001 && (isXWall || isFloor)}\n");

                        if (Math.Abs(rotationRad) > 0.001 && (isXWall || isFloor))
                        {
                            // Get expected placement point
                            XYZ expectedLocation = plan.PlacementPoint ??
                                new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

                            ApplyRotation(doc, instance, rotationRad, expectedLocation);
                        }

                        // Sizes and metadata (batched by SleeveParameterService)
                        _parameterService.SetSleeveParameters(
                            instance,
                            plan.TargetWidthFt,
                            plan.TargetHeightFt,
                            plan.TargetDiameterFt,
                            plan.IsCircular,
                            zone,
                            plan.RequiredDepthFt);

                        _parameterService.SetSleeveInstanceId(instance, elementId.IntegerValue);
                        zone.SleeveInstanceId = elementId.IntegerValue;

                        result.PlacedItems.Add((zone, elementId));
                        result.PlacedCount++;

                        if (result.PlacedCount % batchLogInterval == 0)
                        {
                            _logger($"[BulkPlacement] Apply Rotation & Parameters: processed {result.PlacedCount}/{idList.Count} elements");
                        }
                    }
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

                // ⚡ LOG THE AXIS POINT BEING USED FOR ROTATION
                SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] InstId={instance.Id.IntegerValue}, " +
                    $"RevitAxisPoint=({axisPoint1?.X:F6}, {axisPoint1?.Y:F6}, {axisPoint1?.Z:F6}), " +
                    $"ExpectedPoint=({expectedPlacementPoint.X:F6}, {expectedPlacementPoint.Y:F6}, {expectedPlacementPoint.Z:F6}), " +
                    $"RotationRad={rotationRad:F6}\n");

                if (axisPoint1 == null)
                {
                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] ❌ AxisPoint is NULL! Skipping rotation\n");
                    return;
                }

                // ✅ CRITICAL FIX: If axis point is (0,0,0), MOVE the instance to expected location!
                bool isAtOrigin = Math.Abs(axisPoint1.X) < 0.0001 && Math.Abs(axisPoint1.Y) < 0.0001 && Math.Abs(axisPoint1.Z) < 0.0001;
                if (isAtOrigin)
                {
                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] ⚠️ AXIS AT ORIGIN! Moving to expected point\n");

                    // MOVE the instance instead of setting location (avoids Revit doubling bug)
                    XYZ moveVector = expectedPlacementPoint - axisPoint1; // offset from (0,0,0) to expected
                    Autodesk.Revit.DB.ElementTransformUtils.MoveElement(doc, instance.Id, moveVector);

                    SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] 🔧 Moved instance by ({moveVector.X:F6}, {moveVector.Y:F6}, {moveVector.Z:F6})\n");

                    // ✅ VERIFY: Is it STILL at origin after move?
                    if (axisPoint1 != null && Math.Abs(axisPoint1.X) < 0.0001 && Math.Abs(axisPoint1.Y) < 0.0001)
                    {
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] 📍 Still at origin after move! Forcing absolute location.\n");
                        locPt.Point = expectedPlacementPoint;
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] 📍 Re-read location after move: ({axisPoint1?.X:F6}, {axisPoint1?.Y:F6}, {axisPoint1?.Z:F6})\n");
                    }
                }

                XYZ axisDirection = XYZ.BasisZ;
                XYZ axisPoint2 = axisPoint1 + axisDirection;
                Line axis = Line.CreateBound(axisPoint1, axisPoint2);

                // ⚡ LOG FINAL AXIS
                SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] FinalAxis: ({axisPoint1.X:F6}, {axisPoint1.Y:F6}, {axisPoint1.Z:F6}) to ({axisPoint2.X:F6}, {axisPoint2.Y:F6}, {axisPoint2.Z:F6})\n");

                instance.Location.Rotate(axis, rotationRad);

                SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] ✅ Rotation completed\n");
            }
            catch (Exception ex)
            {
                _logger($"[BulkPlacement] Rotation failed: {ex.Message}");
                SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [APPLY-ROTATION] ❌ EXCEPTION: {ex.Message}\n{ex.StackTrace}\n");
            }
        }

        /// <summary>
        /// Updates ClashZone objects with actual geometry from placed elements.
        /// Critical for ensuring clustering uses \"as-placed\" dimensions.
        /// </summary>
        public void UpdateZonesFromElements(Document doc, List<(ClashZone Zone, ElementId ElementId)> items)
        {
            if (items == null || items.Count == 0) return;

            int updateIndex = 0;
            foreach (var (zone, ElementId) in items)
            {
                try
                {
                    updateIndex++;
                    var element = doc.GetElement(ElementId) as FamilyInstance;
                    if (element == null) continue;

                    // ✅ Update actual placement point from Revit element
                    var locPt = element.Location as LocationPoint;
                    if (locPt != null)
                    {
                        // ⚡ LOG WHAT WE'RE READING FROM REVIT
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [UPDATE-ZONES] [{updateIndex}/{items.Count}] Zone={zone.Id.ToString().Substring(0,8)}, " +
                            $"Reading from Revit: ({locPt.Point.X:F6}, {locPt.Point.Y:F6}, {locPt.Point.Z:F6})\n");

                        zone.SleevePlacementPointX = locPt.Point.X;
                        zone.SleevePlacementPointY = locPt.Point.Y;
                        zone.SleevePlacementPointZ = locPt.Point.Z;

                        zone.SleevePlacementActiveX = locPt.Point.X;
                        zone.SleevePlacementActiveY = locPt.Point.Y;
                        zone.SleevePlacementActiveZ = locPt.Point.Z;

                        // ⚡ LOG WHAT WE'RE WRITING TO ZONE
                        SafeFileLogger.SafeAppendText("bulk_placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [UPDATE-ZONES] [{updateIndex}/{items.Count}] Writing to Zone: " +
                            $"ActiveX={zone.SleevePlacementActiveX:F6}, ActiveY={zone.SleevePlacementActiveY:F6}, ActiveZ={zone.SleevePlacementActiveZ:F6}\n");
                    }

                    // ✅ Extraction of actual Bounding Box corners from Revit (As-Built)
                    var bbox = element.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        zone.SleeveBoundingBoxMinX = bbox.Min.X;
                        zone.SleeveBoundingBoxMinY = bbox.Min.Y;
                        zone.SleeveBoundingBoxMinZ = bbox.Min.Z;
                        zone.SleeveBoundingBoxMaxX = bbox.Max.X;
                        zone.SleeveBoundingBoxMaxY = bbox.Max.Y;
                        zone.SleeveBoundingBoxMaxZ = bbox.Max.Z;
                    }

                    // Minimal implementation: update basic as-placed dimensions if needed
                    var widthParam = element.LookupParameter("Width") ?? element.LookupParameter("Sleeve Width");
                    var heightParam = element.LookupParameter("Height") ?? element.LookupParameter("Sleeve Height");

                    if (widthParam != null && widthParam.StorageType == StorageType.Double)
                        zone.SleeveWidth = widthParam.AsDouble();
                    if (heightParam != null && heightParam.StorageType == StorageType.Double)
                        zone.SleeveHeight = heightParam.AsDouble();

                    var diameterParam = element.LookupParameter("Diameter") ?? element.LookupParameter("Sleeve Diameter");
                    var depthParam = element.LookupParameter("Structural Depth") ?? element.LookupParameter("Sleeve Length") ?? element.LookupParameter("Opening Depth");

                    if (diameterParam != null && diameterParam.StorageType == StorageType.Double)
                        zone.SleeveDiameter = diameterParam.AsDouble();
                    if (depthParam != null && depthParam.StorageType == StorageType.Double)
                        zone.SleeveDepth = depthParam.AsDouble();

                    zone.SleeveFamilyName = element.Symbol.FamilyName;
                }
                catch (Exception ex)
                {
                    _logger($"[BulkPlacement] Failed to update zone {zone?.Id} from element: {ex.Message}");
                }
            }
        }
    }
}
