using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Processes a single floor in isolation (thread-safe)
    /// Handles: Clash Detection → Placement → Clustering
    /// </summary>
    public class SingleFloorProcessor
    {
        private readonly Document _doc;
        private readonly SharedResourceCache _cache;
        private readonly IPerformanceMonitor? _monitor;
        
        public SingleFloorProcessor(
            Document doc, 
            SharedResourceCache cache, 
            IPerformanceMonitor? monitor = null)
        {
            _doc = doc;
            _cache = cache;
            _monitor = monitor;
        }
        
        /// <summary>
        /// Process a single floor with rollback capability
        /// </summary>
        public FloorProcessingResult ProcessFloor(Level level, OpeningFilter filter)
        {
            using (var transaction = new Transaction(_doc, $"Process Floor {level.Name}"))
            {
                transaction.Start();
                
                try
                {
                    // Step 1: Detect clashes
                    List<ClashZone> clashes;
                    using (var clashTracker = _monitor?.TrackOperation($"Detect Clashes - {level.Name}"))
                    {
                        var detector = new FloorIsolatedClashDetector(_doc, _cache);
                        var categories = GetSelectedCategories(filter);
                        clashes = detector.DetectClashesForFloor(level, categories);
                        clashTracker?.SetItemCount(clashes.Count);
                    }
                    
                    if (clashes.Count == 0)
                    {
                        transaction.RollBack();
                        return new FloorProcessingResult
                        {
                            LevelName = level.Name,
                            Success = true,
                            PlacedCount = 0,
                            Message = "No clashes detected"
                        };
                    }
                    
                    // Step 2: Place individual sleeves
                    int placedCount = 0;
                    var placedElements = new List<int>();
                    
                    if (clashes.Count > 0)
                    {
                        using (var placementTracker = _monitor?.TrackOperation($"Place Sleeves - {level.Name}"))
                        {
                            // A. Load Planning Conditions
                            var conditionsService = new ConditionsService(_doc);
                            var categories = GetSelectedCategories(filter);
                            
                            var plannedItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                            
                            foreach (var category in categories)
                            {
                                var categoryZones = clashes.Where(cz => cz.MepElementCategory == category).ToList();
                                if (categoryZones.Count == 0) continue;
                                
                                var categoryConditions = conditionsService.LoadConditions(category);
                                var planner = new Services.Placement.ParallelSleevePlacementPlanner(categoryConditions);
                                var planningResult = planner.Plan(categoryZones);
                                
                                var plannedMap = planningResult.Items.ToDictionary(i => i.ClashZoneId);
                                foreach (var zone in categoryZones)
                                {
                                    if (plannedMap.TryGetValue(zone.Id, out var plan))
                                    {
                                        if (plan.ShouldSkip) continue;
                                        
                                        // Ensure IsCurrentClash is true for placement to pass Context Flag check
                                        zone.IsCurrentClash = true;
                                        plannedItems.Add((zone, plan));
                                    }
                                }
                            }
                            
                            if (plannedItems.Count > 0)
                            {
                                // B. Execute Placement
                                // Instantiate BulkPlacementService (Autonomous placement without DB factory for floor-isolated run)
                                var placementService = new BulkPlacementService(
                                    _doc,
                                    contextFactory: () => new SleeveDbContext(_doc),
                                    logger: msg => SafeFileLogger.SafeAppendText("multifloor.log", $"[{DateTime.Now}] [PLACEMENT] {msg}\n"),
                                    performanceMonitor: _monitor);
                                
                                // ✅ PRIMARY CONSTRAINT: Respect spatial filtering (Section Box) as requested.
                                // If a section box is active, only items within it will be placed across all floors.
                                var placementResult = placementService.ExecuteBulkPlacement(_doc, plannedItems, skipSpatialFiltering: false);
                                
                                placedCount = placementResult.PlacedCount;
                                placedElements = placementResult.PlacedItems.Select(p => p.ElementId.GetIntegerValue()).ToList();
                                
                                SafeFileLogger.SafeAppendText("multifloor.log",
                                    $"[{DateTime.Now}] ✅ Level {level.Name}: Placed {placedCount} sleeves from {plannedItems.Count} planned items.\n");
                            }
                            
                            placementTracker?.SetItemCount(placedCount);
                        }
                    }
                    
                    // Step 3: Cluster analysis (optional)
                    int clusterCount = 0;
                    if (OptimizationFlags.EnableClusteringWorkflow && placedCount > 0)
                    {
                        using (var clusterTracker = _monitor?.TrackOperation($"Cluster Analysis - {level.Name}"))
                        {
                            // TODO: Add clustering logic if needed for multi-floor
                            clusterTracker?.SetItemCount(clusterCount);
                        }
                    }
                    
                    transaction.Commit();
                    
                    return new FloorProcessingResult
                    {
                        LevelName = level.Name,
                        Success = true,
                        PlacedCount = placedCount,
                        ClusterCount = clusterCount,
                        PlacedElements = placedElements
                    };
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    
                    SafeFileLogger.SafeAppendText("multifloor_errors.log",
                        $"[{DateTime.Now}] ❌ Floor {level.Name} rolled back: {ex.Message}\n{ex.StackTrace}\n");
                    
                    throw; // Re-throw to be caught by FloorBatchProcessor
                }
            }
        }
        
        /// <summary>
        /// Revit-only phase of processing (for use in sequential Revit loop)
        /// </summary>
        public FloorRawData ProcessFloor_RevitPhase(Level level, OpeningFilter filter)
        {
            // Placeholder for Revit-only work that can't be parallelized
            return new FloorRawData { LevelName = level.Name };
        }
        
        private List<string> GetSelectedCategories(OpeningFilter filter)
        {
            if (filter.SelectedMepCategoryNames?.Any() == true)
                return filter.SelectedMepCategoryNames;
            
            // Fallback to single category
            return new List<string> { filter.Category.ToString() };
        }
    }
}
