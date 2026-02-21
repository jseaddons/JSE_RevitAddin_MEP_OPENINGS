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

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor.Safety
{
    /// <summary>
    /// Crash-safe single floor processor with comprehensive error handling.
    /// Wraps SingleFloorProcessor with timeout, validation, and failure preprocessing.
    /// </summary>
    public class SafeSingleFloorProcessor
    {
        private readonly Document _doc;
        private readonly SharedResourceCache _cache;
        private readonly IPerformanceMonitor _monitor;
        private readonly MultiFloorValidator _validator;
        private readonly MultiFloorTimeoutMonitor _timeoutMonitor;
        
        public SafeSingleFloorProcessor(
            Document doc,
            SharedResourceCache cache,
            IPerformanceMonitor monitor = null,
            MultiFloorTimeoutMonitor timeoutMonitor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _cache = cache ?? throw new ArgumentNullException(nameof(cache));
            _monitor = monitor;
            _validator = new MultiFloorValidator(msg => SafeFileLogger.SafeAppendText("multifloor.log", msg));
            _timeoutMonitor = timeoutMonitor;
        }
        
        /// <summary>
        /// Process a single floor with comprehensive crash protection
        /// </summary>
        public FloorProcessingResult ProcessFloorSafe(Level level, OpeningFilter filter)
        {
            var floorName = level?.Name ?? "Unknown";
            
            // Start timeout monitoring for this floor
            _timeoutMonitor?.StartFloor(floorName);
            
            try
            {
                // Pre-validation
                var validation = _validator.ValidateLevel(level, _doc);
                if (!validation.IsValid)
                {
                    var errorMsg = $"Pre-validation failed: {string.Join("; ", validation.Errors)}";
                    SafeFileLogger.SafeAppendText("multifloor.log",
                        $"[{DateTime.Now}] ❌ Floor {floorName}: {errorMsg}\n");
                        
                    return new FloorProcessingResult
                    {
                        LevelName = floorName,
                        Success = false,
                        ErrorMessage = errorMsg
                    };
                }
                
                // Execute with transaction and failure handling
                return ExecuteWithTransaction(level, filter);
            }
            catch (TimeoutException)
            {
                var msg = $"Floor {floorName} exceeded timeout limit";
                SafeFileLogger.SafeAppendText("multifloor.log",
                    $"[{DateTime.Now}] ⏱ {msg}\n");
                    
                return new FloorProcessingResult
                {
                    LevelName = floorName,
                    Success = false,
                    ErrorMessage = msg
                };
            }
            catch (Exception ex)
            {
                var msg = $"Unexpected error: {ex.Message}";
                SafeFileLogger.SafeAppendText("multifloor_errors.log",
                    $"[{DateTime.Now}] ❌ Floor {floorName}: {msg}\n{ex.StackTrace}\n");
                    
                return new FloorProcessingResult
                {
                    LevelName = floorName,
                    Success = false,
                    ErrorMessage = msg
                };
            }
            finally
            {
                _timeoutMonitor?.EndFloor(floorName);
            }
        }
        
        /// <summary>
        /// Execute floor processing with transaction and failure preprocessing
        /// </summary>
        private FloorProcessingResult ExecuteWithTransaction(Level level, OpeningFilter filter)
        {
            var floorName = level.Name;
            
            using (var transaction = new Transaction(_doc, $"Process Floor {floorName}"))
            {
                // Configure failure handling
                var failureOptions = transaction.GetFailureHandlingOptions();
                failureOptions.SetFailuresPreprocessor(
                    new MultiFloorFailurePreprocessor(floorName, msg => 
                        SafeFileLogger.SafeAppendText("multifloor.log", $"[{DateTime.Now}] {msg}\n")));
                failureOptions.SetClearAfterRollback(true);
                failureOptions.SetDelayedMiniWarnings(false);
                transaction.SetFailureHandlingOptions(failureOptions);
                
                transaction.Start();
                
                try
                {
                    // Check timeout before starting work
                    if (_timeoutMonitor?.CheckFloorTimeout() == true)
                    {
                        transaction.RollBack();
                        throw new TimeoutException($"Floor {floorName} exceeded timeout");
                    }
                    
                    // Execute the actual processing
                    var result = ProcessFloorInternal(level, filter);
                    
                    // Check timeout before commit
                    if (_timeoutMonitor?.CheckFloorTimeout() == true)
                    {
                        transaction.RollBack();
                        throw new TimeoutException($"Floor {floorName} exceeded timeout during processing");
                    }
                    
                    if (result.Success)
                    {
                        transaction.Commit();
                        SafeFileLogger.SafeAppendText("multifloor.log",
                            $"[{DateTime.Now}] ✅ Floor {floorName}: Committed {result.PlacedCount} placements\n");
                    }
                    else
                    {
                        transaction.RollBack();
                        SafeFileLogger.SafeAppendText("multifloor.log",
                            $"[{DateTime.Now}] ⚠️ Floor {floorName}: Rolled back - {result.ErrorMessage}\n");
                    }
                    
                    return result;
                }
                catch (Exception ex)
                {
                    // Attempt rollback on any exception
                    try
                    {
                        if (transaction.HasStarted())
                        {
                            transaction.RollBack();
                            SafeFileLogger.SafeAppendText("multifloor.log",
                                $"[{DateTime.Now}] 🔄 Floor {floorName}: Rolled back due to exception\n");
                        }
                    }
                    catch (Exception rollbackEx)
                    {
                        SafeFileLogger.SafeAppendText("multifloor_errors.log",
                            $"[{DateTime.Now}] ❌ Floor {floorName}: Rollback failed: {rollbackEx.Message}\n");
                    }
                    
                    throw; // Re-throw to be handled by outer catch
                }
            }
        }
        
        /// <summary>
        /// Internal processing logic (separated for clarity)
        /// </summary>
        private FloorProcessingResult ProcessFloorInternal(Level level, OpeningFilter filter)
        {
            var floorName = level.Name;
            
            // Step 1: Detect clashes
            List<ClashZone> clashes;
            using (var clashTracker = _monitor?.TrackOperation($"Detect Clashes - {floorName}"))
            {
                var detector = new FloorIsolatedClashDetector(_doc, _cache);
                var categories = GetSelectedCategories(filter);
                clashes = detector.DetectClashesForFloor(level, categories);
                
                // Validate clashes
                var clashValidation = _validator.ValidateClashZones(clashes, _doc);
                clashes = clashValidation.ValidClashes;
                
                clashTracker?.SetItemCount(clashes.Count);
            }
            
            if (clashes.Count == 0)
            {
                return new FloorProcessingResult
                {
                    LevelName = floorName,
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
                using (var placementTracker = _monitor?.TrackOperation($"Place Sleeves - {floorName}"))
                {
                    // Check timeout
                    if (_timeoutMonitor?.CheckFloorTimeout() == true)
                    {
                        return new FloorProcessingResult
                        {
                            LevelName = floorName,
                            Success = false,
                            ErrorMessage = "Timeout during placement planning"
                        };
                    }
                    
                    // Load Planning Conditions
                    var conditionsService = new ConditionsService(_doc);
                    var categories = GetSelectedCategories(filter);
                    
                    var plannedItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                    
                    foreach (var category in categories)
                    {
                        // Check timeout periodically
                        if (_timeoutMonitor?.CheckFloorTimeout() == true)
                        {
                            return new FloorProcessingResult
                            {
                                LevelName = floorName,
                                Success = false,
                                ErrorMessage = "Timeout during category processing"
                            };
                        }
                        
                        var categoryZones = clashes.Where(cz => cz.MepElementCategory == category).ToList();
                        if (categoryZones.Count == 0) continue;
                        
                        var categoryConditions = conditionsService.LoadConditions(category);
                        var planner = new ParallelSleevePlacementPlanner(categoryConditions);
                        var planningResult = planner.Plan(categoryZones);
                        
                        var plannedMap = planningResult.Items.ToDictionary(i => i.ClashZoneId);
                        foreach (var zone in categoryZones)
                        {
                            if (plannedMap.TryGetValue(zone.Id, out var plan))
                            {
                                if (plan.ShouldSkip) continue;
                                
                                zone.IsCurrentClash = true;
                                plannedItems.Add((zone, plan));
                            }
                        }
                    }
                    
                    if (plannedItems.Count > 0)
                    {
                        var placementService = new BulkPlacementService(
                            _doc,
                            contextFactory: () => new SleeveDbContext(_doc),
                            logger: msg => SafeFileLogger.SafeAppendText("multifloor.log", $"[{DateTime.Now}] [PLACEMENT] {msg}\n"),
                            performanceMonitor: _monitor);
                        
                        var placementResult = placementService.ExecuteBulkPlacement(_doc, plannedItems, skipSpatialFiltering: false);
                        
                        placedCount = placementResult.PlacedCount;
                        placedElements = placementResult.PlacedItems.Select(p => p.ElementId.GetIntegerValue()).ToList();
                        
                        SafeFileLogger.SafeAppendText("multifloor.log",
                            $"[{DateTime.Now}] ✅ Level {floorName}: Placed {placedCount} sleeves from {plannedItems.Count} planned items.\n");
                    }
                    
                    placementTracker?.SetItemCount(placedCount);
                }
            }
            
            // Step 3: Cluster analysis (optional)
            int clusterCount = 0;
            if (OptimizationFlags.EnableClusteringWorkflow && placedCount > 0)
            {
                using (var clusterTracker = _monitor?.TrackOperation($"Cluster Analysis - {floorName}"))
                {
                    // TODO: Add clustering with timeout checks
                    clusterTracker?.SetItemCount(clusterCount);
                }
            }
            
            return new FloorProcessingResult
            {
                LevelName = floorName,
                Success = true,
                PlacedCount = placedCount,
                ClusterCount = clusterCount,
                PlacedElements = placedElements
            };
        }
        
        private List<string> GetSelectedCategories(OpeningFilter filter)
        {
            if (filter.SelectedMepCategoryNames?.Any() == true)
                return filter.SelectedMepCategoryNames;
            
            return new List<string> { filter.Category.ToString() };
        }
    }
}
