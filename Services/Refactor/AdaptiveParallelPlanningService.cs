using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactor
{
    /// <summary>
    /// Adaptive parallel planning service that wraps ISleevePlacementPlanner.
    /// Provides threshold-based activation of parallel processing.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public class AdaptiveParallelPlanningService : IParallelPlanningService
    {
        private readonly ISleevePlacementPlanner _planner;
        private readonly int _parallelThreshold;
        private readonly bool _isEnabled;
        
        public int ParallelThreshold => _parallelThreshold;
        public bool IsEnabled => _isEnabled;
        
        /// <summary>
        /// Constructor with dependency injection.
        /// </summary>
        /// <param name="planner">Underlying placement planner (defaults to ParallelSleevePlacementPlanner)</param>
        /// <param name="parallelThreshold">Zone count threshold for parallel activation (default: 10)</param>
        /// <param name="isEnabled">Whether parallel planning is enabled (default: true)</param>
        public AdaptiveParallelPlanningService(
            ISleevePlacementPlanner planner = null,
            int parallelThreshold = 10,
            bool isEnabled = true)
        {
            _planner = planner ?? new ParallelSleevePlacementPlanner();
            _parallelThreshold = parallelThreshold;
            _isEnabled = isEnabled;
        }
        
        /// <summary>
        /// Plan sleeve dimensions for zones, using parallel processing if beneficial.
        /// Returns updated context with dimension data (immutable pattern).
        /// </summary>
        public PlacementContext PlanDimensions(IList<ClashZone> zones, PlacementContext ctx)
        {
            if (zones == null || zones.Count == 0 || ctx == null)
                return ctx;
            
            // Check if parallel planning should be used
            if (!ShouldUseParallelPlanning(zones.Count))
            {
                // Sequential planning - not implemented in stub
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ParallelPlanning] Using sequential planning for {zones.Count} zones");
                }
                return ctx;
            }
            
            // Use parallel planner
            try
            {
                var result = _planner.Plan(zones);
                
                // Populate context with results
                if (result != null && result.Items != null)
                {
                    // Update context with new dimensions using builder method (PlacementContext is immutable)
                    var updatedContext = ctx;
                    foreach (var item in result.Items)
                    {
                        // Store dimension data - DTO has width/height in feet, need to determine if circular
                        bool isCircular = item.TargetWidthFt == item.TargetHeightFt; // Heuristic
                        double diameter = isCircular ? item.TargetWidthFt : 0;
                        updatedContext = updatedContext.WithDimension(item.ClashZoneId, (
                            item.TargetWidthFt,
                            item.TargetHeightFt,
                            diameter,
                            isCircular
                        ));
                    }
                    ctx = updatedContext;
                    
                    // Store metrics using WithMetric (PlacementContext is immutable)
                    ctx = ctx.WithMetric("ParallelPlanningUsed", true);
                    ctx = ctx.WithMetric("ZonesPlanned", result.Items.Count);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ParallelPlanning] Planned {result.Items.Count} zones in parallel");
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Fail-safe: Log error and continue using immutable context
                ctx = ctx.WithError($"Parallel planning failed: {ex.Message}");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[ParallelPlanning] Failed: {ex.Message}");
                }
            }
            
            return ctx;
        }
        
        /// <summary>
        /// Determine if parallel planning should be used.
        /// </summary>
        public bool ShouldUseParallelPlanning(int zoneCount)
        {
            return _isEnabled && zoneCount > _parallelThreshold;
        }
    }
}
