using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;  // For PlacementContext from Team A

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Service for adaptive parallel planning of sleeve placements.
    /// Encapsulates parallel clearance calculation with threshold-based activation.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public interface IParallelPlanningService
    {
        /// <summary>
        /// Plan sleeve dimensions for zones, using parallel processing if beneficial.
        /// Returns updated context with dimension data (immutable pattern).
        /// NOTE: Uses PlacementContext from Team A (Services.Placement namespace).
        /// </summary>
        /// <param name="zones">Clash zones requiring dimension planning</param>
        /// <param name="ctx">Placement context to read and update</param>
        /// <returns>Updated placement context with dimension results</returns>
        PlacementContext PlanDimensions(IList<ClashZone> zones, PlacementContext ctx);
        
        /// <summary>
        /// Determine if parallel planning should be used based on zone count and configuration.
        /// </summary>
        /// <param name="zoneCount">Number of zones to process</param>
        /// <returns>True if parallel processing is beneficial, false for sequential</returns>
        bool ShouldUseParallelPlanning(int zoneCount);
        
        /// <summary>
        /// Get the threshold zone count above which parallel planning is activated.
        /// </summary>
        int ParallelThreshold { get; }
        
        /// <summary>
        /// Check if parallel planning is currently enabled via configuration.
        /// </summary>
        bool IsEnabled { get; }
    }
}
