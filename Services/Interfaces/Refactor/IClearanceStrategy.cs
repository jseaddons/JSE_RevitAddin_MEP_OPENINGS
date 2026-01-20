using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// High-level abstraction for clearance calculation strategies.
    /// Delegates to category-specific providers while maintaining consistent interface.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public interface IClearanceStrategy
    {
        /// <summary>
        /// Calculate clearance for a clash zone based on MEP element and structural host.
        /// </summary>
        /// <param name="zone">The clash zone requiring clearance calculation</param>
        /// <param name="conditions">Opening conditions containing clearance rules</param>
        /// <returns>Clearance value in Revit internal units (feet)</returns>
        double CalculateClearance(ClashZone zone, OpeningConditions conditions);
        
        /// <summary>
        /// Determine if clearance calculation should use parallel processing based on zone count.
        /// </summary>
        /// <param name="zoneCount">Number of zones to process</param>
        /// <returns>True if parallel processing is beneficial, false otherwise</returns>
        bool ShouldUseParallelCalculation(int zoneCount);
        
        /// <summary>
        /// Get the name of the strategy for logging and diagnostics.
        /// </summary>
        string StrategyName { get; }
    }
}
