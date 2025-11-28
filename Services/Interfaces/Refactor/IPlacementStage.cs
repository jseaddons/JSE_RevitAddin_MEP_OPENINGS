using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Represents a single stage in the sleeve placement pipeline.
    /// Each stage adheres to Single Responsibility Principle (SRP).
    /// </summary>
    public interface IPlacementStage
    {
        /// <summary>
        /// Logical name of the stage for logging and diagnostics.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Execute this stage's work, returning updated context or errors.
        /// </summary>
        /// <param name="context">Immutable context; stage returns new context via With* methods</param>
        /// <param name="perf">Performance monitor to report stage timing</param>
        /// <returns>Result containing success status, updated context, and any errors</returns>
        JSE_RevitAddin_MEP_OPENINGS.Services.Placement.StageResult Execute(
            JSE_RevitAddin_MEP_OPENINGS.Services.Placement.PlacementContext context, 
            IPerformanceMonitor perf);
    }
}
