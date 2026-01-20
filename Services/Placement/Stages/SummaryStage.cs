using System;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 10: Log summary metrics and finalize performance reporting.
    /// Aggregates data from all previous stages for comprehensive diagnostics.
    /// </summary>
    public class SummaryStage : IPlacementStage
    {
        public string Name => "Summary";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                var placed = context.PlacedInstances.Count;
                var errors = context.Errors.Count;
                var inputCount = context.InputZones.Count;
                var filteredCount = context.Metrics.ContainsKey("FilteredCount") 
                    ? Convert.ToInt32(context.Metrics["FilteredCount"]) 
                    : 0;

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var summary = $@"
=== PLACEMENT PIPELINE SUMMARY ===
CorrelationId: {context.CorrelationId}
Config Version: {context.Config.Version}
Input Zones: {inputCount}
Filtered Zones: {filteredCount}
Placed Instances: {placed}
Errors: {errors}
Success Rate: {(placed > 0 ? (placed * 100.0 / Math.Max(filteredCount, 1)) : 0):F2}%
Features:
  - Parameter Batching: {context.Config.Placement.UseParameterBatching}
  - Smart Replay: {context.Config.Placement.UseSmartReplay}
  - Parallel Clearance: {context.Config.Placement.UseParallelClearance}
  - R-Tree Index: {context.Config.Detection.UseRTreeIndex}
  - Spatial Grid: {context.Config.Detection.UseSpatialGrid}
==================================";

                    DebugLogger.Info($"[{Name}] {summary}");
                    SafeFileLogger.SafeAppendText("placement_summary.log", 
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {summary}");
                }

                perf?.StopOperation(Name, 1);
                perf?.LogMetric("SummaryPlaced", placed);
                perf?.LogMetric("SummaryErrors", errors);

                return StageResult.Succeeded(context);
            }
            catch (Exception ex)
            {
                // Summary stage failure shouldn't affect placement result
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[{Name}] Summary generation failed: {ex.Message}");
                }

                return StageResult.Succeeded(context.WithError($"SummaryStage warning: {ex.Message}"));
            }
        }
    }
}
