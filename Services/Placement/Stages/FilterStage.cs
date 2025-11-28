using System;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 1: Filter clash zones based on ReadyForPlacement flags and other criteria.
    /// Delegates to existing filtering logic until Team C provides IZonePreFilterService.
    /// </summary>
    public class FilterStage : IPlacementStage
    {
        public string Name => "FilterZones";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                // TODO: Replace with IZonePreFilterService when Team C delivers
                // For now, simple passthrough filtering (ReadyForPlacement=true)
                var filtered = context.InputZones
                    .Where(z => z.ReadyForPlacement)
                    .ToList();

                perf?.StopOperation(Name, filtered.Count);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Filtered {context.InputZones.Count} -> {filtered.Count} zones");
                }

                var updatedContext = context
                    .WithFilteredZones(filtered)
                    .WithMetric("FilteredCount", filtered.Count)
                    .WithMetric("FilteredOutCount", context.InputZones.Count - filtered.Count);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"FilterStage failed: {ex.Message}");
            }
        }
    }
}
