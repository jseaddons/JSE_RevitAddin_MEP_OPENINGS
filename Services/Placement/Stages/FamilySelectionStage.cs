using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 4: Select appropriate family symbol based on zone characteristics.
    /// Placeholder until ISleeveFamilySelector abstraction is provided.
    /// </summary>
    public class FamilySelectionStage : IPlacementStage
    {
        public string Name => "SelectFamilies";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                // TODO: Implement ISleeveFamilySelector when abstraction is ready
                // For now, family selection happens inline during placement

                perf?.StopOperation(Name, context.FilteredZones.Count);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Family selection prepared for {context.FilteredZones.Count} zones");
                }

                var updatedContext = context.WithMetric("FamiliesSelected", context.FilteredZones.Count);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"FamilySelectionStage failed: {ex.Message}");
            }
        }
    }
}
