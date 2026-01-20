using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 6: Assign parameters to placed sleeve instances.
    /// Uses IParameterBatchingService if batching is enabled.
    /// </summary>
    public class ParameterAssignStage : IPlacementStage
    {
        private readonly IParameterBatchingService _batchingService;

        public ParameterAssignStage(IParameterBatchingService batchingService)
        {
            _batchingService = batchingService;
        }

        public string Name => "AssignParameters";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                if (!context.Config.Placement.UseParameterBatching)
                {
                    perf?.StopOperation(Name, 0);
                    return StageResult.Succeeded(context.WithMetric("ParameterBatchingEnabled", false));
                }

                // TODO: Defer parameters for placed instances
                // This will be populated when InstancePlacementStage actually places elements
                var deferredCount = _batchingService.DeferredParameterCount;

                perf?.StopOperation(Name, deferredCount);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Deferred {deferredCount} parameters for batching");
                }

                var updatedContext = context
                    .WithMetric("ParameterBatchingEnabled", true)
                    .WithMetric("ParametersDeferred", deferredCount);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"ParameterAssignStage failed: {ex.Message}");
            }
        }
    }
}
