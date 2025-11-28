using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 7: Flush deferred parameters after document regeneration.
    /// Critical for batching performance optimization (4-6× speedup).
    /// </summary>
    public class FlushParametersStage : IPlacementStage
    {
        private readonly IParameterBatchingService _batchingService;

        public FlushParametersStage(IParameterBatchingService batchingService)
        {
            _batchingService = batchingService;
        }

        public string Name => "FlushParameters";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                if (!context.Config.Placement.UseParameterBatching || 
                    _batchingService.DeferredElementCount == 0)
                {
                    perf?.StopOperation(Name, 0);
                    return StageResult.Succeeded(context.WithMetric("ParametersFlushed", 0));
                }

                // Regenerate document once before flushing all parameters
                context.Doc.Regenerate();

                var flushedCount = _batchingService.FlushDeferredParameters(context.Doc);

                perf?.StopOperation(Name, flushedCount);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Flushed {flushedCount} parameters after regeneration");
                }

                var updatedContext = context.WithMetric("ParametersFlushed", flushedCount);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                // Fail-safe: parameter flush errors shouldn't abort entire placement
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[{Name}] Parameter flush failed: {ex.Message}");
                }

                return StageResult.Failed(
                    context.WithMetric("ParametersFlushed", 0),
                    $"FlushParametersStage failed: {ex.Message}"
                );
            }
        }
    }
}
