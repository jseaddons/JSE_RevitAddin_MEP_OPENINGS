using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 8: Persist placement data to database and XML.
    /// Will use ISleevePersistenceService when Team D delivers abstraction.
    /// </summary>
    public class PersistenceStage : IPlacementStage
    {
        public string Name => "PersistData";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                // TODO: Replace with ISleevePersistenceService when Team D delivers
                // For now, placeholder indicating persistence intent
                var persistedCount = 0;

                perf?.StopOperation(Name, persistedCount);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Persisted {persistedCount} placement records");
                }

                var updatedContext = context.WithMetric("RecordsPersisted", persistedCount);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                // Persistence failure is critical but shouldn't lose placed instances
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[{Name}] Persistence failed: {ex.Message}");
                }

                return StageResult.Failed(
                    context,
                    $"PersistenceStage failed: {ex.Message} (sleeves placed but not saved)"
                );
            }
        }
    }
}
