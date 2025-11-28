using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 5: Place sleeve instances in the Revit document.
    /// Currently a placeholder; will delegate to ISleeveInstancePlacer.
    /// </summary>
    public class InstancePlacementStage : IPlacementStage
    {
        public string Name => "PlaceInstances";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                // TODO: Implement actual placement using ISleeveInstancePlacer
                // For now, placeholder that tracks intent
                var placedCount = 0;
                var errors = new List<string>();

                perf?.StopOperation(Name, placedCount);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Placed {placedCount} sleeve instances");
                }

                var updatedContext = context
                    .WithMetric("InstancesPlaced", placedCount)
                    .WithErrors(errors);

                return errors.Count == 0 
                    ? StageResult.Succeeded(updatedContext)
                    : StageResult.Failed(updatedContext, errors);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"InstancePlacementStage failed: {ex.Message}");
            }
        }
    }
}
