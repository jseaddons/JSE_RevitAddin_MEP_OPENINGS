using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 3: Calculate sleeve dimensions based on MEP element size and clearance strategy.
    /// Delegates to IClearanceStrategy when Team C delivers, currently uses placeholder logic.
    /// </summary>
    public class DimensionStage : IPlacementStage
    {
        public string Name => "CalculateDimensions";

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                // TODO: Replace with IClearanceStrategy and IParallelPlanningService from Team C
                var dimensions = new Dictionary<Guid, (double, double, double, bool)>();

                foreach (var zone in context.FilteredZones)
                {
                    // Placeholder: use existing zone dimensions or default clearance
                    var isCircular = zone.MepElementSize > 0;
                    var clearance = 0.1; // 100mm default, TODO: use IClearanceStrategy

                    if (isCircular)
                    {
                        var diameter = zone.MepElementSize + (2 * clearance);
                        dimensions[zone.Id] = (0, 0, diameter, true);
                    }
                    else
                    {
                        var width = zone.MepElementWidth + (2 * clearance);
                        var height = zone.MepElementHeight + (2 * clearance);
                        dimensions[zone.Id] = (width, height, 0, false);
                    }
                }

                perf?.StopOperation(Name, dimensions.Count);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Calculated dimensions for {dimensions.Count} zones");
                }

                var updatedContext = context
                    .WithDimensions(dimensions)
                    .WithMetric("DimensionsCalculated", dimensions.Count);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"DimensionStage failed: {ex.Message}");
            }
        }
    }
}
