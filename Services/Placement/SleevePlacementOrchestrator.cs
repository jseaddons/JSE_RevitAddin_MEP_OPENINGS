using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Orchestrates sleeve placement through a pipeline of ordered stages.
    /// Implements fail-safe execution: stages return results, errors don't abort pipeline.
    /// Adheres to SOLID: depends on IPlacementStage abstractions, no business logic.
    /// </summary>
    public class SleevePlacementOrchestrator : ISleevePlacementOrchestrator
    {
        private readonly IReadOnlyList<IPlacementStage> _stages;
        private readonly IOptimizationConfig _config;
        private readonly Data.Repositories.IClashZoneRepository _clashZoneRepository; // optional
        private readonly Services.Interfaces.IParameterBatchingService _parameterBatchingService; // optional
        private readonly Services.Interfaces.IParameterSnapshotTransferService _snapshotTransferService; // optional

        public SleevePlacementOrchestrator(
            IReadOnlyList<IPlacementStage> stages,
            IOptimizationConfig config,
            Data.Repositories.IClashZoneRepository clashZoneRepository = null,
            Services.Interfaces.IParameterBatchingService parameterBatchingService = null,
            Services.Interfaces.IParameterSnapshotTransferService snapshotTransferService = null)
        {
            _stages = stages ?? throw new ArgumentNullException(nameof(stages));
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _clashZoneRepository = clashZoneRepository;
            _parameterBatchingService = parameterBatchingService;
            _snapshotTransferService = snapshotTransferService;
        }

        public OrchestratorResult Execute(Document doc, IReadOnlyList<ClashZone> zones, IPerformanceMonitor perf)
        {
            if (zones == null || zones.Count == 0)
            {
                return new OrchestratorResult(
                    success: true,
                    placedCount: 0,
                    errorCount: 0,
                    correlationId: Guid.NewGuid().ToString()
                );
            }

            // Initialize context with input zones and config
            var context = new PlacementContext(
                doc: doc,
                zones: zones,
                config: _config
            );

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ORCHESTRATOR] Starting placement pipeline. CorrelationId: {context.CorrelationId}, Version: {_config.Version}");
                DebugLogger.Info($"[ORCHESTRATOR] Input zones: {zones.Count}, Stages: {_stages.Count}");
            }

            perf?.StartOperation("OrchestratorExecute");

            // Execute stages sequentially with error boundaries
            var currentContext = context;
            var stageErrors = new List<string>();

            foreach (var stage in _stages)
            {
                try
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ORCHESTRATOR] Executing stage: {stage.Name}");
                    }

                    perf?.StartOperation($"Stage_{stage.Name}");
                    var stageResult = stage.Execute(currentContext, perf);
                    perf?.StopOperation($"Stage_{stage.Name}", 1);

                    if (!stageResult.Success)
                    {
                        // Fail-safe: log errors but continue to next stage
                        stageErrors.AddRange(stageResult.Errors);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[ORCHESTRATOR] Stage {stage.Name} reported {stageResult.Errors.Count} errors");
                            foreach (var error in stageResult.Errors)
                            {
                                DebugLogger.Warning($"  - {error}");
                            }
                        }
                    }

                    // Update context even on failure (stage may have partial results)
                    currentContext = stageResult.Context;
                }
                catch (Exception ex)
                {
                    // Critical stage failure: log and continue
                    var errorMsg = $"Stage {stage.Name} threw exception: {ex.Message}";
                    stageErrors.Add(errorMsg);

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[ORCHESTRATOR] {errorMsg}");
                        DebugLogger.Error($"[ORCHESTRATOR] Exception: {ex}");
                    }

                    SafeFileLogger.SafeAppendText("placement_errors.log",
                        $"[{DateTime.Now:HH:mm:ss}] [{context.CorrelationId}] Stage {stage.Name}: {ex.Message}");

                    // Add error to context and continue
                    currentContext = currentContext.WithError(errorMsg);
                }
            }

            perf?.StopOperation("OrchestratorExecute", currentContext.PlacedInstances.Count);

            // Aggregate final result
            var finalPlacedCount = currentContext.PlacedInstances.Count;
            var finalErrorCount = currentContext.Errors.Count + stageErrors.Count;
            var overallSuccess = finalPlacedCount > 0 || (finalErrorCount == 0 && zones.Count == 0);

            // OPTIONAL SNAPSHOT PARAMETER TRANSFER (post-pipeline)
            if (OptimizationFlags.EnableSnapshotParameterTransfer && OptimizationFlags.UseBatchedParameterWrites &&
                _snapshotTransferService != null && _parameterBatchingService != null && _clashZoneRepository != null &&
                currentContext.PlacedInstances.Any())
            {
                try
                {
                    var sleeveElementIds = currentContext.PlacedInstances.Select(fi => fi.Id).ToList();
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SNAPSHOT-PARAMS] 🚀 Starting snapshot transfer for {sleeveElementIds.Count} placed sleeves (CorrelationId={context.CorrelationId})");

                    var deferredCount = _snapshotTransferService.TransferSnapshotParameters(doc, sleeveElementIds, _clashZoneRepository, _parameterBatchingService);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SNAPSHOT-PARAMS] 📥 Deferred {deferredCount} parameters; regenerating document once...");

                    // Single regeneration prior to batch flush
                    doc.Regenerate();

                    var flushed = _parameterBatchingService.FlushDeferredParameters(doc);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[SNAPSHOT-PARAMS] ✅ Flush complete: wrote {flushed} parameters for {sleeveElementIds.Count} sleeves");
                }
                catch (Exception snapEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SNAPSHOT-PARAMS] ⚠️ Snapshot transfer failed: {snapEx.Message}");
                }
            }

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ORCHESTRATOR] Pipeline complete. CorrelationId: {context.CorrelationId}");
                DebugLogger.Info($"[ORCHESTRATOR] Placed: {finalPlacedCount}, Errors: {finalErrorCount}, Success: {overallSuccess}");
            }

            perf?.LogMetric("OrchestratorPlacedCount", finalPlacedCount);
            perf?.LogMetric("OrchestratorErrorCount", finalErrorCount);

            return new OrchestratorResult(
                success: overallSuccess,
                placedCount: finalPlacedCount,
                errorCount: finalErrorCount,
                correlationId: context.CorrelationId
            );
        }
    }
}
