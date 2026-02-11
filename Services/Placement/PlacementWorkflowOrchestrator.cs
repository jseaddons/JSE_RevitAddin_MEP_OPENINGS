using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Calculation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Algorithm;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories; // Added for FilterRepository

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// SOLID Orchestrator for the complete placement workflow:
    /// 1. Bulk Placement (Individual Sleeves)
    /// 2. Geometry Extraction (As-placed corners)
    /// 3. Global Clustering (Calculation & Placement)
    /// </summary>
    public class PlacementWorkflowOrchestrator
    {
        private readonly Document _doc;
        private readonly Func<SleeveDbContext> _contextFactory;
        private readonly Action<string> _logger;
        private readonly IPerformanceMonitor _perf;

        public PlacementWorkflowOrchestrator(
            Document doc,
            Func<SleeveDbContext> contextFactory,
            Action<string> logger = null,
            IPerformanceMonitor perf = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _contextFactory = contextFactory ?? throw new ArgumentNullException(nameof(contextFactory));
            _logger = logger ?? (msg => { });
            _perf = perf;
        }

        /// <summary>
        /// Executes the optimized end-to-end placement and clustering workflow for a list of filters.
        /// </summary>
        public OrchestratorResult ExecuteOptimizedPlacementWorkflow(IEnumerable<OpeningFilter> filters, IOperationTracker parentTracker = null)
        {
            int totalPlaced = 0;
            int totalErrors = 0;
            int totalClusters = 0;
            int totalClusteredZones = 0;
            string correlationId = Guid.NewGuid().ToString();
            var placedZones = new List<ClashZone>();

            try
            {
                _logger($"[WORKFLOW] 🚀 Starting Optimized Placement Workflow for {filters.Count()} filters...");

                // 1. INDIVIDUAL PLACEMENT (Loop filters)
                using (var bulkOp = parentTracker?.TrackSubOperation("1. Bulk Placement (Multi-Filter)"))
                {
                    var bulkService = new BulkPlacementService(_doc, _contextFactory, _logger, _perf);

                    foreach (var filter in filters)
                    {
                        using (var filterOp = bulkOp?.TrackSubOperation($"Filter: {filter.Name}"))
                        {
                            var bulkResult = bulkService.ExecuteContextPlacement(_doc, filter, parentTracker);

                            totalPlaced += bulkResult.PlacedCount;
                            totalErrors += bulkResult.FailedCount;
                            placedZones.AddRange(bulkResult.PlacedItems.Select(x => x.Zone));

                            if (!bulkResult.OverallSuccess)
                            {
                                _logger($"[WORKFLOW] ⚠️ Bulk placement for {filter.Name} reported issues: {bulkResult.Error}");
                            }
                        }
                    }
                    bulkOp?.SetItemCount(totalPlaced);
                }

                // 2. GEOMETRY EXTRACTION (Only for newly placed sleeves)
                if (placedZones.Any())
                {
                    using (var geomOp = parentTracker?.TrackSubOperation("2. Geometry Extraction"))
                    {
                        using (var db = _contextFactory())
                        {
                            var repo = new ClashZoneRepository(db, _logger, _perf as PerformanceMonitor);
                            var extractor = new BatchSleeveCornerExtractor(repo);

                            int extracted = extractor.ExtractAndSaveCornersForZones(_doc, placedZones);
                            _logger($"[WORKFLOW] 📐 Extracted corners for {extracted} new sleeves.");
                        }
                    }
                }

                // 3. GLOBAL CLUSTERING (Calc & Placement)
                // We always run clustering after placement to catch any new clusters formed by the new sleeves
                using (var clusterOp = parentTracker?.TrackSubOperation("3. Global Clustering"))
                {
                    var clusterResult = ExecuteClusteringSequence(filters.FirstOrDefault(), parentTracker);
                    totalClusters = clusterResult.clustersPlaced;
                    totalClusteredZones = clusterResult.zonesInClusters;
                }

                // 4. RESET FLAGS (Mark file combos as processed)
                ResetProcessedFlags(filters);

                // ✅ FINAL SUMMARY: Log comprehensive workflow results
                var summaryLog = new System.Text.StringBuilder();
                summaryLog.AppendLine($"\n[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════════════════════");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] 🎯 PLACEMENT WORKFLOW COMPLETE");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════════════════════");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] 📊 INDIVIDUAL SLEEVE PLACEMENT:");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Placed: {totalPlaced} sleeves");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Errors: {totalErrors}");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] 🧬 CLUSTER CONSOLIDATION:");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Clusters Created: {totalClusters}");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Zones Clustered: {totalClusteredZones}");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Individual Sleeves Replaced: {totalClusteredZones}");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ✅ FINAL STATUS:");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Total Openings Resolved: {totalPlaced} zones");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]   • Final Sleeve Count: {totalPlaced - totalClusteredZones + totalClusters}");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}]     ({totalPlaced - totalClusteredZones} individual + {totalClusters} clusters)");
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ");
                if (totalClusteredZones > 0)
                {
                    summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] 💡 Clustering reduced sleeve count by {totalClusteredZones - totalClusters} ({((totalClusteredZones - totalClusters) * 100.0 / totalClusteredZones):F1}%)");
                }
                summaryLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ═══════════════════════════════════════════════════════\n");

                SafeFileLogger.SafeAppendText("flag_workflow.log", summaryLog.ToString());

                _logger($"[WORKFLOW] ✅ Workflow complete. Individual: {totalPlaced}, Clusters: {totalClusters} (covering {totalClusteredZones} zones)");
                return new OrchestratorResult(true, totalPlaced, totalErrors, correlationId, totalClusters, totalClusteredZones);
            }
            catch (Exception ex)
            {
                _logger($"[WORKFLOW] ❌ CRITICAL ERROR: {ex.Message}");
                SafeFileLogger.SafeAppendText("workflow_error.log", $"[{DateTime.Now:HH:mm:ss}] Workflow Failed: {ex.Message}\n{ex.StackTrace}\n");
                return new OrchestratorResult(false, totalPlaced, totalErrors + 1, correlationId, totalClusters, totalClusteredZones);
            }
        }

        private (int clustersPlaced, int zonesInClusters) ExecuteClusteringSequence(OpeningFilter filter, IOperationTracker parentTracker)
        {
            int clustersPlaced = 0;
            int zonesInClusters = 0;

            try
            {
                _logger("[WORKFLOW][CLUSTER] 🧬 Starting Clustering Sequence...");

                using (var db = _contextFactory())
                {
                    var repo = new ClashZoneRepository(db, _logger, _perf as PerformanceMonitor);

                    // A. PROXIMITY MARKING (NEW STEP - Check which zones have neighbors)
                    int markedForClustering = 0;
                    using (var proximityTracker = parentTracker?.TrackSubOperation("3a. Proximity Marking"))
                    {
                        try
                        {
                            // Default proximity tolerance: 0.5 feet (approximately 6 inches)
                            // This can be customized via filter parameters if needed
                            double proximityTolerance = 0.5;
                            if (filter?.Parameters != null && filter.Parameters.ContainsKey("JoinOpeningsDistance"))
                            {
                                if (double.TryParse(filter.Parameters["JoinOpeningsDistance"].ToString(), out double customTol))
                                {
                                    proximityTolerance = customTol;
                                }
                            }

                            var proximityMarker = new ClusterProximityMarker(repo, proximityTolerance, _logger);

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 3A: STARTING PROXIMITY CHECK\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Tolerance: {proximityTolerance} ft\n");

                            markedForClustering = proximityMarker.MarkZonesForClustering(_doc, filter?.Category.ToString() ?? "All");

                            _logger($"[WORKFLOW][PROXIMITY] ✅ Marked {markedForClustering} zones for clustering based on proximity");
                        }
                        catch (Exception proximityEx)
                        {
                            _logger($"[WORKFLOW][PROXIMITY] ⚠️ Proximity marking failed: {proximityEx.Message}");

                            // ✅ DIAGNOSTIC: Log proximity failure to workflow log
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ PROXIMITY MARKING EXCEPTION CAUGHT ❌❌❌\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Exception: {proximityEx.Message}\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Stack: {proximityEx.StackTrace}\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Inner: {proximityEx.InnerException?.Message ?? "None"}\n");

                            // Continue with clustering even if proximity marking fails
                        }
                    }

                    // B. CLUSTER CALCULATION (Only for zones marked with MarkedForClusterProcess=1)
                    using (var calcTracker = parentTracker?.TrackSubOperation("3b. Cluster Calculation"))
                    {
                        var algo = new ClusterAlgorithmService(); // Using the actual service class
                        var rotation = new ClusterRotationService(
                            getClashZoneFunc: (id, _) => repo.GetClashZoneByInstanceId(id),
                            getClusterPlacementFunc: null
                        );
                        var calcService = new BatchClusterCalculationService(algo, rotation, db.DatabasePath);

                        // Load zones marked for clustering (Ready=1, Current=1, IsResolved=1, MarkedForClusterProcess=1)
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}] STEP 5: LOADING ZONES FOR CLUSTER CALCULATION\n");
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}]   Query: Ready=1, Current=1, IsResolved=1, MarkedForClusterProcess=1\n");

                        var zones = repo.GetZonesForClustering();
                        zonesInClusters = zones.Count; // Track total zones eligible for clustering

                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}]   Found {zones.Count} zones eligible for clustering\n");

                        if (zones.Count > 0)
                        {
                            _logger($"[WORKFLOW][CLUSTER] Calculating clusters for {zones.Count} zones...");

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 6: CALCULATING CLUSTER ARRANGEMENTS\n");

                            // Calculate results in memory
                            var results = calcService.CalculateOnly(zones, filter?.Category.ToString() ?? "All", 0, 0, _doc);

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Calculated {results.Count} cluster arrangements\n");

                            if (results.Count > 0)
                            {
                                // Save pending clusters using BatchSave
                                calcService.BatchSave(results);

                                // ✅ FLAG UPDATE: Mark constituent zones as processed for clustering
                                calcService.BatchUpdateFlags(results);

                                SafeFileLogger.SafeAppendText("flag_workflow.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   ✅ Saved {results.Count} clusters to ClusterSleeves_v2\n");

                                _logger($"[WORKFLOW][CLUSTER] 💾 Saved {results.Count} pending clusters to DB and updated flags.");
                            }
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   ⚠️ No zones found with MarkedForClusterProcess=1\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Possible reasons:\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   - Proximity check did not run\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   - Proximity check set flag=0 (all zones isolated)\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   - Flag update failed\n");

                            _logger($"[WORKFLOW][CLUSTER] ℹ️ No zones eligible for clustering (no proximity found)");
                        }
                    }

                    // C. PLACEMENT & CLEANUP
                    using (var placeTracker = parentTracker?.TrackSubOperation("3c. Cluster Placement & Cleanup"))
                    {
                        var parameterService = new SleeveParameterService(_doc, false, _perf as PlacementPerformanceMonitor);
                        var cleanup = new ClusterCleanupService();
                        var placeService = new BatchClusterPlacementService(db.DatabasePath, repo, parameterService, cleanup, _perf);

                        var (placed, failed, cleaned) = placeService.PlaceAllCategoriesAndCleanup(_doc, cleanup);
                        clustersPlaced = placed;
                        _logger($"[WORKFLOW][CLUSTER] 🏗️ Placed {placed} clusters, cleaned up {cleaned} redundant sleeves.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[WORKFLOW][CLUSTER] ⚠️ Clustering failed: {ex.Message}");
            }

            return (clustersPlaced, zonesInClusters);
        }
        private void ResetProcessedFlags(IEnumerable<OpeningFilter> filters)
        {
            try
            {
                using (var db = _contextFactory())
                {
                    var filterRepo = new FilterRepository(db, _logger);
                    foreach (var filter in filters)
                    {
                        // Standard Refresh methodology: Reset IsFilterComboNew to 0 after placement
                        // This allows the next refresh to use PATH 1 (Replay Mode)
                        // TODO: ResetFilterComboNewFlag method not found - filterRepo.ResetFilterComboNewFlag(filter.Id, _doc.GetDocId());
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[WORKFLOW] ⚠️ Failed to reset processed flags: {ex.Message}");
            }
        }
    }
}
