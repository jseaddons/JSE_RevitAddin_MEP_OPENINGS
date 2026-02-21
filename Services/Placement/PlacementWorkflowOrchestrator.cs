using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
using JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
// using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories; // FIX: CS0105 duplicate using directive

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
        /// ✅ CRASH-SAFE: Includes timeout monitoring to prevent Revit hangs
        /// </summary>
        public OrchestratorResult ExecuteOptimizedPlacementWorkflow(IEnumerable<OpeningFilter> filters, IOperationTracker parentTracker = null)
        {
            int totalPlaced = 0;
            int totalErrors = 0;
            int totalClusters = 0;
            int totalClusteredZones = 0;
            string correlationId = Guid.NewGuid().ToString();
            var placedZones = new List<ClashZone>();
            
            // ✅ CRASH SAFETY: Initialize timeout monitor (5 min per filter, 15 min total)
            var timeoutMonitor = new MultiFloorTimeoutMonitor(
                floorTimeoutMinutes: 5, 
                totalTimeoutMinutes: 15, 
                logger: msg => _logger($"[WORKFLOW-TIMEOUT] {msg}"));
            timeoutMonitor.StartOperation();
            
            _logger($"[WORKFLOW] 🚀 Starting Optimized Placement Workflow for {filters.Count()} filters...");
            _logger($"[WORKFLOW] ⏱️ Timeout: 5min per filter, 15min total");
            var workflowStopwatch = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                // 1. INDIVIDUAL PLACEMENT (Loop filters)
                using (var bulkOp = parentTracker?.TrackSubOperation("1. Bulk Placement (Multi-Filter)"))
                {
                    // ✅ Instantiate snapshot transfer service for bulk placement
                    var snapshotTransferService = new ParameterSnapshotTransferService();
                    var bulkService = new BulkPlacementService(_doc, _contextFactory, _logger, _perf, null, snapshotTransferService);

                    foreach (var filter in filters)
                    {
                        // ✅ CRASH SAFETY: Check timeout before processing each filter
                        if (timeoutMonitor.CheckTotalTimeout())
                        {
                            _logger($"[WORKFLOW] ⏱️ TOTAL TIMEOUT: Stopping placement after {totalPlaced} sleeves");
                            throw new TimeoutException($"Placement workflow exceeded 15 minute limit. Placed {totalPlaced} sleeves before timeout.");
                        }

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
                        
                        // ✅ CRASH SAFETY: Yield to allow Revit to process messages
                        System.Windows.Forms.Application.DoEvents();
                    }
                    bulkOp?.SetItemCount(totalPlaced);
                }
                _logger($"[WORKFLOW] ⏱️ PHASE 1 (Individual Placement): {workflowStopwatch.ElapsedMilliseconds}ms for {totalPlaced} sleeves");

                // ✅ CRASH SAFETY: Check timeout before geometry extraction
                if (timeoutMonitor.CheckTotalTimeout())
                {
                    _logger($"[WORKFLOW] ⏱️ TOTAL TIMEOUT: Skipping geometry extraction after {totalPlaced} sleeves");
                    return new OrchestratorResult(true, totalPlaced, totalErrors, correlationId, totalClusters, totalClusteredZones, 
                        $"Partial success - timeout after individual placement ({totalPlaced} sleeves)");
                }

                // 2. GEOMETRY EXTRACTION (Only for newly placed sleeves)
                // ✅ PERF: Compute corners mathematically from placement data (no Revit API calls)
                // Previously called ExtractAndSaveCornersForZones which did doc.GetElement() + get_BoundingBox()
                // per sleeve. Now uses CalculatedSleeveWidth/Height + SleevePlacementPoint already on the zone.
                if (placedZones.Any())
                {
                    var geomSw = System.Diagnostics.Stopwatch.StartNew();
                    using (var geomOp = parentTracker?.TrackSubOperation("2. Geometry Extraction"))
                    {
                        using (var db = _contextFactory())
                        {
                            var repo = new ClashZoneRepository(db, _logger, _perf as PerformanceMonitor);
                            var extractor = new BatchSleeveCornerExtractor(repo);

                            int extracted = extractor.ComputeAndSaveCornersFromPlacementData(placedZones);
                            geomSw.Stop();
                            _logger($"[WORKFLOW] ⏱️ PHASE 2 (Corner Computation): {geomSw.ElapsedMilliseconds}ms for {extracted} sleeves (math-only).");
                        }
                    }
                }

                // ✅ CRASH SAFETY: Check timeout before clustering
                if (timeoutMonitor.CheckTotalTimeout())
                {
                    _logger($"[WORKFLOW] ⏱️ TOTAL TIMEOUT: Skipping clustering after {totalPlaced} sleeves");
                    timeoutMonitor.Dispose();
                    return new OrchestratorResult(true, totalPlaced, totalErrors, correlationId, totalClusters, totalClusteredZones, 
                        $"Partial success - timeout before clustering ({totalPlaced} sleeves placed)");
                }

                // 3. GLOBAL CLUSTERING (Calc & Placement)
                // ✅ FIX: Run clustering for ALL filter categories, not just the first one
                var clusterSw = System.Diagnostics.Stopwatch.StartNew();
                using (var clusterOp = parentTracker?.TrackSubOperation("3. Global Clustering"))
                {
                    var clusterResult = ExecuteClusteringSequence(filters, parentTracker, timeoutMonitor);
                    totalClusters = clusterResult.clustersPlaced;
                    totalClusteredZones = clusterResult.zonesInClusters;
                }
                clusterSw.Stop();
                _logger($"[WORKFLOW] ⏱️ PHASE 3 (Clustering): {clusterSw.ElapsedMilliseconds}ms ({totalClusters} clusters, {totalClusteredZones} zones)");

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

                workflowStopwatch.Stop();
                _logger($"[WORKFLOW] ✅ Workflow complete in {workflowStopwatch.ElapsedMilliseconds}ms. Individual: {totalPlaced}, Clusters: {totalClusters} (covering {totalClusteredZones} zones)");
                timeoutMonitor.Dispose();
                return new OrchestratorResult(true, totalPlaced, totalErrors, correlationId, totalClusters, totalClusteredZones);
            }
            catch (TimeoutException tex)
            {
                _logger($"[WORKFLOW] ⏱️ TIMEOUT: {tex.Message}");
                SafeFileLogger.SafeAppendText("workflow_error.log", $"[{DateTime.Now:HH:mm:ss}] Workflow Timeout: {tex.Message}\n");
                timeoutMonitor?.Dispose();
                return new OrchestratorResult(true, totalPlaced, totalErrors, correlationId, totalClusters, totalClusteredZones, 
                    $"Partial success - timeout ({totalPlaced} sleeves placed)");
            }
            catch (Exception ex)
            {
                _logger($"[WORKFLOW] ❌ CRITICAL ERROR: {ex.Message}");
                SafeFileLogger.SafeAppendText("workflow_error.log", $"[{DateTime.Now:HH:mm:ss}] Workflow Failed: {ex.Message}\n{ex.StackTrace}\n");
                timeoutMonitor?.Dispose();
                return new OrchestratorResult(false, totalPlaced, totalErrors + 1, correlationId, totalClusters, totalClusteredZones, 
                    $"Error: {ex.Message}");
            }
        }

        private (int clustersPlaced, int zonesInClusters) ExecuteClusteringSequence(IEnumerable<OpeningFilter> filters, IOperationTracker parentTracker, MultiFloorTimeoutMonitor timeoutMonitor = null)
        {
            int clustersPlaced = 0;
            int zonesInClusters = 0;

            try
            {
                _logger("[WORKFLOW][CLUSTER] 🧬 Starting Clustering Sequence...");
                
                // ✅ CRASH SAFETY: Check timeout at start of clustering
                if (timeoutMonitor?.CheckTotalTimeout() == true)
                {
                    _logger("[WORKFLOW][CLUSTER] ⏱️ TIMEOUT: Skipping clustering sequence");
                    return (0, 0);
                }

                using (var db = _contextFactory())
                {
                    var repo = new ClashZoneRepository(db, _logger, _perf as PerformanceMonitor);

                    // A. PROXIMITY MARKING — run for ALL categories (not just one)
                    //    A single filter can have multiple combos spanning Ducts, Pipes, Cable Trays, etc.
                    //    CheckProximityToOtherZones already skips cross-category comparisons (line 178-181),
                    //    so passing all categories at once is safe and correct.
                    int markedForClustering = 0;
                    using (var proximityTracker = parentTracker?.TrackSubOperation("3a. Proximity Marking"))
                    {
                        try
                        {
                            // Default proximity tolerance: 0.5 feet (approximately 6 inches)
                            double proximityTolerance = 0.5;
                            var firstFilter = filters?.FirstOrDefault();
                            if (firstFilter?.Parameters != null && firstFilter.Parameters.ContainsKey("JoinOpeningsDistance"))
                            {
                                if (double.TryParse(firstFilter.Parameters["JoinOpeningsDistance"].ToString(), out double customTol))
                                {
                                    proximityTolerance = customTol;
                                }
                            }

                            var proximityMarker = new ClusterProximityMarker(repo, proximityTolerance, _logger);

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 3A: STARTING PROXIMITY CHECK\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Tolerance: {proximityTolerance} ft\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Mode: ALL categories (single filter spans multiple combos/categories)\n");

                            // ✅ FIX: Pass null to check ALL resolved zones regardless of category.
                            // GetZonesReadyForProximityCheck(null) returns zones across all categories.
                            // CheckProximityToOtherZones skips cross-category pairs automatically.
                            markedForClustering = proximityMarker.MarkZonesForClustering(_doc, null);

                            _logger($"[WORKFLOW][PROXIMITY] ✅ Marked {markedForClustering} zones for clustering (all categories)");
                        }
                        catch (Exception proximityEx)
                        {
                            _logger($"[WORKFLOW][PROXIMITY] ⚠️ Proximity marking failed: {proximityEx.Message}");

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
                        var algo = new ClusterAlgorithmService();
                        var rotation = new ClusterRotationService(
                            getClashZoneFunc: (id, _) => repo.GetClashZoneByInstanceId((int)id),
                            getClusterPlacementFunc: null
                        );
                        var calcService = new BatchClusterCalculationService(algo, rotation, db.DatabasePath);

                        // Load zones marked for clustering (Ready=1, Current=1, IsResolved=1, MarkedForClusterProcess=1)
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}] STEP 5: LOADING ZONES FOR CLUSTER CALCULATION\n");
                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}]   Query: Ready=1, Current=1, IsResolved=1, MarkedForClusterProcess=1\n");

                        var zones = repo.GetZonesForClustering();
                        zonesInClusters = zones.Count;

                        SafeFileLogger.SafeAppendText("flag_workflow.log",
                            $"[{DateTime.Now:HH:mm:ss}]   Found {zones.Count} zones eligible for clustering\n");

                        if (zones.Count > 0)
                        {
                            // ✅ CRASH SAFETY: Check timeout before cluster calculation
                            if (timeoutMonitor?.CheckTotalTimeout() == true)
                            {
                                _logger("[WORKFLOW][CLUSTER] ⏱️ TIMEOUT: Skipping cluster calculation");
                                return (clustersPlaced, zonesInClusters);
                            }
                            
                            _logger($"[WORKFLOW][CLUSTER] Calculating clusters for {zones.Count} zones...");

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}] STEP 6: CALCULATING CLUSTER ARRANGEMENTS\n");

                            // ✅ FIX: Pass "All" as category since zones span multiple categories
                            var results = calcService.CalculateOnly(zones, "All", 0, 0, _doc);

                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Calculated {results.Count} cluster arrangements\n");

                            if (results.Count > 0)
                            {
                                calcService.BatchSave(results);
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
                        // ✅ CRASH SAFETY: Check timeout before cluster placement
                        if (timeoutMonitor?.CheckTotalTimeout() == true)
                        {
                            _logger("[WORKFLOW][CLUSTER] ⏱️ TIMEOUT: Skipping cluster placement");
                            return (clustersPlaced, zonesInClusters);
                        }
                        
                        var parameterService = new SleeveParameterService(_doc, false, _perf as PlacementPerformanceMonitor);
                        var cleanup = new ClusterCleanupService();
                        // ✅ Instantiate snapshot transfer service for cluster placement
                        var clusterSnapshotTransferService = new ParameterSnapshotTransferService();
                        var placeService = new BatchClusterPlacementService(db.DatabasePath, repo, parameterService, cleanup, _perf, clusterSnapshotTransferService);

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
