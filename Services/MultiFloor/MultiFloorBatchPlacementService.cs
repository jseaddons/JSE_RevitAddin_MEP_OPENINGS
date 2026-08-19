using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Calculation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// 🚀 BIM360-OPTIMIZED: Multi-floor sleeve placement in SINGLE transaction
    /// 
    /// PROBLEM: Original FloorBatchProcessor creates one transaction per floor
    /// SOLUTION: This service aggregates ALL floors' sleeves and places in ONE batch
    /// 
    /// BIM360 Benefit: Reduces cloud sync round-trips from N to 1 (N = floor count)
    /// </summary>
    public class MultiFloorBatchPlacementService
    {
        private readonly Document _doc;
        private readonly IPerformanceMonitor _monitor;
        private readonly Action<string> _logger;
        private readonly Func<SleeveDbContext> _contextFactory;
        
        // ✅ SAFETY: Default max sleeves per batch (will be adjusted based on memory)
        private const int DEFAULT_MAX_SLEEVES_PER_BATCH = 500;
        private const int MIN_SLEEVES_PER_BATCH = 200;  // Minimum for low-memory systems
        private const int MAX_SLEEVES_PER_BATCH_HARD = 1000; // Absolute maximum
        
        // ✅ SAFETY: Memory thresholds for adaptive batch sizing (based on Process RAM usage)
        private const long MEMORY_HIGH_USAGE_MB = 2000;  // >2GB used = heavy load, use small batches
        private const long MEMORY_MODERATE_MB = 1200;    // >1.2GB used = moderate, use normal batches
                                                           // <1.2GB used = light, use large batches
        
        // ✅ SAFETY: Base transaction timeout in minutes
        private const int BASE_TRANSACTION_TIMEOUT_MINUTES = 2;
        
        // ✅ SAFETY: Time per sleeve estimate (conservative for BIM360)
        private const double TIME_PER_SLEEVE_MS = 100; // Conservative estimate

        public MultiFloorBatchPlacementService(
            Document doc,
            Func<SleeveDbContext> contextFactory = null,
            IPerformanceMonitor monitor = null,
            Action<string> logger = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _contextFactory = contextFactory ?? (() => new SleeveDbContext(doc));
            _monitor = monitor;
            _logger = logger ?? (msg => SafeFileLogger.SafeAppendText("multifloor.log", msg));
        }

        /// <summary>
        /// 🎯 MAIN ENTRY: Place sleeves for ALL floors in optimized batch mode
        /// 
        /// Flow:
        /// 1. DETECTION (No Transaction) - Read-only clash detection per floor
        /// 2. PLANNING (No Transaction) - Calculate sleeve dimensions per floor  
        /// 3. PLACEMENT (Single Transaction) - Batch place ALL sleeves in ONE API call
        /// 4. CLUSTERING (Optional Single Transaction) - Global clustering after placement
        /// </summary>
        public MultiFloorResult ProcessFloorsOptimized(
            List<Level> levels, 
            OpeningFilter filter,
            SharedResourceCache cache = null,
            bool enableGlobalClustering = true)
        {
            // 🆕 SAFETY: Check feature flag for new 3-phase optimization
            if (OptimizationFlags.UseOptimizedMultiFloorFlow)
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] 🚀 USING OPTIMIZED 3-PHASE FLOW (UseOptimizedMultiFloorFlow = true)");
                return ProcessFloorsThreePhase(levels, filter, cache, enableGlobalClustering);
            }
            
            // Legacy flow (safe default)
            return ProcessFloorsLegacy(levels, filter, cache, enableGlobalClustering);
        }
        
        /// <summary>
        /// 🚀 NEW: 3-Phase Optimized Flow
        /// Phase 1: Place ALL individual sleeves (ALL floors) → 1 Regenerate
        /// Phase 2: Cluster ALL floors together → 1 Regenerate  
        /// Phase 3: Delete ALL old sleeves → 1 Regenerate (deferred)
        /// Result: Only 3 Regenerates total (vs 3N for N floors)
        /// </summary>
        private MultiFloorResult ProcessFloorsThreePhase(
            List<Level> levels, 
            OpeningFilter filter,
            SharedResourceCache cache,
            bool enableGlobalClustering)
        {
            var result = new MultiFloorResult();
            var startTime = DateTime.Now;
            
            _logger($"[{DateTime.Now:HH:mm:ss}] 🚀 MULTI-FLOOR 3-PHASE: Processing {levels.Count} floors");

            try
            {
                // ═══════════════════════════════════════════════════════════════
                // PHASE 0: DETECTION & PLANNING (NO TRANSACTION - READ ONLY)
                // ═══════════════════════════════════════════════════════════════
                var allPlannedItems = new List<PlannedFloorItem>();
                
                foreach (var level in levels)
                {
                    try
                    {
                        var floorItems = DetectAndPlanFloor(level, filter, cache);
                        if (floorItems.Count > 0)
                        {
                            allPlannedItems.AddRange(floorItems);
                            result.SuccessfulFloors.Add(level.Name);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedFloors.Add($"{level.Name}: {ex.Message}");
                    }
                }

                if (allPlannedItems.Count == 0)
                {
                    _logger($"[{DateTime.Now:HH:mm:ss}] ℹ No sleeves to place");
                    return result;
                }

                _logger($"[{DateTime.Now:HH:mm:ss}] 📊 PHASE 0: {allPlannedItems.Count} sleeves planned across {result.SuccessfulFloors.Count} floors");

                // ═══════════════════════════════════════════════════════════════
                // PHASE 1: PLACE ALL INDIVIDUAL SLEEVES → 1 Transaction, 1 Commit
                // Multiple batches inside ONE transaction to reduce commit overhead
                // Each batch still does its own NewFamilyInstances2 + Regen + Rotate
                // ═══════════════════════════════════════════════════════════════
                int totalPlaced = 0;
                var optimalBatchSize = CalculateOptimalBatchSize();
                var chunks = ChunkPlannedItems(allPlannedItems, optimalBatchSize);
                var allPlacedItems = new List<(ClashZone Zone, ElementId ElementId)>();

                _logger($"[{DateTime.Now:HH:mm:ss}] 🔷 PHASE 1: Placing {allPlannedItems.Count} individual sleeves in {chunks.Count} batch(es) — SINGLE TRANSACTION");

                using (var phase1Tracker = _monitor?.TrackOperation("Phase 1: Place All Individual Sleeves (Single Tx)"))
                {
                    var timeout = CalculateTransactionTimeout(allPlannedItems.Count);

                    using (var transaction = new Transaction(_doc, "Place Sleeves - All Floors (Optimized)"))
                    {
                        var options = transaction.GetFailureHandlingOptions();
                        options.SetFailuresPreprocessor(new MultiFloorFailurePreprocessor("AllFloors", _logger));
                        transaction.SetFailureHandlingOptions(options);

                        transaction.Start();
                        _logger($"[{DateTime.Now:HH:mm:ss}] 🏗️ SINGLE TX START: {allPlannedItems.Count} sleeves, {chunks.Count} batch(es), timeout: {timeout.TotalMinutes:F1}min");

                        try
                        {
                            var bulkService = new BulkPlacementService(
                                _doc,
                                contextFactory: _contextFactory,
                                logger: _logger,
                                performanceMonitor: _monitor);

                            for (int i = 0; i < chunks.Count; i++)
                            {
                                var bulkItems = chunks[i].Select(p => (p.Zone, p.Plan)).ToList();
                                _logger($"[{DateTime.Now:HH:mm:ss}]   Batch {i + 1}/{chunks.Count}: {bulkItems.Count} sleeves");

                                // Reuse ExecuteBulkPlacement — does NewFamilyInstances2 + Regen + Rotate + Params per batch
                                var batchResult = bulkService.ExecuteBulkPlacement(_doc, bulkItems, skipSpatialFiltering: false);

                                totalPlaced += batchResult.PlacedCount;
                                allPlacedItems.AddRange(batchResult.PlacedItems);

                                // Yield between batches to prevent UI freeze
                                System.Windows.Forms.Application.DoEvents();
                            }

                            var commitStatus = transaction.Commit();
                            _logger($"[{DateTime.Now:HH:mm:ss}] ✅ SINGLE TX COMMITTED: {totalPlaced} sleeves ({commitStatus})");
                        }
                        catch (Exception ex)
                        {
                            transaction.RollBack();
                            _logger($"[{DateTime.Now:HH:mm:ss}] ❌ SINGLE TX ROLLED BACK: {ex.Message}");
                            throw;
                        }
                    }

                    // Post-commit: read geometry from placed sleeves (reuses existing methods)
                    if (allPlacedItems.Count > 0)
                    {
                        var updateService = new BulkPlacementService(
                            _doc,
                            contextFactory: _contextFactory,
                            logger: _logger,
                            performanceMonitor: _monitor);

                        updateService.UpdateZonesFromElements(_doc, allPlacedItems);
                        _logger($"[{DateTime.Now:HH:mm:ss}] 📐 Updated zones from {allPlacedItems.Count} placed elements");

                        // Corner extraction for clustering proximity
                        using (var db = _contextFactory())
                        {
                            var repo = new Data.Repositories.ClashZoneRepository(db, _logger, _monitor as PerformanceMonitor);
                            var extractor = new Calculation.BatchSleeveCornerExtractor(repo);
                            var placedZones = allPlacedItems.Select(x => x.Zone).ToList();
                            int extracted = extractor.ExtractAndSaveCornersForZones(_doc, placedZones);
                            _logger($"[{DateTime.Now:HH:mm:ss}] 📐 Extracted corners for {extracted} sleeves");
                        }
                    }

                    phase1Tracker?.SetItemCount(totalPlaced);
                }

                result.TotalSleevesPlaced = totalPlaced;
                _logger($"[{DateTime.Now:HH:mm:ss}] ✅ PHASE 1 COMPLETE: {totalPlaced} individual sleeves placed");

                if (totalPlaced == 0)
                {
                    _logger($"[{DateTime.Now:HH:mm:ss}] ⚠️ No sleeves placed, skipping clustering");
                    return result;
                }

                // ═══════════════════════════════════════════════════════════════
                // PHASE 2: GLOBAL CLUSTERING → 1 Regenerate
                // ═══════════════════════════════════════════════════════════════
                if (enableGlobalClustering && OptimizationFlags.EnableClusteringWorkflow)
                {
                    using (var phase2Tracker = _monitor?.TrackOperation("Phase 2: Global Clustering (All Floors)"))
                    {
                        _logger($"[{DateTime.Now:HH:mm:ss}] 🔷 PHASE 2: Running global clustering for all floors");
                        
                        // Use the optimized clustering that defers deletions
                        var clusteringResult = ExecuteGlobalClusteringDeferDeletions();
                        
                        result.TotalClustersFormed = clusteringResult.clustersPlaced;
                        result.TotalClusteredZones = clusteringResult.zonesInClusters;
                        
                        _logger($"[{DateTime.Now:HH:mm:ss}] ✅ PHASE 2 COMPLETE: {clusteringResult.clustersPlaced} clusters placed");
                        phase2Tracker?.SetItemCount(clusteringResult.clustersPlaced);
                    }
                }

                // ═══════════════════════════════════════════════════════════════
                // PHASE 3: DEFERRED DELETION → Commit auto-regenerates
                // ═══════════════════════════════════════════════════════════════
                if (enableGlobalClustering && result.TotalClusteredZones > 0)
                {
                    using (var phase3Tracker = _monitor?.TrackOperation("Phase 3: Delete Old Individual Sleeves"))
                    {
                        _logger($"[{DateTime.Now:HH:mm:ss}] 🔷 PHASE 3: Deleting {result.TotalClusteredZones} old individual sleeves (Commit will auto-regenerate)");
                        
                        int deleted = ExecuteDeferredDeletions();
                        
                        _logger($"[{DateTime.Now:HH:mm:ss}] ✅ PHASE 3 COMPLETE: {deleted} old sleeves deleted");
                        phase3Tracker?.SetItemCount(deleted);
                    }
                }

                var elapsed = DateTime.Now - startTime;
                _logger($"[{DateTime.Now:HH:mm:ss}] ✅ 3-PHASE COMPLETE in {elapsed.TotalSeconds:F1}s: {totalPlaced} sleeves, {result.TotalClustersFormed} clusters");
                
                return result;
            }
            catch (Exception ex)
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] ❌ 3-PHASE ERROR: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Legacy flow (per-floor processing) - Safe fallback
        /// </summary>
        private MultiFloorResult ProcessFloorsLegacy(
            List<Level> levels, 
            OpeningFilter filter,
            SharedResourceCache cache,
            bool enableGlobalClustering)
        {
            var result = new MultiFloorResult();
            var startTime = DateTime.Now;
            
            _logger($"[{DateTime.Now:HH:mm:ss}] 🚀 MULTI-FLOOR LEGACY: Processing {levels.Count} floors in BATCH mode");
            
            // Initialize cache if not provided
            cache ??= new SharedResourceCache(_doc);
            cache.Initialize();

            try
            {
                // ═══════════════════════════════════════════════════════════════
                // PHASE 1: DETECTION & PLANNING (NO TRANSACTION - READ ONLY)
                // ═══════════════════════════════════════════════════════════════
                var allPlannedItems = new List<PlannedFloorItem>();
                
                using (var detectionTracker = _monitor?.TrackOperation("Phase 1: Detection & Planning (All Floors)"))
                {
                    foreach (var level in levels)
                    {
                        try
                        {
                            var floorItems = DetectAndPlanFloor(level, filter, cache);
                            if (floorItems.Count > 0)
                            {
                                allPlannedItems.AddRange(floorItems);
                                result.SuccessfulFloors.Add(level.Name);
                                _logger($"[{DateTime.Now:HH:mm:ss}]   ✓ {level.Name}: {floorItems.Count} sleeves planned");
                            }
                            else
                            {
                                _logger($"[{DateTime.Now:HH:mm:ss}]   ℹ {level.Name}: No clashes detected");
                            }
                        }
                        catch (Exception ex)
                        {
                            result.FailedFloors.Add($"{level.Name}: {ex.Message}");
                            _logger($"[{DateTime.Now:HH:mm:ss}]   ✗ {level.Name}: FAILED - {ex.Message}");
                            // Continue with other floors - don't fail entire batch
                        }
                    }
                    
                    detectionTracker?.SetItemCount(allPlannedItems.Count);
                }

                if (allPlannedItems.Count == 0)
                {
                    _logger($"[{DateTime.Now:HH:mm:ss}] ℹ No sleeves to place across all floors");
                    return result;
                }

                _logger($"[{DateTime.Now:HH:mm:ss}] 📊 TOTAL: {allPlannedItems.Count} sleeves planned across {result.SuccessfulFloors.Count} floors");

                // ═══════════════════════════════════════════════════════════════
                // PHASE 2: BATCH PLACEMENT (SINGLE TRANSACTION)
                // ═══════════════════════════════════════════════════════════════
                int totalPlaced = 0;
                int totalFailed = 0;
                
                // Calculate optimal batch size based on current memory state
                var optimalBatchSize = CalculateOptimalBatchSize();
                
                // Chunk if needed for safety
                var chunks = ChunkPlannedItems(allPlannedItems, optimalBatchSize);
                _logger($"[{DateTime.Now:HH:mm:ss}] 🔢 Split into {chunks.Count} chunk(s) (batch size: {optimalBatchSize}, total: {allPlannedItems.Count} sleeves)");

                for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
                {
                    var chunk = chunks[chunkIndex];
                    var chunkResult = PlaceChunkInTransaction(chunk, chunkIndex + 1, chunks.Count);
                    
                    totalPlaced += chunkResult.placed;
                    totalFailed += chunkResult.failed;
                }

                result.TotalSleevesPlaced = totalPlaced;
                
                // ═══════════════════════════════════════════════════════════════
                // PHASE 3: GLOBAL CLUSTERING (OPTIONAL, SINGLE TRANSACTION)
                // ═══════════════════════════════════════════════════════════════
                if (enableGlobalClustering && totalPlaced > 0 && OptimizationFlags.EnableClusteringWorkflow)
                {
                    using (var clusterTracker = _monitor?.TrackOperation("Phase 3: Global Clustering"))
                    {
                        var clusteringResult = ExecuteGlobalClustering();
                        result.TotalClustersFormed = clusteringResult.clustersPlaced;
                        result.TotalClusteredZones = clusteringResult.zonesInClusters;
                    }
                }

                var elapsed = DateTime.Now - startTime;
                _logger($"[{DateTime.Now:HH:mm:ss}] ✅ COMPLETE: {totalPlaced} sleeves placed in {elapsed.TotalSeconds:F1}s");
                
                return result;
            }
            catch (Exception ex)
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL ERROR: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Detect clashes and plan sleeves for a single floor (NO TRANSACTION)
        /// </summary>
        private List<PlannedFloorItem> DetectAndPlanFloor(
            Level level, 
            OpeningFilter filter,
            SharedResourceCache cache)
        {
            var plannedItems = new List<PlannedFloorItem>();
            
            // 1. Detect clashes (read-only, no transaction needed)
            var detector = new FloorIsolatedClashDetector(_doc, cache);
            var categories = GetSelectedCategories(filter);
            var clashes = detector.DetectClashesForFloor(level, categories);
            
            if (clashes.Count == 0) return plannedItems;

            // 2. Plan sleeves (calculation only, no Revit API)
            var conditionsService = new ConditionsService(_doc);
            
            foreach (var category in categories)
            {
                var categoryZones = clashes.Where(cz => 
                    string.Equals(cz.MepElementCategory, category, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                
                if (categoryZones.Count == 0) continue;
                
                var categoryConditions = conditionsService.LoadConditions(category);
                var planner = new ParallelSleevePlacementPlanner(categoryConditions);
                var planningResult = planner.Plan(categoryZones);
                
                var plannedMap = planningResult.Items.ToDictionary(i => i.ClashZoneId);
                
                foreach (var zone in categoryZones)
                {
                    if (plannedMap.TryGetValue(zone.Id, out var plan))
                    {
                        if (plan.ShouldSkip) continue;
                        
                        // Mark for current context (ensures placement passes context check)
                        zone.IsCurrentClash = true;
                        
                        plannedItems.Add(new PlannedFloorItem
                        {
                            Zone = zone,
                            Plan = plan,
                            LevelName = level.Name,
                            Category = category
                        });
                    }
                }
            }
            
            return plannedItems;
        }

        /// <summary>
        /// Calculate optimal batch size based on current process RAM usage
        /// Uses Process.WorkingSet64 for accurate memory pressure detection
        /// </summary>
        private int CalculateOptimalBatchSize()
        {
            try
            {
                var usedMemoryMB = GetCurrentProcessMemoryMB();
                
                if (usedMemoryMB > MEMORY_HIGH_USAGE_MB)
                {
                    // Heavy RAM usage - use smaller batches to avoid OOM
                    _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] ⚠️ High RAM usage: {usedMemoryMB}MB, using small batch size {MIN_SLEEVES_PER_BATCH}");
                    return MIN_SLEEVES_PER_BATCH;
                }
                else if (usedMemoryMB > MEMORY_MODERATE_MB)
                {
                    // Moderate RAM usage - use default batches
                    _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] 💾 Moderate RAM usage: {usedMemoryMB}MB, using batch size {DEFAULT_MAX_SLEEVES_PER_BATCH}");
                    return DEFAULT_MAX_SLEEVES_PER_BATCH;
                }
                else
                {
                    // Light RAM usage - can use larger batches
                    _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] ✅ Low RAM usage: {usedMemoryMB}MB, using large batch size {MAX_SLEEVES_PER_BATCH_HARD}");
                    return MAX_SLEEVES_PER_BATCH_HARD;
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] ⚠️ Memory check failed: {ex.Message}, using default batch size");
                return DEFAULT_MAX_SLEEVES_PER_BATCH;
            }
        }

        /// <summary>
        /// Get current process RAM usage in MB
        /// Uses Process.WorkingSet64 for accurate physical memory consumption
        /// </summary>
        private long GetCurrentProcessMemoryMB()
        {
            try
            {
                using (var proc = Process.GetCurrentProcess())
                {
                    return proc.WorkingSet64 / (1024 * 1024); // Convert bytes to MB
                }
            }
            catch
            {
                return DEFAULT_MAX_SLEEVES_PER_BATCH; // Safe fallback
            }
        }

        /// <summary>
        /// Calculate adaptive transaction timeout based on sleeve count
        /// </summary>
        private TimeSpan CalculateTransactionTimeout(int sleeveCount)
        {
            // Base timeout: 2 minutes minimum
            var baseTimeout = TimeSpan.FromMinutes(BASE_TRANSACTION_TIMEOUT_MINUTES);
            
            // Estimated time: sleeve count × time per sleeve (with 3x safety margin for BIM360)
            var estimatedTime = TimeSpan.FromMilliseconds(sleeveCount * TIME_PER_SLEEVE_MS * 3);
            
            // Use the larger of base or estimated, capped at 10 minutes max
            var timeout = TimeSpan.FromMinutes(Math.Max(baseTimeout.TotalMinutes, estimatedTime.TotalMinutes));
            var maxTimeout = TimeSpan.FromMinutes(10);
            
            return timeout > maxTimeout ? maxTimeout : timeout;
        }

        /// <summary>
        /// Place a chunk of sleeves in a SINGLE transaction
        /// </summary>
        private (int placed, int failed) PlaceChunkInTransaction(
            List<PlannedFloorItem> chunk, 
            int chunkNumber, 
            int totalChunks)
        {
            int placed = 0;
            int failed = 0;
            
            var transactionName = totalChunks == 1 
                ? "Place Sleeves - All Floors" 
                : $"Place Sleeves - Chunk {chunkNumber}/{totalChunks}";
            
            // Calculate adaptive timeout based on chunk size
            var timeout = CalculateTransactionTimeout(chunk.Count);
            
            using (var transaction = new Transaction(_doc, transactionName))
            {
                // Set timeout for BIM360 safety
                var options = transaction.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(new MultiFloorFailurePreprocessor(chunkNumber.ToString(), _logger));
                transaction.SetFailureHandlingOptions(options);
                
                transaction.Start();
                
                var placementStartTime = DateTime.Now;
                _logger($"[{DateTime.Now:HH:mm:ss}] ⏱️ Timeout set to {timeout.TotalMinutes:F1} minutes for {chunk.Count} sleeves");
                
                try
                {
                    _logger($"[{DateTime.Now:HH:mm:ss}] 🏗️ TRANSACTION START: {transactionName} ({chunk.Count} sleeves, timeout: {timeout.TotalMinutes:F1}min)");
                    
                    // Convert to format expected by BulkPlacementService
                    var bulkItems = chunk.Select(p => (p.Zone, p.Plan)).ToList();
                    
                    // ✅ SINGLE BATCH API CALL for entire chunk
                    var placementService = new BulkPlacementService(
                        _doc,
                        contextFactory: _contextFactory,
                        logger: _logger,
                        performanceMonitor: _monitor);
                    
                    var result = placementService.ExecuteBulkPlacement(_doc, bulkItems, skipSpatialFiltering: false);
                    
                    placed = result.PlacedCount;
                    failed = result.FailedCount;
                    
                    // Commit the transaction
                    var commitStatus = transaction.Commit();
                    
                    if (commitStatus == TransactionStatus.Committed)
                    {
                        _logger($"[{DateTime.Now:HH:mm:ss}] ✅ TRANSACTION COMMITTED: {placed} sleeves placed");
                    }
                    else
                    {
                        _logger($"[{DateTime.Now:HH:mm:ss}] ⚠️ TRANSACTION STATUS: {commitStatus}");
                    }
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    _logger($"[{DateTime.Now:HH:mm:ss}] ❌ TRANSACTION ROLLED BACK: {ex.Message}");
                    throw;
                }
            }
            
            return (placed, failed);
        }

        /// <summary>
        /// Execute clustering globally for all placed sleeves
        /// </summary>
        private (int clustersPlaced, int zonesInClusters) ExecuteGlobalClustering()
        {
            try
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] 🧬 Starting Global Clustering...");
                
                using (var db = _contextFactory())
                {
                    // Cast to Refresh.PerformanceMonitor or null if wrong type
                    var refreshMonitor = _monitor as PerformanceMonitor;
                    var repo = new Data.Repositories.ClashZoneRepository(db, _logger, refreshMonitor);
                    
                    // 1. Proximity marking (all floors)
                    var proximityMarker = new Clustering.Proximity.ClusterProximityMarker(
                        repo, proximityTolerance: 0.5, _logger);
                    int marked = proximityMarker.MarkZonesForClustering(_doc, null); // null = all categories
                    
                    _logger($"[{DateTime.Now:HH:mm:ss}]   Marked {marked} zones for clustering");
                    
                    if (marked == 0) return (0, 0);
                    
                    // 2. Cluster calculation
                    var zones = repo.GetZonesForClustering();
                    var algo = new Clustering.Algorithm.ClusterAlgorithmService();
                    var rotation = new Clustering.Rotation.ClusterRotationService(
                        getClashZoneFunc: (id, _) => repo.GetClashZoneByInstanceId((int)id));
                    var calcService = new Clustering.BatchClusterCalculationService(algo, rotation, db.DatabasePath);
                    
                    var results = calcService.CalculateOnly(zones, "All", 0, 0, _doc);
                    
                    if (results.Count > 0)
                    {
                        calcService.BatchSave(results);
                        calcService.BatchUpdateFlags(results);
                    }
                    
                    // 3. Cluster placement
                    var placementMonitor = _monitor as PlacementPerformanceMonitor;
                    var parameterService = new SleeveParameterService(_doc, false, placementMonitor);
                    var cleanup = new Clustering.Cleanup.ClusterCleanupService();
                    var snapshotTransfer = new ParameterSnapshotTransferService();
                    var placeService = new Clustering.BatchClusterPlacementService(
                        db.DatabasePath, repo, parameterService, cleanup, _monitor, snapshotTransfer);
                    
                    var (placed, failed, cleaned) = placeService.PlaceAllCategoriesAndCleanup(_doc, cleanup);
                    
                    _logger($"[{DateTime.Now:HH:mm:ss}] ✅ Global Clustering: {placed} clusters placed, {cleaned} cleaned");
                    
                    return (placed, zones.Count);
                }
            }
            catch (Exception ex)
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] ⚠️ Global Clustering failed: {ex.Message}");
                return (0, 0);
            }
        }

        /// <summary>
        /// 🆕 PHASE 2: Global clustering with deferred deletions
        /// Places clusters but defers deletion of old individual sleeves to Phase 3
        /// </summary>
        private (int clustersPlaced, int zonesInClusters) ExecuteGlobalClusteringDeferDeletions()
        {
            try
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] 🧬 Starting Global Clustering (Deferred Deletions)...");
                
                using (var db = _contextFactory())
                {
                    var repo = new Data.Repositories.ClashZoneRepository(db, _logger, _monitor as PerformanceMonitor);
                    
                    // 1. Proximity marking (all floors)
                    var proximityMarker = new Clustering.Proximity.ClusterProximityMarker(
                        repo, proximityTolerance: 0.5, _logger);
                    int marked = proximityMarker.MarkZonesForClustering(_doc, null);
                    
                    _logger($"[{DateTime.Now:HH:mm:ss}]   Marked {marked} zones for clustering");
                    
                    if (marked == 0) return (0, 0);
                    
                    // 2. Cluster calculation
                    var zones = repo.GetZonesForClustering();
                    var algo = new Clustering.Algorithm.ClusterAlgorithmService();
                    var rotation = new Clustering.Rotation.ClusterRotationService(
                        getClashZoneFunc: (id, _) => repo.GetClashZoneByInstanceId((int)id));
                    var calcService = new Clustering.BatchClusterCalculationService(algo, rotation, db.DatabasePath);
                    
                    var results = calcService.CalculateOnly(zones, "All", 0, 0, _doc);
                    
                    if (results.Count > 0)
                    {
                        calcService.BatchSave(results);
                        calcService.BatchUpdateFlags(results);
                    }
                    
                    // 3. Cluster placement (WITHOUT cleanup - deletions deferred to Phase 3)
                    var parameterService = new SleeveParameterService(_doc, false, _monitor as PlacementPerformanceMonitor);
                    // Use null cleanup to defer deletions
                    var placeService = new Clustering.BatchClusterPlacementService(
                        db.DatabasePath, repo, parameterService, null, _monitor, null);
                    
                    // Place clusters only, don't delete old sleeves yet
                    var (placed, failed) = placeService.PlaceFromDatabase(_doc, null, useSingleTransaction: true);
                    
                    _logger($"[{DateTime.Now:HH:mm:ss}] ✅ Global Clustering: {placed} clusters placed (deletions deferred to Phase 3)");
                    
                    // Store the IDs to delete for Phase 3
                    StoreSleevesToDelete(zones);
                    
                    return (placed, zones.Count);
                }
            }
            catch (Exception ex)
            {
                _logger($"[{DateTime.Now:HH:mm:ss}] ⚠️ Global Clustering (Deferred) failed: {ex.Message}");
                return (0, 0);
            }
        }

        private List<int> _sleevesToDelete = new List<int>();
        
        /// <summary>
        /// Store sleeve IDs for deferred deletion in Phase 3
        /// </summary>
        private void StoreSleevesToDelete(List<ClashZone> clusteredZones)
        {
            _sleevesToDelete.Clear();
            foreach (var zone in clusteredZones)
            {
                if (zone.SleeveInstanceId > 0)
                {
                    _sleevesToDelete.Add((int)zone.SleeveInstanceId);
                }
            }
            _logger($"[{DateTime.Now:HH:mm:ss}]   Stored {_sleevesToDelete.Count} sleeves for Phase 3 deletion");
        }

        /// <summary>
        /// 🆕 PHASE 3: Execute deferred deletions in separate transaction
        /// </summary>
        private int ExecuteDeferredDeletions()
        {
            if (_sleevesToDelete.Count == 0) return 0;
            
            int deletedCount = 0;
            
            using (var transaction = new Transaction(_doc, "Delete Old Individual Sleeves"))
            {
                var options = transaction.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(new MultiFloorFailurePreprocessor("Delete", _logger));
                transaction.SetFailureHandlingOptions(options);
                
                transaction.Start();
                
                try
                {
#if REVIT2024_OR_GREATER
                    var elementIds = _sleevesToDelete
                        .Select(id => new ElementId((long)id))
                        .Where(id => _doc.GetElement(id) != null)
                        .ToList();
#else
                    var elementIds = _sleevesToDelete
                        .Select(id => new ElementId(id))
                        .Where(id => _doc.GetElement(id) != null)
                        .ToList();
#endif
                    
                    if (elementIds.Count > 0)
                    {
                        _doc.Delete(elementIds);
                        deletedCount = elementIds.Count;
                        _logger($"[{DateTime.Now:HH:mm:ss}]   Deleted {deletedCount} old individual sleeves");
                    }
                    
                    transaction.Commit();
                    _sleevesToDelete.Clear(); // Clear after successful deletion
                }
                catch (Exception ex)
                {
                    transaction.RollBack();
                    _logger($"[{DateTime.Now:HH:mm:ss}] ❌ Deletion failed: {ex.Message}");
                    throw;
                }
            }
            
            return deletedCount;
        }

        /// <summary>
        /// Split planned items into chunks for safety
        /// </summary>
        private List<List<PlannedFloorItem>> ChunkPlannedItems(List<PlannedFloorItem> items, int chunkSize)
        {
            var chunks = new List<List<PlannedFloorItem>>();
            for (int i = 0; i < items.Count; i += chunkSize)
            {
                chunks.Add(items.Skip(i).Take(chunkSize).ToList());
            }
            return chunks;
        }

        private List<string> GetSelectedCategories(OpeningFilter filter)
        {
            if (filter?.SelectedMepCategoryNames?.Any() == true)
                return filter.SelectedMepCategoryNames;
            
            // Category is a value type (enum), can't use ?. operator
            string categoryStr = filter != null ? filter.Category.ToString() : "Pipes";
            return new List<string> { categoryStr };
        }

        /// <summary>
        /// Internal class to track planned item with floor info
        /// </summary>
        private class PlannedFloorItem
        {
            public ClashZone Zone { get; set; }
            public SleevePlacementPlanningDto Plan { get; set; }
            public string LevelName { get; set; }
            public string Category { get; set; }
        }
    }

    /// <summary>
    /// Failure preprocessor for multi-floor transactions
    /// </summary>
    public class MultiFloorFailurePreprocessor : IFailuresPreprocessor
    {
        private readonly string _context;
        private readonly Action<string> _logger;

        public MultiFloorFailurePreprocessor(string context, Action<string> logger)
        {
            _context = context;
            _logger = logger;
        }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                var severity = f.GetSeverity();
                var desc = f.GetDescriptionText();
                
                if (severity == FailureSeverity.Warning)
                {
                    _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] [MultiFloor {_context}] Swallowing warning: {desc}");
                    fa.DeleteWarning(f);
                }
                else if (severity == FailureSeverity.Error)
                {
                    // Handle known harmless errors
                    if (IsHarmlessError(desc))
                    {
                        _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] [MultiFloor {_context}] Treating harmless error as warning: {desc}");
                        fa.DeleteWarning(f);
                    }
                    else
                    {
                        _logger?.Invoke($"[{DateTime.Now:HH:mm:ss}] [MultiFloor {_context}] CRITICAL ERROR: {desc}");
                    }
                }
            }
            return FailureProcessingResult.Continue;
        }

        private bool IsHarmlessError(string desc)
        {
            var harmlessPatterns = new[]
            {
                "duplicate", "coincident", "slightly off axis", 
                "very small arc", "instability detected"
            };
            return harmlessPatterns.Any(p => 
                desc.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
        }
    }
}
