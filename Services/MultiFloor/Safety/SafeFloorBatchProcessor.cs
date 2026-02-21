using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor.Safety
{
    /// <summary>
    /// Crash-safe multi-floor batch processor with comprehensive error handling.
    /// Integrates timeout monitoring, validation, circuit breaker, and checkpoint recovery.
    /// </summary>
    public class SafeFloorBatchProcessor
    {
        private readonly Document _doc;
        private readonly IPerformanceMonitor _monitor;
        private readonly CheckpointManager _checkpointMgr;
        private readonly MultiFloorValidator _validator;
        private readonly MultiFloorTimeoutMonitor _timeoutMonitor;
        private readonly CircuitBreaker _circuitBreaker;
        private readonly int _maxParallelism;
        private readonly int _chunkSize;
        
        public SafeFloorBatchProcessor(
            Document doc,
            IPerformanceMonitor monitor = null,
            string checkpointPath = null,
            int floorTimeoutMinutes = 1,
            int totalTimeoutMinutes = 10,
            int chunkSize = 5,
            int circuitBreakerThreshold = 3)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _monitor = monitor;
            _chunkSize = chunkSize;
            _maxParallelism = Math.Min(Environment.ProcessorCount, 8);
            
            _checkpointMgr = new CheckpointManager(checkpointPath ?? GetDefaultCheckpointPath());
            _validator = new MultiFloorValidator(msg => Log(msg));
            _timeoutMonitor = new MultiFloorTimeoutMonitor(floorTimeoutMinutes, totalTimeoutMinutes, msg => Log(msg));
            _circuitBreaker = new CircuitBreaker(circuitBreakerThreshold);
            
            // Subscribe to timeout events
            _timeoutMonitor.FloorTimeout += (s, e) => Log($"⏱ Floor timeout: {e.Message}");
            _timeoutMonitor.TotalTimeout += (s, e) => Log($"⏱ Total timeout: {e.Message}");
        }
        
        /// <summary>
        /// Main entry point: Process multiple floors with crash protection
        /// </summary>
        public MultiFloorResult ProcessFloorsSafe(List<Level> levels, OpeningFilter filter)
        {
            Log($"🚀 Starting crash-safe multi-floor processing: {levels?.Count ?? 0} levels, chunk size {_chunkSize}, floor timeout {_timeoutMonitor.GetStats().FloorTimeoutMs/60000}min");
            
            // Pre-flight validation
            var preValidation = _validator.ValidateBeforeProcessing(levels, _doc);
            if (!preValidation.IsValid)
            {
                var errorMsg = $"Pre-flight validation failed: {string.Join("; ", preValidation.Errors)}";
                Log($"❌ {errorMsg}");
                throw new InvalidOperationException(errorMsg);
            }
            
            Log($"✅ Pre-validation passed: {preValidation.GetSummary()}");
            
            // Verify database connection - prompts user and throws if fails
            Log("🔍 Verifying database connection...");
            SleeveDbDiagnostics.VerifyConnectionOrPrompt(_doc);
            Log("✅ Database connection OK");
            
            // Start timeout monitoring
            _timeoutMonitor.StartOperation();
            
            // Check for checkpoint (resume capability)
            var checkpoint = _checkpointMgr.LoadCheckpoint();
            if (checkpoint != null)
            {
                Log($"🔄 Resuming from checkpoint: {checkpoint.ProcessedFloors.Count} floors already done");
                levels = levels.Where(l => !checkpoint.ProcessedFloors.Contains(l.Name)).ToList();
            }
            
            if (levels.Count == 0)
            {
                Log("✅ All floors already processed");
                return new MultiFloorResult();
            }
            
            // Pre-cache symbols ONCE for all floors
            SharedResourceCache cache;
            using (var cacheTracker = _monitor?.TrackOperation("Pre-load Shared Resources"))
            {
                cache = new SharedResourceCache(_doc);
                cache.Initialize();
                cacheTracker?.SetItemCount(cache.Symbols.Count);
                Log($"📦 Pre-cached {cache.Symbols.Count} symbols");
            }
            
            var totalResult = new MultiFloorResult();
            var failedFloors = new List<string>();
            int chunkIndex = 0;
            
            // Process in chunks for memory safety
            for (int i = 0; i < levels.Count; i += _chunkSize)
            {
                // Check total timeout
                if (_timeoutMonitor.CheckTotalTimeout())
                {
                    Log("⏱ TOTAL TIMEOUT: Stopping multi-floor processing");
                    SaveCheckpointWithErrors(totalResult, failedFloors);
                    throw new TimeoutException("Multi-floor processing exceeded total time limit");
                }
                
                // Check circuit breaker
                if (!_circuitBreaker.ShouldProcess())
                {
                    Log("🚫 CIRCUIT BREAKER: Too many consecutive failures, stopping");
                    SaveCheckpointWithErrors(totalResult, failedFloors);
                    throw new InvalidOperationException("Circuit breaker triggered due to repeated failures");
                }
                
                var chunk = levels.Skip(i).Take(_chunkSize).ToList();
                chunkIndex++;
                
                Log($"🚀 Processing chunk {chunkIndex}: Floors {string.Join(", ", chunk.Select(l => l.Name))}");
                
                using (var chunkTracker = _monitor?.TrackOperation($"Floor Chunk {chunkIndex}"))
                {
                    var chunkResult = ProcessFloorChunkSafe(chunk, filter, cache);
                    totalResult.Merge(chunkResult);
                    chunkTracker?.SetItemCount(chunk.Count);
                    
                    // Track failures for circuit breaker
                    var chunkFailedFloors = chunkResult.FailedFloors;
                    failedFloors.AddRange(chunkFailedFloors);
                    
                    if (chunkFailedFloors.Count == 0)
                    {
                        _circuitBreaker.RecordSuccess();
                    }
                    else
                    {
                        _circuitBreaker.RecordFailure(chunkFailedFloors.Count);
                        Log($"⚠️ Chunk {chunkIndex} had {chunkFailedFloors.Count} failures");
                    }
                    
                    // Save checkpoint after each chunk
                    SaveCheckpointWithErrors(totalResult, failedFloors);
                    
                    // Memory management
                    ForceGarbageCollection(chunkIndex);
                }
            }
            
            // Clear checkpoint on success
            _checkpointMgr.ClearCheckpoint();
            _timeoutMonitor.Dispose();
            
            Log($"✅ Multi-floor processing complete: {totalResult.SuccessfulFloors.Count} succeeded, {failedFloors.Count} failed");
            
            return totalResult;
        }
        
        /// <summary>
        /// Process a chunk of floors with error isolation
        /// </summary>
        private MultiFloorResult ProcessFloorChunkSafe(List<Level> chunk, OpeningFilter filter, SharedResourceCache cache)
        {
            var results = new List<FloorProcessingResult>();
            
            foreach (var level in chunk)
            {
                // Check timeouts before each floor
                if (_timeoutMonitor.CheckTotalTimeout())
                {
                    Log($"⏱ Total timeout before floor {level.Name}");
                    break;
                }
                
                // Process with safe wrapper
                var processor = new SafeSingleFloorProcessor(_doc, cache, _monitor, _timeoutMonitor);
                var result = processor.ProcessFloorSafe(level, filter);
                results.Add(result);
            }
            
            return new MultiFloorResult(results);
        }
        
        /// <summary>
        /// Save checkpoint with error tracking
        /// </summary>
        private void SaveCheckpointWithErrors(MultiFloorResult result, List<string> failedFloors)
        {
            try
            {
                var progress = new MultiFloorProgress
                {
                    CompletedFloors = result.SuccessfulFloors
                };
                
                _checkpointMgr.SaveCheckpoint(progress);
                Log($"💾 Checkpoint saved: {progress.CompletedFloors.Count} done, {failedFloors.Count} failed");
            }
            catch (Exception ex)
            {
                Log($"⚠️ Failed to save checkpoint: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Force garbage collection between chunks
        /// </summary>
        private void ForceGarbageCollection(int chunkIndex)
        {
            try
            {
                Log($"🧹 Forcing GC after chunk {chunkIndex}...");
                GC.Collect(2, GCCollectionMode.Optimized, blocking: false);
                GC.WaitForPendingFinalizers();
                
                var memAfter = GC.GetTotalMemory(false) / 1024 / 1024;
                Log($"💾 Memory after GC: {memAfter}MB");
            }
            catch (Exception ex)
            {
                Log($"⚠️ GC failed: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get default checkpoint path
        /// </summary>
        private string GetDefaultCheckpointPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = System.IO.Path.Combine(appData, "JSE_MEP_Openings", "Checkpoints");
            System.IO.Directory.CreateDirectory(dir);
            return System.IO.Path.Combine(dir, $"checkpoint_{_doc.Title}.json");
        }
        
        /// <summary>
        /// Log with timestamp
        /// </summary>
        private void Log(string message)
        {
            var timestamped = $"[{DateTime.Now:HH:mm:ss}] {message}";
            SafeFileLogger.SafeAppendText("multifloor.log", timestamped + "\n");
        }
    }
    
    /// <summary>
    /// Circuit breaker to stop processing after repeated failures
    /// </summary>
    public class CircuitBreaker
    {
        private readonly int _failureThreshold;
        private int _consecutiveFailures;
        private readonly object _lock = new object();
        
        public CircuitBreaker(int failureThreshold = 3)
        {
            _failureThreshold = failureThreshold;
        }
        
        public bool ShouldProcess()
        {
            lock (_lock)
            {
                return _consecutiveFailures < _failureThreshold;
            }
        }
        
        public void RecordSuccess()
        {
            lock (_lock)
            {
                _consecutiveFailures = 0;
            }
        }
        
        public void RecordFailure(int count = 1)
        {
            lock (_lock)
            {
                _consecutiveFailures += count;
            }
        }
        
        public int ConsecutiveFailures 
        { 
            get 
            { 
                lock (_lock) return _consecutiveFailures; 
            } 
        }
    }
}
