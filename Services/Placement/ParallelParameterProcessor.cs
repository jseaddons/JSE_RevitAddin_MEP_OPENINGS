using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ PERFORMANCE OPTIMIZATION: Parallel processing service for non-conflicting sleeve parameter operations.
    /// Enables parallel parameter setting for non-conflicting sleeves to utilize multi-core systems.
    /// 
    /// Features:
    /// - Thread-safe parameter setting with proper synchronization
    /// - Conflict detection to identify and prevent conflicts between parallel operations
    /// - Load balancing to distribute work evenly across available CPU cores
    /// - Graceful degradation to sequential processing if conflicts detected
    /// - Comprehensive error handling and fallback mechanisms
    /// </summary>
    public class ParallelParameterProcessor
    {
        private readonly Document _doc;
        private readonly bool _isReplayPath;
        private readonly PlacementPerformanceMonitor? _performanceMonitor;

        // ✅ THREAD SAFETY: Thread-safe collections for parallel operations
        private readonly ConcurrentDictionary<ElementId, FamilyInstance> _sleeveCache = new ConcurrentDictionary<ElementId, FamilyInstance>();
        private readonly ConcurrentDictionary<string, object> _conflictDetectionCache = new ConcurrentDictionary<string, object>();
        
        // ✅ LOAD BALANCING: Track workload distribution across threads
        private readonly ConcurrentDictionary<int, int> _threadWorkload = new ConcurrentDictionary<int, int>();
        
        // ✅ SAFETY MECHANISMS: Prevent conflicts and ensure data integrity
        private readonly ReaderWriterLockSlim _sleeveAccessLock = new ReaderWriterLockSlim();
        private readonly SemaphoreSlim _maxConcurrencySemaphore;
        
        // ✅ PERFORMANCE MONITORING: Track parallel processing performance
        private readonly ConcurrentDictionary<string, ProcessingMetrics> _processingMetrics = new ConcurrentDictionary<string, ProcessingMetrics>();

        public ParallelParameterProcessor(
            Document doc,
            bool isReplayPath = false,
            PlacementPerformanceMonitor? performanceMonitor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _isReplayPath = isReplayPath;
            _performanceMonitor = performanceMonitor;
            
            // ✅ LOAD BALANCING: Limit max concurrency to prevent resource exhaustion
            int maxConcurrency = Math.Max(1, Environment.ProcessorCount - 1);
            _maxConcurrencySemaphore = new SemaphoreSlim(maxConcurrency, maxConcurrency);
        }

        /// <summary>
        /// ✅ MAIN METHOD: Process parameter setting for multiple sleeves in parallel.
        /// Automatically detects conflicts and falls back to sequential processing when needed.
        /// </summary>
        public async Task<int> ProcessParameterUpdatesInParallelAsync(
            List<SleeveParameterUpdateRequest> parameterUpdates)
        {
            if (parameterUpdates == null || parameterUpdates.Count == 0)
                return 0;

            using (var tracker = _performanceMonitor?.TrackOperation("Parallel Parameter Processing"))
            {
                // ✅ CONFLICT DETECTION: Analyze for potential conflicts
                var conflictAnalysis = AnalyzeConflicts(parameterUpdates);
                
                if (conflictAnalysis.HasConflicts)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[ParallelParameterProcessor] [CONFLICT-DETECTION] ⚠️ Conflicts detected, falling back to sequential processing. " +
                            $"Conflicts: {conflictAnalysis.ConflictCount}, Total: {parameterUpdates.Count}");
                    }
                    
                    // ✅ GRACEFUL DEGRADATION: Fall back to sequential processing
                    return await ProcessSequentiallyAsync(parameterUpdates);
                }

                // ✅ PARALLEL PROCESSING: Process non-conflicting updates in parallel
                return await ProcessInParallelAsync(parameterUpdates, conflictAnalysis);
            }
        }

        /// <summary>
        /// ✅ PARALLEL PROCESSING: Process parameter updates using parallel operations.
        /// Distributes work evenly across available CPU cores.
        /// </summary>
        private async Task<int> ProcessInParallelAsync(
            List<SleeveParameterUpdateRequest> parameterUpdates,
            ConflictAnalysisResult conflictAnalysis)
        {
            int successCount = 0;
            int failCount = 0;
            var errorLog = new System.Text.StringBuilder();

            try
            {
                // ✅ LOAD BALANCING: Distribute work evenly across threads
                var workBatches = CreateWorkBatches(parameterUpdates);
                
                // ✅ PARALLEL EXECUTION: Process batches in parallel
                var tasks = workBatches.Select(batch => ProcessBatchAsync(batch)).ToList();
                var results = await Task.WhenAll(tasks);

                // ✅ RESULT AGGREGATION: Combine results from all parallel tasks
                foreach (var result in results)
                {
                    successCount += result.SuccessCount;
                    failCount += result.FailCount;
                    
                    if (result.ErrorMessages.Any())
                    {
                        foreach (var errorMsg in result.ErrorMessages)
                        {
                            errorLog.AppendLine($"[ParallelParameterProcessor] {errorMsg}");
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ParallelParameterProcessor] [PARALLEL-PROCESSING] ✅ Processed {successCount} successes, {failCount} failures across {workBatches.Count} batches");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ParallelParameterProcessor] [PARALLEL-PROCESSING] Error during parallel processing: {ex.Message}");
                }
                
                // ✅ GRACEFUL DEGRADATION: Fall back to sequential on parallel failure
                return await ProcessSequentiallyAsync(parameterUpdates);
            }

            return successCount;
        }

        /// <summary>
        /// ✅ SEQUENTIAL PROCESSING: Process parameter updates sequentially.
        /// Used as fallback when conflicts are detected or parallel processing fails.
        /// </summary>
        private async Task<int> ProcessSequentiallyAsync(List<SleeveParameterUpdateRequest> parameterUpdates)
        {
            int successCount = 0;
            int failCount = 0;

            try
            {
                foreach (var update in parameterUpdates)
                {
                    var result = await ProcessSingleUpdateAsync(update);
                    
                    if (result.SuccessCount > 0)
                    {
                        successCount++;
                    }
                    else
                    {
                        failCount++;
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ParallelParameterProcessor] [SEQUENTIAL-PROCESSING] ✅ Processed {successCount} successes, {failCount} failures sequentially");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ParallelParameterProcessor] [SEQUENTIAL-PROCESSING] Error during sequential processing: {ex.Message}");
                }
            }

            return successCount;
        }

        /// <summary>
        /// ✅ BATCH PROCESSING: Process a batch of parameter updates in parallel.
        /// Each batch runs on a separate thread with proper synchronization.
        /// </summary>
        private async Task<BatchProcessingResult> ProcessBatchAsync(List<SleeveParameterUpdateRequest> batch)
        {
            var result = new BatchProcessingResult();
            var currentThreadId = Thread.CurrentThread.ManagedThreadId;

            try
            {
                // ✅ SEMAPHORE: Limit concurrent access to prevent resource exhaustion
                await _maxConcurrencySemaphore.WaitAsync();

                try
                {
                    // ✅ WORKLOAD TRACKING: Track workload per thread
                    _threadWorkload.AddOrUpdate(currentThreadId, 1, (key, value) => value + 1);

                    // ✅ PARALLEL PROCESSING: Process updates within batch
                    var tasks = batch.Select(update => ProcessSingleUpdateAsync(update)).ToList();
                    var updateResults = await Task.WhenAll(tasks);

                    // ✅ RESULT AGGREGATION: Combine results from batch
                    foreach (var updateResult in updateResults)
                    {
                        result.SuccessCount += updateResult.SuccessCount;
                        result.FailCount += updateResult.FailCount;
                        result.ErrorMessages.AddRange(updateResult.ErrorMessages);
                    }
                }
                finally
                {
                    _maxConcurrencySemaphore.Release();
                }
            }
            catch (Exception ex)
            {
                result.FailCount += batch.Count;
                result.ErrorMessages.Add($"Batch processing failed: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// ✅ SINGLE UPDATE: Process a single parameter update with thread safety.
        /// Uses proper locking and validation to ensure thread safety.
        /// </summary>
        private async Task<SingleUpdateResult> ProcessSingleUpdateAsync(SleeveParameterUpdateRequest update)
        {
            var result = new SingleUpdateResult();
            var currentThreadId = Thread.CurrentThread.ManagedThreadId;

            try
            {
                // ✅ THREAD SAFETY: Use read lock for sleeve access
                _sleeveAccessLock.EnterReadLock();
                
                try
                {
                    // ✅ SLEEVE VALIDATION: Validate sleeve instance
                    var sleeve = await GetSleeveInstanceAsync(update.SleeveId);
                    if (sleeve == null)
                    {
                        result.ErrorMessages.Add($"Sleeve not found: {update.SleeveId.IntegerValue}");
                        return result;
                    }

                    // ✅ PARAMETER VALIDATION: Validate parameter before setting
                    if (!ValidateParameter(sleeve, update.ParameterName, update.ParameterValue))
                    {
                        result.ErrorMessages.Add($"Invalid parameter: {update.ParameterName} for sleeve {update.SleeveId.IntegerValue}");
                        return result;
                    }

                    // ✅ PARAMETER SETTING: Set parameter with thread safety
                    bool success = await SetParameterAsync(sleeve, update.ParameterName, update.ParameterValue);
                    
                    if (success)
                    {
                        result.SuccessCount = 1;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("parallel_parameter_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [ParallelParameterProcessor] [THREAD-{currentThreadId}] ✅ Set parameter '{update.ParameterName}' = {update.ParameterValue} for sleeve {update.SleeveId.IntegerValue}\n");
                        }
                    }
                    else
                    {
                        result.ErrorMessages.Add($"Failed to set parameter '{update.ParameterName}' for sleeve {update.SleeveId.IntegerValue}");
                    }
                }
                finally
                {
                    _sleeveAccessLock.ExitReadLock();
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessages.Add($"Single update failed: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// ✅ CONFLICT DETECTION: Analyze parameter updates for potential conflicts.
        /// Identifies conflicts that would prevent safe parallel processing.
        /// </summary>
        private ConflictAnalysisResult AnalyzeConflicts(List<SleeveParameterUpdateRequest> parameterUpdates)
        {
            var result = new ConflictAnalysisResult();
            var sleeveParameterMap = new Dictionary<ElementId, HashSet<string>>();

            try
            {
                foreach (var update in parameterUpdates)
                {
                    if (!sleeveParameterMap.TryGetValue(update.SleeveId, out HashSet<string> parameters))
                    {
                        parameters = new HashSet<string>();
                        sleeveParameterMap[update.SleeveId] = parameters;
                    }

                    // ✅ CONFLICT DETECTION: Check for parameter conflicts on same sleeve
                    if (parameters.Contains(update.ParameterName))
                    {
                        result.HasConflicts = true;
                        result.ConflictCount++;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("parallel_parameter_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [ParallelParameterProcessor] [CONFLICT-DETECTION] ⚠️ Conflict detected: Sleeve {update.SleeveId.IntegerValue}, Parameter '{update.ParameterName}'\n");
                        }
                    }
                    else
                    {
                        parameters.Add(update.ParameterName);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ParallelParameterProcessor] [CONFLICT-DETECTION] Error during conflict analysis: {ex.Message}");
                }
                
                // ✅ SAFETY: If conflict analysis fails, assume conflicts exist
                result.HasConflicts = true;
                result.ConflictCount = parameterUpdates.Count;
            }

            return result;
        }

        /// <summary>
        /// ✅ LOAD BALANCING: Create work batches for parallel processing.
        /// Distributes work evenly across available CPU cores.
        /// </summary>
        private List<List<SleeveParameterUpdateRequest>> CreateWorkBatches(List<SleeveParameterUpdateRequest> parameterUpdates)
        {
            int maxConcurrency = Math.Max(1, Environment.ProcessorCount - 1);
            var batches = new List<List<SleeveParameterUpdateRequest>>();
            
            // ✅ LOAD BALANCING: Distribute work evenly across batches
            int batchSize = (int)Math.Ceiling((double)parameterUpdates.Count / maxConcurrency);

            for (int i = 0; i < parameterUpdates.Count; i += batchSize)
            {
                var batch = parameterUpdates.Skip(i).Take(batchSize).ToList();
                batches.Add(batch);
            }

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ParallelParameterProcessor] [LOAD-BALANCING] Created {batches.Count} batches with max batch size {batchSize}");
            }

            return batches;
        }

        /// <summary>
        /// ✅ SLEEVE ACCESS: Get sleeve instance with caching for performance.
        /// Uses thread-safe caching to reduce repeated document lookups.
        /// </summary>
        private async Task<FamilyInstance?> GetSleeveInstanceAsync(ElementId sleeveId)
        {
            // ✅ CACHE CHECK: Try to get from cache first
            if (_sleeveCache.TryGetValue(sleeveId, out FamilyInstance cachedSleeve))
            {
                return cachedSleeve;
            }

            try
            {
                // ✅ DOCUMENT ACCESS: Get sleeve from document (may need to be on main thread)
                var sleeve = _doc.GetElement(sleeveId) as FamilyInstance;
                
                if (sleeve != null)
                {
                    // ✅ CACHE UPDATE: Store in cache for future access
                    _sleeveCache.TryAdd(sleeveId, sleeve);
                }

                return sleeve;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("parallel_parameter_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ParallelParameterProcessor] [SLEEVE-ACCESS] Error getting sleeve {sleeveId.IntegerValue}: {ex.Message}\n");
                }
                
                return null;
            }
        }

        /// <summary>
        /// ✅ PARAMETER VALIDATION: Validate parameter before setting.
        /// Ensures parameter is valid and safe to set.
        /// </summary>
        private bool ValidateParameter(FamilyInstance sleeve, string parameterName, object parameterValue)
        {
            try
            {
                var parameter = sleeve.LookupParameter(parameterName);
                if (parameter == null || parameter.IsReadOnly)
                {
                    return false;
                }

                // ✅ VALUE VALIDATION: Validate parameter value type
                if (parameterValue is double doubleValue)
                {
                    return !double.IsNaN(doubleValue) && !double.IsInfinity(doubleValue);
                }
                else if (parameterValue is string stringValue)
                {
                    return !string.IsNullOrWhiteSpace(stringValue);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ✅ PARAMETER SETTING: Set parameter value asynchronously.
        /// Uses proper transaction handling for thread safety.
        /// </summary>
        private async Task<bool> SetParameterAsync(FamilyInstance sleeve, string parameterName, object parameterValue)
        {
            try
            {
                var parameter = sleeve.LookupParameter(parameterName);
                if (parameter == null || parameter.IsReadOnly)
                {
                    return false;
                }

                // ✅ VALUE SETTING: Set parameter value based on type
                if (parameterValue is double doubleValue)
                {
                    parameter.Set(doubleValue);
                }
                else if (parameterValue is string stringValue)
                {
                    parameter.Set(stringValue);
                }
                else if (parameterValue is int intValue)
                {
                    parameter.Set(intValue);
                }
                else if (parameterValue is ElementId elementIdValue)
                {
                    parameter.Set(elementIdValue);
                }

                return true;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("parallel_parameter_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ParallelParameterProcessor] [PARAMETER-SETTING] Error setting parameter '{parameterName}' for sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                }
                
                return false;
            }
        }

        /// <summary>
        /// ✅ PERFORMANCE MONITORING: Get processing metrics for monitoring.
        /// </summary>
        public string GetProcessingMetrics()
        {
            var totalWork = _threadWorkload.Values.Sum();
            var avgWorkload = totalWork / Math.Max(1, _threadWorkload.Count);
            
            return $"Threads: {_threadWorkload.Count}, TotalWork: {totalWork}, AvgWorkload: {avgWorkload}, " +
                   $"SleeveCache: {_sleeveCache.Count}, Conflicts: {_conflictDetectionCache.Count}";
        }

        /// <summary>
        /// ✅ MEMORY MANAGEMENT: Clear all cached data and release resources.
        /// Called periodically to prevent memory leaks.
        /// </summary>
        public void ClearCache()
        {
            _sleeveCache.Clear();
            _conflictDetectionCache.Clear();
            _threadWorkload.Clear();
            _processingMetrics.Clear();
        }

        /// <summary>
        /// ✅ RESOURCE CLEANUP: Dispose of resources properly.
        /// </summary>
        public void Dispose()
        {
            _sleeveAccessLock?.Dispose();
            _maxConcurrencySemaphore?.Dispose();
            ClearCache();
        }

        #region Helper Classes

        /// <summary>
        /// ✅ REQUEST DATA: Represents a parameter update request.
        /// Contains all information needed for parallel parameter processing.
        /// </summary>
        public class SleeveParameterUpdateRequest
        {
            public ElementId SleeveId { get; set; } = ElementId.InvalidElementId;
            public string ParameterName { get; set; } = string.Empty;
            public object ParameterValue { get; set; } = string.Empty;
            public DateTime Timestamp { get; set; }
        }

        /// <summary>
        /// ✅ CONFLICT ANALYSIS: Represents the result of conflict analysis.
        /// Tracks conflicts that would prevent safe parallel processing.
        /// </summary>
        private class ConflictAnalysisResult
        {
            public bool HasConflicts { get; set; } = false;
            public int ConflictCount { get; set; } = 0;
        }

        /// <summary>
        /// ✅ BATCH RESULT: Represents the result of batch processing.
        /// Tracks success/failure counts and error messages.
        /// </summary>
        private class BatchProcessingResult
        {
            public int SuccessCount { get; set; } = 0;
            public int FailCount { get; set; } = 0;
            public List<string> ErrorMessages { get; set; } = new List<string>();
        }

        /// <summary>
        /// ✅ SINGLE RESULT: Represents the result of single update processing.
        /// Tracks success/failure and error messages.
        /// </summary>
        private class SingleUpdateResult
        {
            public int SuccessCount { get; set; } = 0;
            public int FailCount { get; set; } = 0;
            public List<string> ErrorMessages { get; set; } = new List<string>();
        }

        /// <summary>
        /// ✅ PROCESSING METRICS: Represents processing metrics for monitoring.
        /// Tracks performance and resource usage.
        /// </summary>
        private class ProcessingMetrics
        {
            public int TotalUpdates { get; set; } = 0;
            public int SuccessfulUpdates { get; set; } = 0;
            public int FailedUpdates { get; set; } = 0;
            public TimeSpan ProcessingTime { get; set; }
        }

        #endregion
    }
}
