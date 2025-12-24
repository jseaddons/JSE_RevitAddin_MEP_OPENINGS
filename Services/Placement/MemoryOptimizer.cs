using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ PERFORMANCE OPTIMIZATION: Memory optimization service for parameter operations.
    /// Reduces memory allocations during parameter operations to minimize garbage collection.
    /// 
    /// Features:
    /// - Object pooling for frequently created objects
    /// - String optimization to minimize string allocations
    /// - Memory leak detection and prevention
    /// - GC optimization to reduce garbage collection pressure
    /// - Comprehensive memory usage monitoring
    /// </summary>
    public class MemoryOptimizer
    {
        private readonly bool _isReplayPath;
        private readonly PlacementPerformanceMonitor? _performanceMonitor;

        // ✅ OBJECT POOLING: Pool frequently allocated objects
        private readonly ConcurrentQueue<ParameterUpdateRequest> _parameterRequestPool = new ConcurrentQueue<ParameterUpdateRequest>();
        private readonly ConcurrentQueue<string> _stringPool = new ConcurrentQueue<string>();
        private readonly ConcurrentDictionary<string, string> _stringInternCache = new ConcurrentDictionary<string, string>();
        
        // ✅ MEMORY LEAK DETECTION: Track object lifecycles
        private readonly ConcurrentDictionary<long, WeakReference> _trackedObjects = new ConcurrentDictionary<long, WeakReference>();
        private readonly ConcurrentDictionary<string, int> _allocationCounts = new ConcurrentDictionary<string, int>();
        
        // ✅ MEMORY MONITORING: Track memory usage patterns
        private readonly ConcurrentDictionary<string, MemoryMetrics> _memoryMetrics = new ConcurrentDictionary<string, MemoryMetrics>();
        private readonly System.Threading.Timer _memoryMonitoringTimer;
        
        // ✅ GC OPTIMIZATION: Control garbage collection behavior
        private readonly GCMemoryPressureTracker _gcPressureTracker;
        private readonly bool _targetPressureLevel = true; // Simplified for compatibility

        public MemoryOptimizer(
            bool isReplayPath = false,
            PlacementPerformanceMonitor? performanceMonitor = null)
        {
            _isReplayPath = isReplayPath;
            _performanceMonitor = performanceMonitor;
            
        // ✅ MEMORY MONITORING: Start periodic memory monitoring
        _memoryMonitoringTimer = new System.Threading.Timer(MonitorMemoryUsage, null, TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30));
            
            // ✅ GC OPTIMIZATION: Initialize GC pressure tracking
            _gcPressureTracker = new GCMemoryPressureTracker();
        }

        /// <summary>
        /// ✅ MAIN METHOD: Optimize memory usage for parameter operations.
        /// Applies all memory optimization techniques to reduce allocations.
        /// </summary>
        public T OptimizeMemoryUsage<T>(Func<T> operation, string operationName = "MemoryOperation")
        {
            if (operation == null)
                throw new ArgumentNullException(nameof(operation));

            var metrics = GetOrCreateMetrics(operationName);
            var initialMemory = GC.GetTotalMemory(false);
            var initialGcCounts = GetGCCounts();

            try
            {
                // ✅ MEMORY TRACKING: Start tracking memory usage
                metrics.StartOperation();
                
                // ✅ GC OPTIMIZATION: Set memory pressure level
                _gcPressureTracker.SetMemoryPressure(_targetPressureLevel);
                
                // ✅ EXECUTE OPERATION: Run the memory-intensive operation
                var result = operation();
                
                // ✅ MEMORY TRACKING: Record successful completion
                var finalMemory = GC.GetTotalMemory(false);
                var memoryDelta = finalMemory - initialMemory;
                var gcCountsAfter = GetGCCounts();
                var gcCollections = GetGCDelta(initialGcCounts, gcCountsAfter);
                
                metrics.RecordSuccess(memoryDelta, gcCollections);
                
                // ✅ MEMORY LEAK DETECTION: Check for potential leaks
                DetectMemoryLeaks(operationName);
                
                return result;
            }
            catch (Exception ex)
            {
                // ✅ ERROR HANDLING: Record failed operation
                var finalMemory = GC.GetTotalMemory(false);
                var memoryDelta = finalMemory - initialMemory;
                var gcCountsAfter = GetGCCounts();
                var gcCollections = GetGCDelta(initialGcCounts, gcCountsAfter);
                
                metrics.RecordFailure(memoryDelta, gcCollections, ex.Message);
                
                throw;
            }
            finally
            {
                // ✅ GC OPTIMIZATION: Reset memory pressure
                _gcPressureTracker.ResetMemoryPressure();
                
                // ✅ MEMORY CLEANUP: Force cleanup if needed
                if (ShouldForceCleanup())
                {
                    ForceMemoryCleanup();
                }
            }
        }

        /// <summary>
        /// ✅ OBJECT POOLING: Get a pooled parameter update request object.
        /// Reuses objects to reduce memory allocations.
        /// </summary>
        public ParameterUpdateRequest GetPooledParameterRequest()
        {
            if (_parameterRequestPool.TryDequeue(out ParameterUpdateRequest request))
            {
                // ✅ OBJECT REUSE: Reset object state for reuse
                request.Reset();
                return request;
            }

            // ✅ OBJECT CREATION: Create new object if pool is empty
            return new ParameterUpdateRequest();
        }

        /// <summary>
        /// ✅ OBJECT POOLING: Return a parameter update request object to the pool.
        /// Clears object state and returns it to the pool for reuse.
        /// </summary>
        public void ReturnToPool(ParameterUpdateRequest request)
        {
            if (request == null)
                return;

            // ✅ OBJECT CLEANUP: Clear object state before pooling
            request.Clear();
            
            // ✅ POOL LIMIT: Prevent pool from growing too large
            if (_parameterRequestPool.Count < 1000)
            {
                _parameterRequestPool.Enqueue(request);
            }
        }

        /// <summary>
        /// ✅ STRING OPTIMIZATION: Get an interned string to reduce allocations.
        /// Uses string interning to reduce duplicate string allocations.
        /// </summary>
        public string GetInternedString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // ✅ STRING INTERNING: Use interned strings to reduce allocations
            return _stringInternCache.GetOrAdd(value, s => string.Intern(s));
        }

        /// <summary>
        /// ✅ STRING OPTIMIZATION: Format strings efficiently to minimize allocations.
        /// Uses string builders and pooling to reduce string allocations.
        /// </summary>
        public string FormatString(string format, params object[] args)
        {
            // ✅ STRING BUILDER: Use StringBuilder to reduce allocations
            var builder = new System.Text.StringBuilder(format.Length + (args.Length * 10));
            builder.AppendFormat(format, args);
            
            var result = builder.ToString();
            
            // ✅ STRING INTERNING: Intern common strings
            if (result.Length < 100) // Only intern short strings
            {
                return GetInternedString(result);
            }
            
            return result;
        }

        /// <summary>
        /// ✅ MEMORY LEAK DETECTION: Track object for potential memory leaks.
        /// Uses weak references to detect objects that aren't being properly disposed.
        /// </summary>
        public void TrackObject(object obj, string objectName)
        {
            if (obj == null)
                return;

            var objectId = obj.GetHashCode();
            var weakRef = new WeakReference(obj, trackResurrection: true);
            
            _trackedObjects.TryAdd(objectId, weakRef);
            _allocationCounts.AddOrUpdate(objectName, 1, (key, value) => value + 1);
        }

        /// <summary>
        /// ✅ MEMORY LEAK DETECTION: Check for potential memory leaks.
        /// Identifies objects that may not be getting properly disposed.
        /// </summary>
        public void DetectMemoryLeaks(string operationName)
        {
            var leakedObjects = new List<string>();
            
            foreach (var kvp in _trackedObjects)
            {
                if (kvp.Value.Target == null)
                {
                    // Object has been garbage collected - remove from tracking
                    _trackedObjects.TryRemove(kvp.Key, out _);
                }
                else if (kvp.Value.IsAlive && !kvp.Value.Target.GetType().IsValueType)
                {
                    // Potential memory leak detected
                    var objectType = kvp.Value.Target.GetType().Name;
                    leakedObjects.Add($"{objectType} (ID: {kvp.Key})");
                }
            }

            if (leakedObjects.Any() && !DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("memory_leak_detection.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [MemoryOptimizer] [MEMORY-LEAK-DETECTION] ⚠️ Potential memory leaks detected in {operationName}:\n" +
                    string.Join(", ", leakedObjects) + "\n");
            }
        }

        /// <summary>
        /// ✅ MEMORY CLEANUP: Force garbage collection and cleanup.
        /// Called when memory pressure is high to free up memory.
        /// </summary>
        public void ForceMemoryCleanup()
        {
            try
            {
                // ✅ GC OPTIMIZATION: Force garbage collection
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true);
                GC.WaitForPendingFinalizers();
                
                // ✅ MEMORY CLEANUP: Clear object pools periodically
                if (_parameterRequestPool.Count > 500)
                {
                    ClearObjectPools();
                }
                
                // ✅ MEMORY CLEANUP: Clear string intern cache periodically
                if (_stringInternCache.Count > 10000)
                {
                    ClearStringInternCache();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[MemoryOptimizer] [MEMORY-CLEANUP] Forced garbage collection completed");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[MemoryOptimizer] [MEMORY-CLEANUP] Error during forced cleanup: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ MEMORY MONITORING: Monitor memory usage and detect issues.
        /// Runs periodically to track memory patterns and detect problems.
        /// </summary>
        private void MonitorMemoryUsage(object state)
        {
            try
            {
                var totalMemory = GC.GetTotalMemory(false);
                var generationCounts = new int[GC.MaxGeneration + 1];
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    generationCounts[i] = GC.CollectionCount(i);
                }
                
                var memoryInfo = new ProcessMemoryInfo
                {
                    TotalMemory = totalMemory,
                    GenerationCounts = generationCounts,
                    Timestamp = DateTime.Now
                };

                // ✅ MEMORY ANALYSIS: Analyze memory patterns
                AnalyzeMemoryPatterns(memoryInfo);
                
                // ✅ MEMORY ALERTS: Check for memory issues
                CheckMemoryAlerts(totalMemory);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[MemoryOptimizer] [MEMORY-MONITORING] Error during memory monitoring: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ MEMORY ANALYSIS: Analyze memory usage patterns.
        /// Identifies trends and potential memory issues.
        /// </summary>
        private void AnalyzeMemoryPatterns(ProcessMemoryInfo memoryInfo)
        {
            // ✅ MEMORY TREND: Track memory usage over time
            var currentMetrics = _memoryMetrics.Values
                .OrderByDescending(m => m.LastUpdated)
                .FirstOrDefault();

            if (currentMetrics != null)
            {
                var memoryDelta = memoryInfo.TotalMemory - currentMetrics.TotalMemory;
                var timeDelta = (DateTime.Now - currentMetrics.LastUpdated).TotalSeconds;

                if (memoryDelta > 100 * 1024 * 1024 && timeDelta < 60) // 100MB increase in 60 seconds
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("memory_analysis.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [MemoryOptimizer] [MEMORY-ANALYSIS] ⚠️ Rapid memory growth detected: {memoryDelta / 1024 / 1024:F1}MB in {timeDelta:F1}s\n");
                    }
                }
            }
        }

        /// <summary>
        /// ✅ MEMORY ALERTS: Check for memory issues and alert if needed.
        /// Monitors memory pressure and triggers cleanup when needed.
        /// </summary>
        private void CheckMemoryAlerts(long totalMemory)
        {
            var memoryThreshold = 1024L * 1024 * 1024; // 1GB threshold
            
            if (totalMemory > memoryThreshold)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("memory_alerts.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [MemoryOptimizer] [MEMORY-ALERTS] ⚠️ High memory usage detected: {totalMemory / 1024 / 1024:F1}MB\n");
                }
                
                // ✅ MEMORY CLEANUP: Trigger cleanup for high memory usage
                ForceMemoryCleanup();
            }
        }

        /// <summary>
        /// ✅ MEMORY UTILIZATION: Get current memory utilization statistics.
        /// Provides detailed memory usage information for monitoring.
        /// </summary>
        public MemoryUtilizationStats GetMemoryUtilizationStats()
        {
            var totalMemory = GC.GetTotalMemory(false);
            var generationCounts = new int[GC.MaxGeneration + 1];
            for (int i = 0; i <= GC.MaxGeneration; i++)
            {
                generationCounts[i] = GC.CollectionCount(i);
            }

            return new MemoryUtilizationStats
            {
                TotalMemoryBytes = totalMemory,
                TotalMemoryMB = totalMemory / (1024.0 * 1024.0),
                GenerationCounts = generationCounts,
                PoolSize = _parameterRequestPool.Count,
                InternCacheSize = _stringInternCache.Count,
                TrackedObjectsCount = _trackedObjects.Count,
                AllocationCounts = _allocationCounts.ToDictionary(kvp => kvp.Key, kvp => kvp.Value),
                Timestamp = DateTime.Now
            };
        }

        /// <summary>
        /// ✅ MEMORY CLEANUP: Clear object pools to free memory.
        /// Called periodically to prevent pools from growing too large.
        /// </summary>
        private void ClearObjectPools()
        {
            while (_parameterRequestPool.TryDequeue(out _))
            {
                // Clear pool contents
            }
        }

        /// <summary>
        /// ✅ MEMORY CLEANUP: Clear string intern cache to free memory.
        /// Called periodically to prevent cache from growing too large.
        /// </summary>
        private void ClearStringInternCache()
        {
            _stringInternCache.Clear();
        }

        /// <summary>
        /// ✅ MEMORY CLEANUP: Check if forced cleanup is needed.
        /// Determines if memory pressure requires immediate cleanup.
        /// </summary>
        private bool ShouldForceCleanup()
        {
            var totalMemory = GC.GetTotalMemory(false);
            var memoryThreshold = 512L * 1024 * 1024; // 512MB threshold
            
            return totalMemory > memoryThreshold;
        }

        /// <summary>
        /// ✅ GC COUNTS: Get current garbage collection counts.
        /// Used to track GC activity during operations.
        /// </summary>
        private int[] GetGCCounts()
        {
            var counts = new int[GC.MaxGeneration + 1];
            for (int i = 0; i <= GC.MaxGeneration; i++)
            {
                counts[i] = GC.CollectionCount(i);
            }
            return counts;
        }

        /// <summary>
        /// ✅ GC DELTA: Calculate GC collection delta between two measurements.
        /// Used to track GC activity during operations.
        /// </summary>
        private int[] GetGCDelta(int[] initialCounts, int[] finalCounts)
        {
            var delta = new int[initialCounts.Length];
            for (int i = 0; i < initialCounts.Length; i++)
            {
                delta[i] = finalCounts[i] - initialCounts[i];
            }
            return delta;
        }

        /// <summary>
        /// ✅ METRICS: Get or create memory metrics for an operation.
        /// Tracks memory usage patterns for specific operations.
        /// </summary>
        private MemoryMetrics GetOrCreateMetrics(string operationName)
        {
            return _memoryMetrics.GetOrAdd(operationName, name => new MemoryMetrics(name));
        }

        /// <summary>
        /// ✅ RESOURCE CLEANUP: Dispose of resources properly.
        /// </summary>
        public void Dispose()
        {
            _memoryMonitoringTimer?.Dispose();
            _gcPressureTracker?.Dispose();
            ClearObjectPools();
            ClearStringInternCache();
            _trackedObjects.Clear();
            _allocationCounts.Clear();
            _memoryMetrics.Clear();
        }

        #region Helper Classes

        /// <summary>
        /// ✅ REQUEST DATA: Represents a parameter update request with pooling support.
        /// Designed for reuse to minimize memory allocations.
        /// </summary>
        public class ParameterUpdateRequest
        {
            public ElementId SleeveId { get; set; } = ElementId.InvalidElementId;
            public string ParameterName { get; set; } = string.Empty;
            public object ParameterValue { get; set; } = string.Empty;
            public DateTime Timestamp { get; set; }

            /// <summary>
            /// ✅ OBJECT POOLING: Reset object state for reuse.
            /// </summary>
            public void Reset()
            {
                SleeveId = ElementId.InvalidElementId;
                ParameterName = string.Empty;
                ParameterValue = string.Empty;
                Timestamp = DateTime.MinValue;
            }

            /// <summary>
            /// ✅ OBJECT POOLING: Clear object state completely.
            /// </summary>
            public void Clear()
            {
                Reset();
            }
        }

        /// <summary>
        /// ✅ MEMORY METRICS: Represents memory usage metrics for an operation.
        /// Tracks memory allocation patterns and performance.
        /// </summary>
        private class MemoryMetrics
        {
            public string OperationName { get; }
            public long TotalMemory { get; set; }
            public int SuccessCount { get; set; }
            public int FailureCount { get; set; }
            public DateTime LastUpdated { get; set; }
            public List<MemoryOperation> Operations { get; } = new List<MemoryOperation>();

            public MemoryMetrics(string operationName)
            {
                OperationName = operationName;
                LastUpdated = DateTime.Now;
            }

            public void StartOperation()
            {
                TotalMemory = GC.GetTotalMemory(false);
                LastUpdated = DateTime.Now;
            }

            public void RecordSuccess(long memoryDelta, int[] gcCollections)
            {
                SuccessCount++;
                Operations.Add(new MemoryOperation
                {
                    Success = true,
                    MemoryDelta = memoryDelta,
                    GCCollections = gcCollections,
                    Timestamp = DateTime.Now
                });
                LastUpdated = DateTime.Now;
            }

            public void RecordFailure(long memoryDelta, int[] gcCollections, string errorMessage)
            {
                FailureCount++;
                Operations.Add(new MemoryOperation
                {
                    Success = false,
                    MemoryDelta = memoryDelta,
                    GCCollections = gcCollections,
                    ErrorMessage = errorMessage,
                    Timestamp = DateTime.Now
                });
                LastUpdated = DateTime.Now;
            }
        }

        /// <summary>
        /// ✅ MEMORY OPERATION: Represents a single memory operation.
        /// Tracks memory usage for individual operations.
        /// </summary>
        private class MemoryOperation
        {
            public bool Success { get; set; }
            public long MemoryDelta { get; set; }
            public int[] GCCollections { get; set; }
            public string ErrorMessage { get; set; } = string.Empty;
            public DateTime Timestamp { get; set; }
        }

        /// <summary>
        /// ✅ MEMORY UTILIZATION: Represents current memory utilization statistics.
        /// Provides detailed memory usage information.
        /// </summary>
        public class MemoryUtilizationStats
        {
            public long TotalMemoryBytes { get; set; }
            public double TotalMemoryMB { get; set; }
            public int[] GenerationCounts { get; set; } = new int[0];
            public int PoolSize { get; set; }
            public int InternCacheSize { get; set; }
            public int TrackedObjectsCount { get; set; }
            public Dictionary<string, int> AllocationCounts { get; set; } = new Dictionary<string, int>();
            public DateTime Timestamp { get; set; }
        }

        /// <summary>
        /// ✅ PROCESS MEMORY: Represents process memory information.
        /// Used for memory monitoring and analysis.
        /// </summary>
        private class ProcessMemoryInfo
        {
            public long TotalMemory { get; set; }
            public int[] GenerationCounts { get; set; } = new int[0];
            public DateTime Timestamp { get; set; }
        }

        /// <summary>
        /// ✅ GC PRESSURE: Tracks and manages garbage collection pressure.
        /// Helps optimize GC behavior during memory-intensive operations.
        /// </summary>
        private class GCMemoryPressureTracker : IDisposable
        {
            private bool _disposed = false;
            private bool _noGcRegionStarted = false;

            public void SetMemoryPressure(bool enablePressure)
            {
                if (_disposed || _noGcRegionStarted) return;

                try
                {
                    // ✅ GC PRESSURE: Set memory pressure level
                    if (enablePressure)
                    {
                        GC.TryStartNoGCRegion(100 * 1024 * 1024, // 100MB
                            disallowFullBlockingGC: true);
                        _noGcRegionStarted = true;
                    }
                }
                catch
                {
                    // Ignore GC pressure setting failures
                }
            }

            public void ResetMemoryPressure()
            {
                if (_disposed || !_noGcRegionStarted) return;

                try
                {
                    // ✅ GC PRESSURE: End no-GC region
                    GC.EndNoGCRegion();
                    _noGcRegionStarted = false;
                }
                catch
                {
                    // Ignore GC pressure reset failures
                }
            }

            public void Dispose()
            {
                if (!_disposed)
                {
                    try
                    {
                        if (_noGcRegionStarted)
                        {
                            GC.EndNoGCRegion();
                        }
                    }
                    catch
                    {
                        // Ignore cleanup failures
                    }
                    _disposed = true;
                }
            }
        }

        #endregion
    }
}
