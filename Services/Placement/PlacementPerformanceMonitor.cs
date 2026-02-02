using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Tracks performance metrics for placement operations (individual and cluster sleeves)
    /// Similar to Refresh PerformanceMonitor but tailored for placement operations
    /// </summary>
    public class PlacementPerformanceMonitor : JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor
    {
        private readonly string _logFileName;
        private readonly Stopwatch _totalTimer;
        private readonly Dictionary<string, OperationMetrics> _operations;
        private long _startMemoryBytes;
        
        public PlacementPerformanceMonitor(string logFileName)
        {
            _logFileName = logFileName;
            _totalTimer = Stopwatch.StartNew();
            _operations = new Dictionary<string, OperationMetrics>();
            _startMemoryBytes = GC.GetTotalMemory(false);
            _activeTrackers = new Dictionary<string, JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker>();
            
            // ✅ INITIALIZE LOG: Log start time (ALWAYS log, even in deployment mode)
            SafeFileLogger.SafeAppendTextAlways($"performance_{_logFileName}", 
                $"=== PLACEMENT PERFORMANCE MONITOR STARTED ===\n" +
                $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                $"Log File: {_logFileName}\n" +
                $"Start Memory: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB\n\n");
        }

        private readonly Dictionary<string, JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker> _activeTrackers;
        
        public bool IsEnabled => true;

        public void StartOperation(string operationName)
        {
            if (!_activeTrackers.ContainsKey(operationName))
            {
                _activeTrackers[operationName] = TrackOperation(operationName);
            }
        }

        public void StopOperation(string operationName, int itemCount = 0)
        {
            if (_activeTrackers.ContainsKey(operationName))
            {
                var tracker = _activeTrackers[operationName];
                tracker.SetItemCount(itemCount);
                tracker.Dispose();
                _activeTrackers.Remove(operationName);
            }
        }

        public void LogMetric(string metricName, object value)
        {
             SafeFileLogger.SafeAppendTextAlways($"performance_{_logFileName}", 
                $"[{DateTime.Now:HH:mm:ss.fff}] METRIC: {metricName} = {value}\n");
        }
        
        /// <summary>
        /// Start tracking an operation
        /// </summary>
        /// <summary>
        /// Start tracking an operation
        /// </summary>
        public JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker TrackOperation(string operationName)
        {
            return new OperationTracker(this, operationName);
        }
        
        internal void RecordOperation(string name, long milliseconds, long memoryBytes, int itemCount)
        {
            if (!_operations.ContainsKey(name))
            {
                _operations[name] = new OperationMetrics { Name = name };
            }
            
            var metrics = _operations[name];
            metrics.CallCount++;
            metrics.TotalMilliseconds += milliseconds;
            metrics.TotalMemoryBytes += memoryBytes;
            metrics.TotalItemCount += itemCount;
            
            if (milliseconds > metrics.MaxMilliseconds)
                metrics.MaxMilliseconds = milliseconds;
            
            if (milliseconds < metrics.MinMilliseconds || metrics.MinMilliseconds == 0)
                metrics.MinMilliseconds = milliseconds;
        }
        
        /// <summary>
        /// Generate final performance report with separate sections for individual and cluster placement
        /// </summary>
        public void GenerateReport(int totalIndividualSleeves, int totalClusters)
        {
            _totalTimer.Stop();
            
            var report = new StringBuilder();
            
            // ✅ BUILD TIMESTAMP: Get build timestamp to verify latest code is running
            string buildTimestamp = "unknown";
            string assemblyPath = "unknown";
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                assemblyPath = assembly?.Location ?? "unknown";
                if (!string.IsNullOrWhiteSpace(assemblyPath) && System.IO.File.Exists(assemblyPath))
                {
                    buildTimestamp = System.IO.File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            catch { }
            
            // ✅ FIX: Extract actual counts from tracked operations if parameters are 0
            if (totalIndividualSleeves == 0)
            {
                var individualOps = _operations.Values
                    .Where(op => IsIndividualPlacementOperation(op.Name) && op.Name.Contains("Bulk Individual Sleeve Placement"))
                    .ToList();
                totalIndividualSleeves = individualOps.Sum(op => op.TotalItemCount);
            }
            
            if (totalClusters == 0)
            {
                var clusterOps = _operations.Values
                    .Where(op => IsClusterPlacementOperation(op.Name) && op.Name.Contains("Bulk Cluster Sleeve Placement"))
                    .ToList();
                totalClusters = clusterOps.Sum(op => op.TotalItemCount);
            }
            
            report.AppendLine($"=== PLACEMENT PERFORMANCE REPORT ===");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"🔨 Build Timestamp: {buildTimestamp} | Assembly: {System.IO.Path.GetFileName(assemblyPath)}");
            report.AppendLine($"Total Workflow Time: {_totalTimer.ElapsedMilliseconds}ms ({_totalTimer.Elapsed:mm\\:ss})");
            // Placement phase = duration of "Bulk Individual Sleeve Placement" (matches log timestamps: create + params + commit)
            var bulkOp = _operations.Values.FirstOrDefault(op => op.Name != null && op.Name.Contains("Bulk Individual Sleeve Placement"));
            if (bulkOp != null && bulkOp.TotalMilliseconds > 0)
                report.AppendLine($"Placement Phase (Bulk block): {bulkOp.TotalMilliseconds}ms ({TimeSpan.FromMilliseconds(bulkOp.TotalMilliseconds):mm\\:ss})");
            report.AppendLine($"Total Individual Sleeves: {totalIndividualSleeves}");
            report.AppendLine($"Total Clusters: {totalClusters}");
            report.AppendLine();
            
            // Memory summary
            long currentMemory = GC.GetTotalMemory(false);
            long memoryDelta = currentMemory - _startMemoryBytes;
            report.AppendLine($"=== MEMORY USAGE ===");
            report.AppendLine($"Start: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"End: {currentMemory / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"Delta: {memoryDelta / 1024.0 / 1024.0:F2} MB");
            
            int totalItems = totalIndividualSleeves + totalClusters;
            if (totalItems > 0)
            {
                double memoryPerItem = (double)memoryDelta / totalItems / 1024.0; // KB
                report.AppendLine($"Memory per Item: {memoryPerItem:F2} KB");
            }
            report.AppendLine();
            
            // ✅ SEPARATE SECTIONS: Individual vs Cluster
            GenerateIndividualPlacementSection(report, totalIndividualSleeves);
            GenerateClusterPlacementSection(report, totalClusters);
            
            report.AppendLine();
            report.AppendLine($"=== END OF REPORT ===");
            
            // ✅ CRITICAL: Write report using SafeAppendTextAlways (ALWAYS log, even in deployment mode)
            SafeFileLogger.SafeAppendTextAlways($"performance_{_logFileName}", report.ToString());
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string reportPath = SafeFileLogger.GetLogFilePath($"performance_{_logFileName}");
                SafeFileLogger.SafeAppendText("placement_performance.log", 
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Performance report written to: {reportPath}\n");
            }
        }
        
        /// <summary>
        /// Generate summary section for individual sleeve placement
        /// </summary>
        private void GenerateIndividualPlacementSection(StringBuilder report, int totalIndividualSleeves)
        {
            report.AppendLine($"╔════════════════════════════════════════════════════════════════════════════════════╗");
            report.AppendLine($"║                    INDIVIDUAL BULK PLACEMENT PERFORMANCE                          ║");
            report.AppendLine($"╚════════════════════════════════════════════════════════════════════════════════════╝");
            report.AppendLine();
            
            // Filter operations related to individual placement
            var individualOps = _operations.Values
                .Where(op => IsIndividualPlacementOperation(op.Name))
                .OrderByDescending(o => o.TotalMilliseconds)
                .ToList();
            
            if (individualOps.Count == 0)
            {
                report.AppendLine("No individual placement operations recorded.");
                report.AppendLine();
                return;
            }
            
            // Calculate total time for individual operations
            long totalIndividualTime = individualOps.Sum(op => op.TotalMilliseconds);
            
            // Summary table
            report.AppendLine($"=== INDIVIDUAL PLACEMENT SUMMARY ===");
            report.AppendLine($"{"Metric",-30} {"Value",20}");
            report.AppendLine(new string('-', 52));
            report.AppendLine($"{"Total Sleeves Placed",-30} {totalIndividualSleeves,20}");
            report.AppendLine($"{"Total Time",-30} {totalIndividualTime,19}ms ({TimeSpan.FromMilliseconds(totalIndividualTime):mm\\:ss})");
            
            if (totalIndividualSleeves > 0 && totalIndividualTime > 0)
            {
                double sleevesPerSec = (double)totalIndividualSleeves / totalIndividualTime * 1000;
                double avgTimePerSleeve = (double)totalIndividualTime / totalIndividualSleeves;
                report.AppendLine($"{"Sleeves/Second",-30} {sleevesPerSec,19:F0}");
                report.AppendLine($"{"Avg Time per Sleeve",-30} {avgTimePerSleeve,19:F1}ms");
                
                bool meetsTarget = sleevesPerSec >= 50;
                report.AppendLine($"{"Performance Status",-30} {(meetsTarget ? "✅ MEETS TARGET (50+/s)" : "⚠️ BELOW TARGET (<50/s)"),20}");
            }
            report.AppendLine();
            
            // Operation breakdown table
            report.AppendLine($"=== INDIVIDUAL PLACEMENT OPERATION BREAKDOWN ===");
            report.AppendLine($"{"Operation",-45} {"Calls",8} {"Total",12} {"Avg",10} {"Min",10} {"Max",10} {"Items",10} {"Items/s",10} {"%",6}");
            report.AppendLine(new string('-', 120));
            
            foreach (var op in individualOps)
            {
                double avgMs = op.CallCount > 0 ? (double)op.TotalMilliseconds / op.CallCount : 0;
                double avgItemsPerSec = op.TotalMilliseconds > 0 
                    ? (double)op.TotalItemCount / op.TotalMilliseconds * 1000 
                    : 0;
                double percentage = totalIndividualTime > 0 
                    ? (double)op.TotalMilliseconds / totalIndividualTime * 100 
                    : 0;
                
                report.AppendLine($"{op.Name,-45} {op.CallCount,8} {op.TotalMilliseconds,12}ms {avgMs,9:F1}ms {op.MinMilliseconds,9}ms {op.MaxMilliseconds,9}ms {op.TotalItemCount,10} {avgItemsPerSec,9:F0}/s {percentage,5:F1}%");
            }
            
            report.AppendLine();
        }
        
        /// <summary>
        /// Generate summary section for cluster placement
        /// </summary>
        private void GenerateClusterPlacementSection(StringBuilder report, int totalClusters)
        {
            report.AppendLine($"╔════════════════════════════════════════════════════════════════════════════════════╗");
            report.AppendLine($"║                        CLUSTER BULK PLACEMENT PERFORMANCE                          ║");
            report.AppendLine($"╚════════════════════════════════════════════════════════════════════════════════════╝");
            report.AppendLine();
            
            // Filter operations related to cluster placement
            var clusterOps = _operations.Values
                .Where(op => IsClusterPlacementOperation(op.Name))
                .OrderByDescending(o => o.TotalMilliseconds)
                .ToList();
            
            if (clusterOps.Count == 0)
            {
                report.AppendLine("No cluster placement operations recorded.");
                report.AppendLine();
                return;
            }
            
            // Calculate total time for cluster operations
            // ✅ FIX: Show actual time even when no clusters placed - this helps identify why operations are slow
            long totalClusterTime = clusterOps.Sum(op => op.TotalMilliseconds);
            
            // Summary table
            report.AppendLine($"=== CLUSTER PLACEMENT SUMMARY ===");
            report.AppendLine($"{"Metric",-30} {"Value",20}");
            report.AppendLine(new string('-', 52));
            report.AppendLine($"{"Total Clusters Placed",-30} {totalClusters,20}");
            report.AppendLine($"{"Total Time",-30} {totalClusterTime,19}ms ({TimeSpan.FromMilliseconds(totalClusterTime):mm\\:ss})");
            
            // ✅ FIX: Show warning if time was spent but no clusters were placed
            if (totalClusters == 0 && totalClusterTime > 0)
            {
                report.AppendLine($"{"⚠️ WARNING",-30} {"Time spent but no clusters placed!",20}");
            }
            
            if (totalClusters > 0 && totalClusterTime > 0)
            {
                double clustersPerSec = (double)totalClusters / totalClusterTime * 1000;
                double avgTimePerCluster = (double)totalClusterTime / totalClusters;
                report.AppendLine($"{"Clusters/Second",-30} {clustersPerSec,19:F0}");
                report.AppendLine($"{"Avg Time per Cluster",-30} {avgTimePerCluster,19:F1}ms");
                
                bool meetsTarget = clustersPerSec >= 10;
                report.AppendLine($"{"Performance Status",-30} {(meetsTarget ? "✅ MEETS TARGET (10+/s)" : "⚠️ BELOW TARGET (<10/s)"),20}");
            }
            report.AppendLine();
            
            // Operation breakdown table
            report.AppendLine($"=== CLUSTER PLACEMENT OPERATION BREAKDOWN ===");
            report.AppendLine($"{"Operation",-45} {"Calls",8} {"Total",12} {"Avg",10} {"Min",10} {"Max",10} {"Items",10} {"Items/s",10} {"%",6}");
            report.AppendLine(new string('-', 120));
            
            foreach (var op in clusterOps)
            {
                double avgMs = op.CallCount > 0 ? (double)op.TotalMilliseconds / op.CallCount : 0;
                double avgItemsPerSec = op.TotalMilliseconds > 0 
                    ? (double)op.TotalItemCount / op.TotalMilliseconds * 1000 
                    : 0;
                double percentage = totalClusterTime > 0 
                    ? (double)op.TotalMilliseconds / totalClusterTime * 100 
                    : 0;
                
                report.AppendLine($"{op.Name,-45} {op.CallCount,8} {op.TotalMilliseconds,12}ms {avgMs,9:F1}ms {op.MinMilliseconds,9}ms {op.MaxMilliseconds,9}ms {op.TotalItemCount,10} {avgItemsPerSec,9:F0}/s {percentage,5:F1}%");
            }
            
            report.AppendLine();
        }
        
        /// <summary>
        /// Determine if an operation is related to individual placement
        /// </summary>
        private bool IsIndividualPlacementOperation(string operationName)
        {
            string name = operationName.ToLowerInvariant();
            return name.Contains("individual") ||
                   name.Contains("bulk individual") ||
                   name.Contains("step 1") ||
                   name.Contains("step 4") ||
                   name.Contains("step 5") ||
                   name.Contains("step 6") ||
                   name.Contains("load clash zones") ||
                   name.Contains("loading from db") ||
                   name.Contains("pre-filter") ||
                   name.Contains("pre-activate") ||
                   name.Contains("build creation") ||
                   name.Contains("newfamilyinstances") ||
                   name.Contains("apply rotation") ||
                   name.Contains("sync flags") ||
                   name.Contains("delete invalidated") ||
                   name.Contains("plan placement") ||
                   name.Contains("pre-save") ||
                   name.Contains("executebulkplacement") ||
                   name.Contains("regenerate") ||
                   name.Contains("flush deferred") ||
                   name.Contains("update zones") ||
                   name.Contains("transaction commit") ||
                   name.Contains("batchupdate") ||
                   name.Contains("savesleevesnapshots") ||
                   name.Contains("retrieved placed data") ||
                   name.Contains("save placed data");
        }
        
        /// <summary>
        /// Cluster-only operations. Excludes "Apply Rotation & Parameters" (used by BulkPlacementService for individual sleeves)
        /// so when EnableClusteringWorkflow is false we do not attribute that time to cluster.
        /// </summary>
        private bool IsClusterPlacementOperation(string operationName)
        {
            string name = operationName.ToLowerInvariant();
            return name.Contains("cluster") ||
                   name.Contains("prepare creation") ||
                   name.Contains("collect individual sleeves") ||
                   name.Contains("delete individual sleeves") ||
                   (name.Contains("newfamilyinstances2") && name.Contains("cluster")) ||
                   name.Contains("post-placement") ||
                   name.Contains("database updates") ||
                   name.Contains("cleanup");
        }
        
        private class OperationMetrics
        {
            public string Name { get; set; }
            public int CallCount { get; set; }
            public long TotalMilliseconds { get; set; }
            public long TotalMemoryBytes { get; set; }
            public int TotalItemCount { get; set; }
            public long MinMilliseconds { get; set; }
            public long MaxMilliseconds { get; set; }
        }
        
        public class OperationTracker : JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker
        {
            private readonly PlacementPerformanceMonitor _monitor;
            private readonly string _operationName;
            private readonly Stopwatch _timer;
            private readonly long _startMemory;
            private int _itemCount;
            private readonly Dictionary<string, OperationMetrics> _subOperations;
            
            public OperationTracker(PlacementPerformanceMonitor monitor, string operationName)
            {
                _monitor = monitor;
                _operationName = operationName;
                _timer = Stopwatch.StartNew();
                _startMemory = GC.GetTotalMemory(false);
                _subOperations = new Dictionary<string, OperationMetrics>();
            }
            
            public void SetItemCount(int count)
            {
                _itemCount = count;
            }
            
            /// <summary>
            /// Track a sub-operation within this operation
            /// </summary>
            /// <summary>
            /// Track a sub-operation within this operation
            /// </summary>
            public JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker TrackSubOperation(string subOperationName)
            {
                return new SubOperationTracker(this, subOperationName);
            }
            
            internal void RecordSubOperation(string name, long milliseconds, long memoryBytes, int itemCount)
            {
                if (!_subOperations.ContainsKey(name))
                {
                    _subOperations[name] = new OperationMetrics { Name = name };
                }
                
                var metrics = _subOperations[name];
                metrics.CallCount++;
                metrics.TotalMilliseconds += milliseconds;
                metrics.TotalMemoryBytes += memoryBytes;
                metrics.TotalItemCount += itemCount;
                
                if (milliseconds > metrics.MaxMilliseconds)
                    metrics.MaxMilliseconds = milliseconds;
                
                if (milliseconds < metrics.MinMilliseconds || metrics.MinMilliseconds == 0)
                    metrics.MinMilliseconds = milliseconds;
            }
            
            public void Dispose()
            {
                _timer.Stop();
                long endMemory = GC.GetTotalMemory(false);
                long memoryDelta = endMemory - _startMemory;
                
                _monitor.RecordOperation(_operationName, _timer.ElapsedMilliseconds, memoryDelta, _itemCount);
                
                // ✅ LOG OPERATION COMPLETION: Log each operation completion (ALWAYS log, even in deployment mode)
                SafeFileLogger.SafeAppendTextAlways($"performance_{_monitor._logFileName}",
                    $"[{DateTime.Now:HH:mm:ss.fff}] {_operationName}: {_timer.ElapsedMilliseconds}ms, Memory: {memoryDelta / 1024.0:F2} KB, Items: {_itemCount}\n");
                
                // Log sub-operations (ALWAYS log, even in deployment mode)
                SafeFileLogger.SafeAppendTextAlways($"performance_{_monitor._logFileName}",
                     $"[{DateTime.Now:HH:mm:ss.fff}] DEBUG: Disposing OperationTracker '{_operationName}'. SubOpCount: {_subOperations.Count}\n");

                foreach (var subOp in _subOperations.Values.OrderByDescending(o => o.TotalMilliseconds))
                {
                    double avgMs = subOp.CallCount > 0 ? (double)subOp.TotalMilliseconds / subOp.CallCount : 0;
                    SafeFileLogger.SafeAppendTextAlways($"performance_{_monitor._logFileName}",
                        $"[{DateTime.Now:HH:mm:ss.fff}]   (Sub) {subOp.Name}: {subOp.TotalMilliseconds}ms (avg: {avgMs:F1}ms, calls: {subOp.CallCount}, items: {subOp.TotalItemCount})\n");
                }
            }
            
            public class SubOperationTracker : JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker
            {
                private readonly OperationTracker _parent;
                private readonly string _subOperationName;
                private readonly Stopwatch _timer;
                private readonly long _startMemory;
                private int _itemCount;
                
                public SubOperationTracker(OperationTracker parent, string subOperationName)
                {
                    _parent = parent;
                    _subOperationName = subOperationName;
                    _timer = Stopwatch.StartNew();
                    _startMemory = GC.GetTotalMemory(false);
                }
                
                public void SetItemCount(int count)
                {
                    _itemCount = count;
                }

                /// <summary>
                /// Track a sub-operation (delegates to parent to keep hierarchy flat for now)
                /// </summary>
                public JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker TrackSubOperation(string subOperationName)
                {
                    // For now, flatten sub-operations by tracking them on the parent operation
                    return _parent.TrackSubOperation(subOperationName);
                }
                
                public void Dispose()
                {
                    _timer.Stop();
                    long endMemory = GC.GetTotalMemory(false);
                    long memoryDelta = endMemory - _startMemory;
                    
                    _parent.RecordSubOperation(_subOperationName, _timer.ElapsedMilliseconds, memoryDelta, _itemCount);
                }
            }
        }
    }
}

