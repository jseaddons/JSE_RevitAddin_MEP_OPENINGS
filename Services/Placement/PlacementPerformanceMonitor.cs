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
    public class PlacementPerformanceMonitor
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
            
            // ✅ INITIALIZE LOG: Log start time
            SafeFileLogger.SafeAppendText($"performance_{_logFileName}", 
                $"=== PLACEMENT PERFORMANCE MONITOR STARTED ===\n" +
                $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                $"Log File: {_logFileName}\n" +
                $"Start Memory: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB\n\n");
        }
        
        /// <summary>
        /// Start tracking an operation
        /// </summary>
        public OperationTracker TrackOperation(string operationName)
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
        /// Generate final performance report
        /// </summary>
        public void GenerateReport(int totalIndividualSleeves, int totalClusters)
        {
            _totalTimer.Stop();
            
            var report = new StringBuilder();
            report.AppendLine($"=== PLACEMENT PERFORMANCE REPORT ===");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Total Time: {_totalTimer.ElapsedMilliseconds}ms ({_totalTimer.Elapsed:mm\\:ss})");
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
            
            // Operation breakdown
            report.AppendLine($"=== OPERATION BREAKDOWN ===");
            report.AppendLine($"{"Operation",-40} {"Calls",8} {"Total",12} {"Avg",10} {"Min",10} {"Max",10} {"Items",10} {"Items/s",10}");
            report.AppendLine(new string('-', 110));
            
            foreach (var op in _operations.Values.OrderByDescending(o => o.TotalMilliseconds))
            {
                double avgMs = op.CallCount > 0 ? (double)op.TotalMilliseconds / op.CallCount : 0;
                double avgItemsPerSec = op.TotalMilliseconds > 0 
                    ? (double)op.TotalItemCount / op.TotalMilliseconds * 1000 
                    : 0;
                
                report.AppendLine($"{op.Name,-40} {op.CallCount,8} {op.TotalMilliseconds,12}ms {avgMs,9:F1}ms {op.MinMilliseconds,9}ms {op.MaxMilliseconds,9}ms {op.TotalItemCount,10} {avgItemsPerSec,9:F0}/s");
            }
            
            report.AppendLine();
            
            // Performance targets
            report.AppendLine($"=== PERFORMANCE TARGETS ===");
            
            if (totalIndividualSleeves > 0)
            {
                double sleevesPerSec = (double)totalIndividualSleeves / _totalTimer.ElapsedMilliseconds * 1000;
                report.AppendLine($"Individual Sleeves/Second: {sleevesPerSec:F0} (target: 50+)");
                
                bool meetsTarget = sleevesPerSec >= 50;
                report.AppendLine($"Status: {(meetsTarget ? "✅ MEETS TARGET" : "⚠️ BELOW TARGET")}");
            }
            
            if (totalClusters > 0)
            {
                double clustersPerSec = (double)totalClusters / _totalTimer.ElapsedMilliseconds * 1000;
                report.AppendLine($"Clusters/Second: {clustersPerSec:F0} (target: 10+)");
                
                bool meetsTarget = clustersPerSec >= 10;
                report.AppendLine($"Status: {(meetsTarget ? "✅ MEETS TARGET" : "⚠️ BELOW TARGET")}");
            }
            
            report.AppendLine();
            report.AppendLine($"=== END OF REPORT ===");
            
            // Write report using SafeFileLogger
            SafeFileLogger.SafeAppendText($"performance_{_logFileName}", report.ToString());
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string reportPath = SafeFileLogger.GetLogFilePath($"performance_{_logFileName}");
                SafeFileLogger.SafeAppendText("placement_performance.log", 
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Performance report written to: {reportPath}\n");
            }
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
        
        public class OperationTracker : IDisposable
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
            public SubOperationTracker TrackSubOperation(string subOperationName)
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
                
                // ✅ LOG OPERATION COMPLETION: Log each operation completion
                SafeFileLogger.SafeAppendText($"performance_{_monitor._logFileName}",
                    $"[{DateTime.Now:HH:mm:ss.fff}] {_operationName}: {_timer.ElapsedMilliseconds}ms, Memory: {memoryDelta / 1024.0:F2} KB, Items: {_itemCount}\n");
                
                // Log sub-operations
                foreach (var subOp in _subOperations.Values.OrderByDescending(o => o.TotalMilliseconds))
                {
                    double avgMs = subOp.CallCount > 0 ? (double)subOp.TotalMilliseconds / subOp.CallCount : 0;
                    SafeFileLogger.SafeAppendText($"performance_{_monitor._logFileName}",
                        $"[{DateTime.Now:HH:mm:ss.fff}]   {subOp.Name}: {subOp.TotalMilliseconds}ms (avg: {avgMs:F1}ms, calls: {subOp.CallCount}, items: {subOp.TotalItemCount})\n");
                }
            }
            
            public class SubOperationTracker : IDisposable
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

