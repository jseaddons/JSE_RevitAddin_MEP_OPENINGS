using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Tracks performance metrics for refresh operation
    /// </summary>
    public class PerformanceMonitor
    {
        private readonly string _logFileName;
        private readonly Stopwatch _totalTimer;
        private readonly Dictionary<string, OperationMetrics> _operations;
        private long _startMemoryBytes;
        
        public PerformanceMonitor(string logFileName)
        {
            _logFileName = logFileName;
            _totalTimer = Stopwatch.StartNew();
            _operations = new Dictionary<string, OperationMetrics>();
            _startMemoryBytes = GC.GetTotalMemory(false);
        }
        
        /// <summary>
        /// Start tracking an operation
        /// </summary>
        public IDisposable TrackOperation(string operationName)
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
        public void GenerateReport(int totalClashZones)
        {
            _totalTimer.Stop();
            
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
            
            var report = new StringBuilder();
            report.AppendLine($"=== REFRESH PERFORMANCE REPORT ===");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"🔨 Build Timestamp: {buildTimestamp} | Assembly: {System.IO.Path.GetFileName(assemblyPath)}");
            report.AppendLine($"Total Time: {_totalTimer.ElapsedMilliseconds}ms ({_totalTimer.Elapsed:mm\\:ss})");
            report.AppendLine($"Total Clash Zones: {totalClashZones}");
            report.AppendLine();
            
            // Memory summary
            long currentMemory = GC.GetTotalMemory(false);
            long memoryDelta = currentMemory - _startMemoryBytes;
            report.AppendLine($"=== MEMORY USAGE ===");
            report.AppendLine($"Start: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"End: {currentMemory / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"Delta: {memoryDelta / 1024.0 / 1024.0:F2} MB");
            
            if (totalClashZones > 0)
            {
                double memoryPerZone = (double)memoryDelta / totalClashZones / 1024.0; // KB
                report.AppendLine($"Memory per Zone: {memoryPerZone:F2} KB (target: <5 KB)");
            }
            report.AppendLine();
            
            // Operation breakdown
            report.AppendLine($"=== OPERATION BREAKDOWN ===");
            report.AppendLine($"{"Operation",-30} {"Calls",8} {"Total",10} {"Avg",10} {"Min",10} {"Max",10} {"Items",10}");
            report.AppendLine(new string('-', 100));
            
            foreach (var op in _operations.Values.OrderByDescending(o => o.TotalMilliseconds))
            {
                double avgMs = op.CallCount > 0 ? (double)op.TotalMilliseconds / op.CallCount : 0;
                double avgItemsPerSec = op.TotalMilliseconds > 0 
                    ? (double)op.TotalItemCount / op.TotalMilliseconds * 1000 
                    : 0;
                
                report.AppendLine($"{op.Name,-30} {op.CallCount,8} {op.TotalMilliseconds,10}ms {avgMs,9:F1}ms {op.MinMilliseconds,9}ms {op.MaxMilliseconds,9}ms {avgItemsPerSec,9:F0}/s");
            }
            
            report.AppendLine();
            
            // Performance targets
            report.AppendLine($"=== PERFORMANCE TARGETS ===");
            
            if (totalClashZones > 0)
            {
                double zonesPerSec = (double)totalClashZones / _totalTimer.ElapsedMilliseconds * 1000;
                report.AppendLine($"Clash Zones/Second: {zonesPerSec:F0} (target: 500+)");
                
                bool meetsTarget = zonesPerSec >= 500;
                report.AppendLine($"Status: {(meetsTarget ? "✅ MEETS TARGET" : "⚠️ BELOW TARGET")}");
            }
            
            // Write report - bypass deployment mode check for performance reports
            string reportPath = SafeFileLogger.GetLogFilePath($"performance_{_logFileName}");
            SafeFileLogger.SafeAppendTextAlways($"performance_{_logFileName}", report.ToString());
            
            // ✅ NEW: Generate CSV report for Excel analysis
            GenerateCsvReport(totalClashZones);

            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[PERFORMANCE] Report written to: {reportPath}");
        }

        /// <summary>
        /// Generate CSV performance report for Excel analysis
        /// </summary>
        public void GenerateCsvReport(int totalClashZones)
        {
            try
            {
                var csv = new StringBuilder();
                csv.AppendLine("Operation,Calls,Total(ms),Avg(ms),Min(ms),Max(ms),Items,Items/s");

                foreach (var op in _operations.Values.OrderByDescending(o => o.TotalMilliseconds))
                {
                    double avgMs = op.CallCount > 0 ? (double)op.TotalMilliseconds / op.CallCount : 0;
                    double itemsPerSec = op.TotalMilliseconds > 0 
                        ? (double)op.TotalItemCount / op.TotalMilliseconds * 1000 
                        : 0;

                    csv.AppendLine($"{op.Name},{op.CallCount},{op.TotalMilliseconds},{avgMs:F2},{op.MinMilliseconds},{op.MaxMilliseconds},{op.TotalItemCount},{itemsPerSec:F0}");
                }

                string csvFileName = $"performance_{_logFileName.Replace(".log", "")}.csv";
                string csvPath = SafeFileLogger.GetLogFilePath(csvFileName);
                
                // Clear file if exists and write header
                if (System.IO.File.Exists(csvPath)) System.IO.File.Delete(csvPath);
                SafeFileLogger.SafeAppendTextAlways(csvFileName, csv.ToString());
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[PERFORMANCE] CSV Report written to: {csvPath}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PERFORMANCE] Failed to generate CSV report: {ex.Message}");
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
            private readonly PerformanceMonitor _monitor;
            private readonly string _operationName;
            private readonly Stopwatch _timer;
            private readonly long _startMemory;
            private int _itemCount;
            
            public OperationTracker(PerformanceMonitor monitor, string operationName)
            {
                _monitor = monitor;
                _operationName = operationName;
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
                
                _monitor.RecordOperation(_operationName, _timer.ElapsedMilliseconds, memoryDelta, _itemCount);
            }
        }
    }
}
