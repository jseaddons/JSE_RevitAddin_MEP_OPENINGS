using System;
using System.Diagnostics;
using System.Text;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Performance monitor for parameter operations (Parameter Transfer, Apply Marks, Remark Selected)
    /// </summary>
    public class ParameterOperationPerformanceMonitor : IDisposable
    {
        private readonly string _operationName;
        private readonly string _logFileName;
        private readonly Stopwatch _timer;
        private readonly long _startMemoryBytes;
        private int _itemCount;
        private bool _disposed = false;
        
        public ParameterOperationPerformanceMonitor(string operationName)
        {
            _operationName = operationName;
            _timer = Stopwatch.StartNew();
            _startMemoryBytes = GC.GetTotalMemory(false);
            
            // Generate log file name with timestamp
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            _logFileName = $"performance_ParameterOperation_{operationName.Replace(" ", "_")}_{timestamp}.log";
            
            // ✅ INITIALIZE LOG: Log start time
            SafeFileLogger.SafeAppendText(_logFileName, 
                $"=== {_operationName.ToUpper()} PERFORMANCE MONITOR STARTED ===\n" +
                $"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                $"Operation: {_operationName}\n" +
                $"Start Memory: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB\n\n");
        }
        
        public void SetItemCount(int count)
        {
            _itemCount = count;
        }
        
        public void Dispose()
        {
            if (_disposed) return;
            
            _timer.Stop();
            long endMemory = GC.GetTotalMemory(false);
            long memoryDelta = endMemory - _startMemoryBytes;
            
            var report = new StringBuilder();
            report.AppendLine($"=== {_operationName.ToUpper()} PERFORMANCE REPORT ===");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Operation: {_operationName}");
            report.AppendLine($"Total Time: {_timer.ElapsedMilliseconds}ms ({_timer.Elapsed:mm\\:ss\\.fff})");
            report.AppendLine($"Items Processed: {_itemCount}");
            report.AppendLine();
            
            // Memory summary
            report.AppendLine($"=== MEMORY USAGE ===");
            report.AppendLine($"Start: {_startMemoryBytes / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"End: {endMemory / 1024.0 / 1024.0:F2} MB");
            report.AppendLine($"Delta: {memoryDelta / 1024.0 / 1024.0:F2} MB");
            
            if (_itemCount > 0)
            {
                double memoryPerItem = (double)memoryDelta / _itemCount / 1024.0; // KB
                double itemsPerSecond = _timer.ElapsedMilliseconds > 0 
                    ? (double)_itemCount / _timer.ElapsedMilliseconds * 1000 
                    : 0;
                double timePerItem = _itemCount > 0 
                    ? (double)_timer.ElapsedMilliseconds / _itemCount 
                    : 0;
                
                report.AppendLine($"Memory per Item: {memoryPerItem:F2} KB");
                report.AppendLine($"Items per Second: {itemsPerSecond:F0}");
                report.AppendLine($"Time per Item: {timePerItem:F1} ms");
            }
            
            report.AppendLine();
            report.AppendLine($"=== END OF REPORT ===");
            
            // Write report
            SafeFileLogger.SafeAppendText(_logFileName, report.ToString());
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string reportPath = SafeFileLogger.GetLogFilePath(_logFileName);
                DebugLogger.Info($"[{_operationName}] Performance report written to: {reportPath}");
            }
            
            _disposed = true;
        }
    }
}

