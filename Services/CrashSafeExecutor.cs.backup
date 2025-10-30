using System;
using System.Diagnostics;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Provides crash-safe execution with timeout protection and error handling.
    /// Ensures operations never hang indefinitely and always fail gracefully.
    /// </summary>
    public class CrashSafeExecutor
    {
        private const int MAX_EXECUTION_TIME_MS = 300000; // 5 minutes
        private readonly Stopwatch _stopwatch;
        private DateTime _operationStartTime;
        
        public CrashSafeExecutor()
        {
            _stopwatch = new Stopwatch();
        }
        
        /// <summary>
        /// Execute an operation with timeout protection and comprehensive error handling.
        /// </summary>
        public Result ExecuteWithTimeout(Func<Result> operation, string operationName)
        {
            _stopwatch.Restart();
            _operationStartTime = DateTime.Now;
            
            try
            {
                DebugLogger.Info($"[CrashSafe] ▶ Starting: {operationName}");
                
                var result = operation();
                
                _stopwatch.Stop();
                
                if (result == Result.Succeeded)
                {
                    DebugLogger.Info($"[CrashSafe] ✓ Completed: {operationName} in {_stopwatch.ElapsedMilliseconds}ms");
                }
                else if (result == Result.Cancelled)
                {
                    DebugLogger.Warning($"[CrashSafe] ⚠ Cancelled: {operationName} after {_stopwatch.ElapsedMilliseconds}ms");
                }
                else
                {
                    DebugLogger.Warning($"[CrashSafe] ✗ Failed: {operationName} after {_stopwatch.ElapsedMilliseconds}ms");
                }
                
                return result;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                _stopwatch.Stop();
                DebugLogger.Warning($"[CrashSafe] User cancelled: {operationName} after {_stopwatch.ElapsedMilliseconds}ms");
                
                MessageBox.Show(
                    $"Operation '{operationName}' was cancelled by user.",
                    "Operation Cancelled",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                return Result.Cancelled;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                _stopwatch.Stop();
                DebugLogger.Error($"[CrashSafe] Invalid operation in {operationName} after {_stopwatch.ElapsedMilliseconds}ms: {ex.Message}");
                
                MessageBox.Show(
                    $"Operation '{operationName}' failed due to invalid state.\n\nError: {ex.Message}\n\nThis usually happens when:\n- Document is in read-only mode\n- Element was deleted\n- Transaction conflict occurred",
                    "Invalid Operation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                
                return Result.Failed;
            }
            catch (Exception ex)
            {
                _stopwatch.Stop();
                DebugLogger.Error($"[CrashSafe] EXCEPTION in {operationName} after {_stopwatch.ElapsedMilliseconds}ms:");
                DebugLogger.Error($"[CrashSafe]   Message: {ex.Message}");
                DebugLogger.Error($"[CrashSafe]   Type: {ex.GetType().Name}");
                DebugLogger.Error($"[CrashSafe]   StackTrace: {ex.StackTrace}");
                
                MessageBox.Show(
                    $"Operation '{operationName}' failed with an error.\n\nError: {ex.Message}\n\nType: {ex.GetType().Name}\n\nPlease check the log file for details.",
                    "Operation Failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                
                return Result.Failed;
            }
        }
        
        /// <summary>
        /// Check if operation has exceeded timeout limit.
        /// Returns true if timeout occurred (operation should stop).
        /// </summary>
        public bool CheckTimeout(string operationPhase)
        {
            if (_stopwatch.ElapsedMilliseconds > MAX_EXECUTION_TIME_MS)
            {
                _stopwatch.Stop();
                
                DebugLogger.Error($"[CrashSafe] ⏱ TIMEOUT: {operationPhase} exceeded {MAX_EXECUTION_TIME_MS / 1000} second limit");
                DebugLogger.Error($"[CrashSafe]   Elapsed: {_stopwatch.ElapsedMilliseconds}ms");
                DebugLogger.Error($"[CrashSafe]   Started: {_operationStartTime:HH:mm:ss}");
                
                MessageBox.Show(
                    $"Operation is taking too long and has been cancelled.\n\nPhase: {operationPhase}\nTime: {_stopwatch.ElapsedMilliseconds / 1000} seconds\nLimit: {MAX_EXECUTION_TIME_MS / 1000} seconds\n\nThis usually indicates:\n• Very large model\n• No filter selected\n• Infinite loop\n• Section box too large\n\nPlease:\n• Select a specific filter\n• Reduce section box size\n• Process in smaller batches",
                    "Operation Timeout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// Get elapsed time in milliseconds since operation started.
        /// </summary>
        public long GetElapsedMilliseconds()
        {
            return _stopwatch.ElapsedMilliseconds;
        }
        
        /// <summary>
        /// Get elapsed time formatted as string.
        /// </summary>
        public string GetElapsedTimeFormatted()
        {
            var elapsed = _stopwatch.Elapsed;
            
            if (elapsed.TotalMinutes >= 1)
            {
                return $"{elapsed.Minutes}m {elapsed.Seconds}s";
            }
            else if (elapsed.TotalSeconds >= 1)
            {
                return $"{elapsed.Seconds}.{elapsed.Milliseconds:D3}s";
            }
            else
            {
                return $"{elapsed.Milliseconds}ms";
            }
        }
    }
}

