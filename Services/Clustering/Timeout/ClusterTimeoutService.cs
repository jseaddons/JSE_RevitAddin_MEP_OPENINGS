using System;
using System.Diagnostics;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout
{
    /// <summary>
    /// Service for timeout protection during clustering operations.
    /// Prevents infinite loops and provides user feedback for long-running operations.
    /// </summary>
    public class ClusterTimeoutService : IClusterTimeoutService
    {
        private readonly Stopwatch _timer;
        private readonly int _timeoutLimitMs;

        public long ElapsedMilliseconds => _timer?.ElapsedMilliseconds ?? 0;
        public int TimeoutLimitMs => _timeoutLimitMs;

        /// <summary>
        /// Create a new timeout service with specified timeout limit.
        /// </summary>
        /// <param name="timeoutLimitMs">Timeout limit in milliseconds (default: 300000 = 5 minutes)</param>
        public ClusterTimeoutService(int timeoutLimitMs = 300000)
        {
            _timeoutLimitMs = timeoutLimitMs;
            _timer = new Stopwatch();
        }

        /// <summary>
        /// Start the timeout timer for clustering operation.
        /// </summary>
        public void StartTimer()
        {
            _timer.Restart();
        }

        /// <summary>
        /// Check if the operation has exceeded the timeout limit.
        /// </summary>
        /// <returns>True if timeout exceeded, false otherwise</returns>
        public bool IsTimedOut()
        {
            return _timer.ElapsedMilliseconds > _timeoutLimitMs;
        }

        /// <summary>
        /// Log timeout warning (no UI blocking - per user request to remove Phase 10 progress UI).
        /// </summary>
        /// <param name="context">Context message (e.g., "during FormClusters", "after processing N clusters")</param>
        /// <param name="processedCount">Number of items processed (optional)</param>
        /// <param name="totalCount">Total number of items (optional)</param>
        public void ShowTimeoutWarning(string context, int? processedCount = null, int? totalCount = null)
        {
            var timeInSeconds = _timer.ElapsedMilliseconds / 1000;
            var limitInSeconds = _timeoutLimitMs / 1000;

            var message = $"Clustering operation is taking too long and has been cancelled.\n";
            
            if (processedCount.HasValue && totalCount.HasValue)
            {
                message += $"Processed: {processedCount.Value} of {totalCount.Value} clusters\n";
            }
            
            message += $"Time: {timeInSeconds} seconds\n";
            message += $"Limit: {limitInSeconds} seconds\n";
            message += $"Context: {context}";

            // ✅ REMOVED: MessageBox.Show() - was blocking clustering process
            // ✅ CHANGED: Log to file instead (non-blocking)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[ClusterTimeoutService] {message}");
            }
            
            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] ⏱ TIMEOUT: {message}\n");
        }

        /// <summary>
        /// Reset the timer.
        /// </summary>
        public void Reset()
        {
            _timer.Reset();
        }
    }
}
