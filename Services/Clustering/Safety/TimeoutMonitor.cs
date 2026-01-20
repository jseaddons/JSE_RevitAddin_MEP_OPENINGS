using System;
using System.Diagnostics;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety
{
    /// <summary>
    /// Monitors elapsed time for clustering operations and enforces timeouts.
    /// Phase 2: Timeout protection with 5-minute timeout per group (configurable).
    /// </summary>
    public class TimeoutMonitor
    {
        private readonly Stopwatch _stopwatch;
        private readonly TimeSpan _timeout;
        private readonly int _checkInterval;
        private int _iterationCount;

        /// <summary>
        /// Create a timeout monitor with specified timeout duration.
        /// </summary>
        /// <param name="timeoutMinutes">Timeout in minutes (default: 5 minutes)</param>
        /// <param name="checkInterval">Check timeout every N iterations (default: 100)</param>
        public TimeoutMonitor(double timeoutMinutes = 5.0, int checkInterval = 100)
        {
            _stopwatch = Stopwatch.StartNew();
            _timeout = TimeSpan.FromMinutes(timeoutMinutes);
            _checkInterval = checkInterval;
            _iterationCount = 0;
        }

        /// <summary>
        /// Check if timeout has been exceeded.
        /// Should be called periodically (every N iterations) to avoid performance overhead.
        /// </summary>
        /// <param name="operationName">Name of the operation for logging (optional)</param>
        /// <returns>True if timeout exceeded, false otherwise</returns>
        public bool IsTimeoutExceeded(string operationName = null)
        {
            _iterationCount++;

            // ✅ PERFORMANCE OPTIMIZATION: Check timeout every N iterations (not every iteration)
            if (_iterationCount % _checkInterval != 0)
            {
                return false;
            }

            if (_stopwatch.Elapsed >= _timeout)
            {
                string message = $"[TimeoutMonitor] ⚠️ TIMEOUT EXCEEDED: Operation '{operationName ?? "unknown"}' exceeded {_timeout.TotalMinutes:F1} minute timeout after {_iterationCount} iterations";
                SafeFileLogger.SafeAppendText("geometry_errors.log", message);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Check if timeout has been exceeded (does not increment iteration counter).
        /// Use this for immediate checks without affecting iteration tracking.
        /// </summary>
        /// <returns>True if timeout exceeded, false otherwise</returns>
        public bool HasTimeoutExceeded()
        {
            return _stopwatch.Elapsed >= _timeout;
        }

        /// <summary>
        /// Get elapsed time since monitor started.
        /// </summary>
        public TimeSpan Elapsed => _stopwatch.Elapsed;

        /// <summary>
        /// Get remaining time before timeout.
        /// </summary>
        public TimeSpan RemainingTime => _timeout - _stopwatch.Elapsed;

        /// <summary>
        /// Reset the timeout monitor (restart timer).
        /// </summary>
        public void Reset()
        {
            _stopwatch.Restart();
            _iterationCount = 0;
        }

        /// <summary>
        /// Stop the timeout monitor.
        /// </summary>
        public void Stop()
        {
            _stopwatch.Stop();
        }

        /// <summary>
        /// Get the number of iterations checked so far.
        /// </summary>
        public int IterationCount => _iterationCount;
    }
}

