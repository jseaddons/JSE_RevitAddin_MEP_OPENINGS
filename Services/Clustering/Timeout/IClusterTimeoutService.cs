using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Timeout
{
    /// <summary>
    /// Service interface for timeout protection during clustering operations.
    /// Prevents infinite loops and provides user feedback for long-running operations.
    /// </summary>
    public interface IClusterTimeoutService
    {
        /// <summary>
        /// Start the timeout timer for clustering operation.
        /// </summary>
        void StartTimer();

        /// <summary>
        /// Check if the operation has exceeded the timeout limit.
        /// </summary>
        /// <returns>True if timeout exceeded, false otherwise</returns>
        bool IsTimedOut();

        /// <summary>
        /// Get elapsed time in milliseconds since timer started.
        /// </summary>
        long ElapsedMilliseconds { get; }

        /// <summary>
        /// Get timeout limit in milliseconds.
        /// </summary>
        int TimeoutLimitMs { get; }

        /// <summary>
        /// Show timeout warning dialog to user.
        /// </summary>
        /// <param name="context">Context message (e.g., "during FormClusters", "after processing N clusters")</param>
        /// <param name="processedCount">Number of items processed (optional)</param>
        /// <param name="totalCount">Total number of items (optional)</param>
        void ShowTimeoutWarning(string context, int? processedCount = null, int? totalCount = null);

        /// <summary>
        /// Reset the timer.
        /// </summary>
        void Reset();
    }
}
