using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Service for monitoring and logging performance metrics.
    /// Follows Single Responsibility Principle - only handles performance tracking.
    /// </summary>
    public interface IPerformanceMonitor
    {
        /// <summary>
        /// Start timing an operation.
        /// </summary>
        /// <param name="operationName">Name of the operation being timed</param>
        void StartOperation(string operationName);

        /// <summary>
        /// Stop timing an operation and log the result.
        /// </summary>
        /// <param name="operationName">Name of the operation</param>
        /// <param name="itemCount">Number of items processed (optional)</param>
        void StopOperation(string operationName, int itemCount = 0);

        /// <summary>
        /// Log a performance metric without timing.
        /// </summary>
        /// <param name="metricName">Name of the metric</param>
        /// <param name="value">Value to log</param>
        void LogMetric(string metricName, object value);

        /// <summary>
        /// Check if performance monitoring is enabled.
        /// </summary>
        /// <summary>
        /// Track an operation with a using block for automatic timing.
        /// </summary>
        /// <param name="operationName">Name of the operation</param>
        /// <returns>Disposable tracker that logs completion on disposal</returns>
        IOperationTracker TrackOperation(string operationName);

        /// <summary>
        /// Generate final performance report.
        /// </summary>
        void GenerateReport(int totalIndividualSleeves, int totalClusters);

        /// <summary>
        /// Check if performance monitoring is enabled.
        /// </summary>
        bool IsEnabled { get; }
    }

    /// <summary>
    /// Tracker for a single operation, allows updating item counts.
    /// </summary>
    public interface IOperationTracker : IDisposable
    {
        /// <summary>
        /// update the item count for this operation
        /// </summary>
        void SetItemCount(int count);

        /// <summary>
        /// Track a sub-operation within this operation.
        /// </summary>
        IOperationTracker TrackSubOperation(string subOperationName);
    }
}
