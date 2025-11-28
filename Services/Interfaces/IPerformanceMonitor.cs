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
        bool IsEnabled { get; }
    }
}
