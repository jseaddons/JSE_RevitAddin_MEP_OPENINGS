using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team D: Error policy abstraction that decides log level, continuation, and aggregation.
    /// Centralizes error handling decisions instead of scattering try-catch blocks.
    /// </summary>
    public interface IErrorPolicy
    {
        /// <summary>
        /// Handles an exception according to the policy.
        /// </summary>
        /// <param name="exception">The exception that occurred</param>
        /// <param name="scope">Context where error occurred (e.g., "PlaceSleeveInstance", "PersistPlacement")</param>
        /// <param name="context">Optional context object (e.g., PlacementContext, ClashZone)</param>
        /// <param name="critical">Whether this is a critical error that should stop processing</param>
        /// <returns>Error handling result indicating whether to continue processing</returns>
        ErrorHandlingResult Handle(
            Exception exception,
            string scope,
            object? context = null,
            bool critical = false);
        
        /// <summary>
        /// Aggregates multiple errors for batch reporting.
        /// </summary>
        /// <param name="errors">Collection of error handling results</param>
        /// <returns>Aggregated error summary</returns>
        ErrorSummary AggregateErrors(IEnumerable<ErrorHandlingResult> errors);
    }
    
    /// <summary>
    /// Result of error handling operation
    /// </summary>
    public class ErrorHandlingResult
    {
        public bool ShouldContinue { get; set; } = true;
        public string LogLevel { get; set; } = "Error"; // Debug, Info, Warning, Error
        public string Message { get; set; } = string.Empty;
        public string Scope { get; set; } = string.Empty;
        public Exception? Exception { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
        public object? Context { get; set; }
    }
    
    /// <summary>
    /// Aggregated summary of multiple errors
    /// </summary>
    public class ErrorSummary
    {
        public int TotalErrors { get; set; }
        public int CriticalErrors { get; set; }
        public int Warnings { get; set; }
        public List<string> ErrorMessages { get; set; } = new List<string>();
        public bool ShouldAbort { get; set; }
    }
}

