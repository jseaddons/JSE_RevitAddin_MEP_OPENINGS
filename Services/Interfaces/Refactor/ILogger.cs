using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team D: Logger abstraction that wraps DebugLogger + SafeFileLogger.
    /// Provides structured logging capability with correlation IDs.
    /// Preserves existing log file names and message prefixes.
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Correlation ID for tracing operations across multiple log entries
        /// </summary>
        string CorrelationId { get; set; }
        
        /// <summary>
        /// Logs an informational message
        /// </summary>
        void Info(string message, string? scope = null);
        
        /// <summary>
        /// Logs a warning message
        /// </summary>
        void Warning(string message, string? scope = null);
        
        /// <summary>
        /// Logs an error message
        /// </summary>
        void Error(string message, Exception? exception = null, string? scope = null);
        
        /// <summary>
        /// Logs a debug message (only in development mode)
        /// </summary>
        void Debug(string message, string? scope = null);
        
        /// <summary>
        /// Logs structured data with key-value pairs
        /// </summary>
        void LogStructured(string level, string message, object? data = null, string? scope = null);
    }
}

