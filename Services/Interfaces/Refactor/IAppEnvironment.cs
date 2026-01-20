using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team B: Application environment abstraction.
    /// Encapsulates deployment mode, paths, and logging policy.
    /// Provides coexistence with existing DeploymentConfiguration static class.
    /// </summary>
    public interface IAppEnvironment
    {
        /// <summary>
        /// Whether the application is running in deployment mode.
        /// In deployment mode, verbose logging and performance monitoring are disabled.
        /// </summary>
        bool IsDeploymentMode { get; }
        
        /// <summary>
        /// Base directory for application data (logs, databases, etc.)
        /// </summary>
        string AppDataDirectory { get; }
        
        /// <summary>
        /// Directory for log files
        /// </summary>
        string LogDirectory { get; }
        
        /// <summary>
        /// Directory for database files
        /// </summary>
        string DatabaseDirectory { get; }
        
        /// <summary>
        /// Logging policy that determines what should be logged
        /// </summary>
        LoggingPolicy LoggingPolicy { get; }
    }
    
    /// <summary>
    /// Defines what types of logging are enabled.
    /// Immutable class (read-only properties) for .NET Framework 4.8 compatibility.
    /// </summary>
    public class LoggingPolicy
    {
        public bool EnableDebugLogging { get; }
        public bool EnablePerformanceLogging { get; }
        public bool EnableErrorLogging { get; }
        public bool EnableInfoLogging { get; }
        public bool EnableWarningLogging { get; }
        
        public LoggingPolicy(
            bool enableDebugLogging,
            bool enablePerformanceLogging,
            bool enableErrorLogging,
            bool enableInfoLogging,
            bool enableWarningLogging)
        {
            EnableDebugLogging = enableDebugLogging;
            EnablePerformanceLogging = enablePerformanceLogging;
            EnableErrorLogging = enableErrorLogging;
            EnableInfoLogging = enableInfoLogging;
            EnableWarningLogging = enableWarningLogging;
        }
        
        /// <summary>
        /// Default logging policy for deployment mode (minimal logging)
        /// </summary>
        public static LoggingPolicy DeploymentMode => new LoggingPolicy(
            enableDebugLogging: false,
            enablePerformanceLogging: false,
            enableErrorLogging: true,  // Always log errors
            enableInfoLogging: false,
            enableWarningLogging: true  // Always log warnings
        );
        
        /// <summary>
        /// Default logging policy for development mode (full logging)
        /// </summary>
        public static LoggingPolicy DevelopmentMode => new LoggingPolicy(
            enableDebugLogging: true,
            enablePerformanceLogging: true,
            enableErrorLogging: true,
            enableInfoLogging: true,
            enableWarningLogging: true
        );
    }
}

