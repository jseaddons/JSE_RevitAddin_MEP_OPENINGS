using System;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Configuration
{
    /// <summary>
    /// Team B: Adapter that reads from current DeploymentConfiguration into IAppEnvironment.
    /// Provides coexistence with existing static configuration - does not delete or alter DeploymentConfiguration.
    /// </summary>
    public class AppEnvironmentAdapter : IAppEnvironment
    {
        private readonly bool _isDeploymentMode;
        private readonly string _appDataDirectory;
        private readonly string _logDirectory;
        private readonly string _databaseDirectory;
        private readonly LoggingPolicy _loggingPolicy;
        
        /// <summary>
        /// Creates an adapter that reads from current DeploymentConfiguration static class.
        /// This maintains backward compatibility while enabling DI.
        /// </summary>
        public AppEnvironmentAdapter()
        {
            // ✅ COEXISTENCE: Read from existing DeploymentConfiguration (no changes to static class)
            _isDeploymentMode = DeploymentConfiguration.DeploymentMode;
            
            // Get paths from existing utilities (SafeFileLogger, etc.)
            // These paths are typically in %APPDATA%\JSE_MEP_Openings
            string appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "JSE_MEP_Openings"
            );
            
            _appDataDirectory = appDataPath;
            _logDirectory = Path.Combine(appDataPath, "Logs");
            _databaseDirectory = Path.Combine(appDataPath, "Database");
            
            // ✅ LOGGING POLICY: Based on deployment mode
            _loggingPolicy = _isDeploymentMode 
                ? LoggingPolicy.DeploymentMode 
                : LoggingPolicy.DevelopmentMode;
        }
        
        /// <summary>
        /// Creates an environment with explicit values (for testing or custom scenarios)
        /// </summary>
        public AppEnvironmentAdapter(
            bool isDeploymentMode,
            string appDataDirectory,
            string logDirectory,
            string databaseDirectory,
            LoggingPolicy loggingPolicy)
        {
            _isDeploymentMode = isDeploymentMode;
            _appDataDirectory = appDataDirectory ?? throw new ArgumentNullException(nameof(appDataDirectory));
            _logDirectory = logDirectory ?? throw new ArgumentNullException(nameof(logDirectory));
            _databaseDirectory = databaseDirectory ?? throw new ArgumentNullException(nameof(databaseDirectory));
            _loggingPolicy = loggingPolicy ?? throw new ArgumentNullException(nameof(loggingPolicy));
        }
        
        public bool IsDeploymentMode => _isDeploymentMode;
        public string AppDataDirectory => _appDataDirectory;
        public string LogDirectory => _logDirectory;
        public string DatabaseDirectory => _databaseDirectory;
        public LoggingPolicy LoggingPolicy => _loggingPolicy;
    }
    
    /// <summary>
    /// Factory for creating IAppEnvironment instances.
    /// Provides default implementation that reads from DeploymentConfiguration.
    /// </summary>
    public static class AppEnvironmentFactory
    {
        private static IAppEnvironment? _defaultInstance;
        private static readonly object _lock = new object();
        
        /// <summary>
        /// Gets the default environment instance (reads from DeploymentConfiguration).
        /// Thread-safe singleton pattern.
        /// </summary>
        public static IAppEnvironment Default
        {
            get
            {
                if (_defaultInstance == null)
                {
                    lock (_lock)
                    {
                        if (_defaultInstance == null)
                        {
                            _defaultInstance = new AppEnvironmentAdapter();
                        }
                    }
                }
                return _defaultInstance;
            }
        }
        
        /// <summary>
        /// Creates a custom environment instance (for testing or special scenarios).
        /// </summary>
        public static IAppEnvironment Create(
            bool isDeploymentMode,
            string appDataDirectory,
            string logDirectory,
            string databaseDirectory,
            LoggingPolicy loggingPolicy)
        {
            return new AppEnvironmentAdapter(
                isDeploymentMode,
                appDataDirectory,
                logDirectory,
                databaseDirectory,
                loggingPolicy);
        }
        
        /// <summary>
        /// Resets the default instance (useful for testing).
        /// </summary>
        public static void Reset()
        {
            lock (_lock)
            {
                _defaultInstance = null;
            }
        }
    }
}

