using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized logging configuration with individual switches for different services
    /// </summary>
    public static class LoggingConfiguration
    {
        // GLOBAL SWITCH: Disable ALL hardcoded file logging
        public static bool DisableAllHardcodedLogging = false; // Temporarily enabled for DuctSleeveCommand debugging
        
        // Individual log switches - set to true only for the service you want to debug
        // ✅ REMOVED: EnableMainUI - Main UI logging removed per user request
        public static bool EnableOKButton = true;          // OK button operations (ENABLED for debugging)
        public static bool EnableRefreshButton = false;   // Refresh button operations
        public static bool EnableProfileManagement = false; // Profile save/restore (DISABLED)
        public static bool EnableClashDetection = false; // Clash zone detection
        public static bool EnableOrchestrator = true;    // Command orchestrator (ENABLED for debugging)
        public static bool EnableMepIntersection = false; // MEP intersection service
        public static bool EnableSleevePlacers = true;   // Sleeve placement services (ENABLED for DuctSleeveCommand debugging)
        public static bool EnableParameterTransfer = false; // Parameter operations
        public static bool EnableExternalEvents = false; // External event handlers (DISABLED to reduce profile logs)
        public static bool EnableProgressDialog = false;  // Progress dialog operations (DISABLED)
        
        // Log file paths - ✅ FIXED: Use SafeFileLogger for deployment-compatible paths
        private static string LogDirectory => SafeFileLogger.GetLogDirectory();
        private static string MainLogFile => SafeFileLogger.GetLogFilePath("main_debug.log");
        
        /// <summary>
        /// Gets the appropriate log file path based on the service
        /// </summary>
        public static string GetLogFilePath(string serviceName)
        {
            // ✅ REMOVED: MainUI log file path - Main UI logging removed per user request
            if (serviceName == "OKButton")
                return MainLogFile;
            
            return Path.Combine(LogDirectory, $"{serviceName.ToLower()}_debug.log");
        }
        
        /// <summary>
        /// Checks if logging is enabled for a specific service
        /// </summary>
        public static bool IsLoggingEnabled(string serviceName)
        {
            return serviceName switch
            {
                // ✅ REMOVED: "MainUI" => EnableMainUI, - Main UI logging removed per user request
                "OKButton" => EnableOKButton,
                "RefreshButton" => EnableRefreshButton,
                "ProfileManagement" => EnableProfileManagement,
                "ClashDetection" => EnableClashDetection,
                "Orchestrator" => EnableOrchestrator,
                "MepIntersection" => EnableMepIntersection,
                "SleevePlacers" => EnableSleevePlacers,
                "ParameterTransfer" => EnableParameterTransfer,
                "ExternalEvents" => EnableExternalEvents,
"ProgressDialog" => EnableProgressDialog,
                _ => false
            };
        }
        
        /// <summary>
        /// Checks if hardcoded file logging should be disabled
        /// </summary>
        public static bool ShouldDisableHardcodedLogging()
        {
            return DisableAllHardcodedLogging;
        }
        
        /// <summary>
        /// Conditionally appends text to a file only if hardcoded logging is enabled
        /// </summary>
        public static void ConditionalAppendAllText(string filePath, string content)
        {
            // ✅ DEPLOYMENT MODE: Skip all logging if deployment mode is enabled
            if (DeploymentConfiguration.DeploymentMode)
                return;
                
            if (!DisableAllHardcodedLogging)
            {
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(filePath, content);
                }
            }
        }
        
        /// <summary>
        /// Disables all logging except the specified service
        /// </summary>
        public static void FocusOnService(string serviceName)
        {
            // Disable all
            // ✅ REMOVED: EnableMainUI = false; - Main UI logging removed per user request
            EnableOKButton = false;
            EnableRefreshButton = false;
            EnableProfileManagement = false;
            EnableClashDetection = false;
            EnableOrchestrator = false;
            EnableMepIntersection = false;
            EnableSleevePlacers = false;
            EnableParameterTransfer = false;
            EnableExternalEvents = false;
            
            // Enable only the specified service
            switch (serviceName)
            {
                // ✅ REMOVED: case "MainUI": - Main UI logging removed per user request
                case "OKButton":
                    EnableOKButton = true;
                    // ✅ REMOVED: EnableMainUI = true; - Main UI logging removed per user request
                    break;
                case "RefreshButton":
                    EnableRefreshButton = true;
                    // ✅ REMOVED: EnableMainUI = true; - Main UI logging removed per user request
                    break;
                case "ProfileManagement":
                    EnableProfileManagement = true;
                    break;
                case "ClashDetection":
                    EnableClashDetection = true;
                    break;
                case "Orchestrator":
                    EnableOrchestrator = true;
                    break;
                case "MepIntersection":
                    EnableMepIntersection = true;
                    break;
                case "SleevePlacers":
                    EnableSleevePlacers = true;
                    break;
                case "ParameterTransfer":
                    EnableParameterTransfer = true;
                    break;
                case "ExternalEvents":
                    EnableExternalEvents = true;
                    break;
                case "ProgressDialog":
                    EnableProgressDialog = true;
                    break;
            }
        }
        
        /// <summary>
        /// Disables all logging
        /// </summary>
        public static void DisableAllLogging()
        {
            // ✅ REMOVED: EnableMainUI = false; - Main UI logging removed per user request
            EnableOKButton = false;
            EnableRefreshButton = false;
            EnableProfileManagement = false;
            EnableClashDetection = false;
            EnableOrchestrator = false;
            EnableMepIntersection = false;
            EnableSleevePlacers = false;
            EnableParameterTransfer = false;
            EnableExternalEvents = false;
            EnableProgressDialog = false;
        }
        
        /// <summary>
        /// Enables only critical logging (OK button and main UI)
        /// </summary>
        public static void EnableCriticalLoggingOnly()
        {
            DisableAllLogging();
            EnableOKButton = true;
            // ✅ REMOVED: EnableMainUI = true; - Main UI logging removed per user request
        }
    }
}
