using System;
using System.IO;
using System.Reflection;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Dedicated logger for structural elements sleeve placement with timestamps
    /// </summary>
    public static class StructuralElementLogger
    {
        private static string? LogFilePath;
        private static bool IsInitialized = false;

        /// <summary>
        /// Initialize the structural elements log file in the project's Logs folder
        /// </summary>
        public static void InitializeLogger()
        {
            if (!DebugLogger.IsEnabled) return;
            // ...existing code...
            try
            {
                // Try to find the project directory more reliably
            string? projectDir = null;
                
                // Method 1: Look for the known project directory structure
                string[] potentialPaths = {
                    @"c:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS",
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                };
                
                foreach (string path in potentialPaths)
                {
                    if (Directory.Exists(path))
                    {
                        projectDir = path;
                        break;
                    }
                }
                
                // Method 2: Navigate up from assembly location to find .csproj file
                if (projectDir == null)
                {
                    string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                    projectDir = Path.GetDirectoryName(assemblyLocation);
                    
                    // Navigate up to find the project root (containing .csproj file)
                    int maxLevels = 5; // Prevent infinite loop
                    while (projectDir != null && Directory.GetFiles(projectDir, "*.csproj").Length == 0 && maxLevels > 0)
                    {
                        projectDir = Directory.GetParent(projectDir)?.FullName;
                        maxLevels--;
                    }
                }

                if (projectDir == null || !Directory.Exists(projectDir))
                {
                    // Fallback to Documents folder
                    projectDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                // Create Logs folder in project directory
                string logsDir = Path.Combine(projectDir!, "Logs");
                if (!Directory.Exists(logsDir))
                {
                    Directory.CreateDirectory(logsDir);
                }

                // Create timestamped log filename
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string logName = $"StructuralElements_SleeveLog_{timestamp}.log";
                LogFilePath = logName;

                // Write header
                string header = 
                    $"===== STRUCTURAL ELEMENTS SLEEVE PLACEMENT LOG =====\n" +
                    $"Session Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                    $"Log File: {logName}\n" +
                    $"JSE RevitAddin MEP Openings - Structural Elements Support\n" +
                    $"====================================================\n";

                SafeFileLogger.SafeAppendTextAlways(logName, header);
                IsInitialized = true;

                // Also log to main debug logger
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL-LOG] Initialized structural elements log: {LogFilePath}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL-LOG] Failed to initialize structural logger: {ex.Message}");
            }
        }

        /// <summary>
        /// Log structural element detection and processing
        /// </summary>
        public static void LogStructuralElement(string elementType, Autodesk.Revit.DB.ElementId elementId, string action, string details = "")
        {
            if (!DebugLogger.IsEnabled) return;
            if (!IsInitialized)
                InitializeLogger();

            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string logEntry = $"[{timestamp}] {elementType.ToUpper()} ID={elementId.GetIntegerValue()}: {action}";
                if (!string.IsNullOrEmpty(details))
                    logEntry += $" - {details}";
                logEntry += "\n";

                // ✅ PERFORMANCE: Safe logging via SafeFileLogger
                SafeFileLogger.SafeAppendText(LogFilePath, logEntry);

                // Also log to main debug logger with special prefix
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL] {elementType} ID={elementId.GetIntegerValue()}: {action} {details}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL-LOG] Failed to write structural log: {ex.Message}");
            }
        }

        /// <summary>
        /// Log structural element sleeve placement success
        /// </summary>
        public static void LogSleevePlace(string elementType, Autodesk.Revit.DB.ElementId elementId, Autodesk.Revit.DB.ElementId sleeveId, string position, string size)
        {
            if (!DebugLogger.IsEnabled) return;
            LogStructuralElement(elementType, elementId, "SLEEVE PLACED", $"Sleeve ID={sleeveId.GetIntegerValue()}, Position={position}, Size={size}");
        }

        /// <summary>
        /// Log structural element sleeve placement failure
        /// </summary>
        public static void LogSleeveFailure(string elementType, Autodesk.Revit.DB.ElementId elementId, string reason)
        {
            if (!DebugLogger.IsEnabled) return;
            LogStructuralElement(elementType, elementId, "SLEEVE FAILED", $"Reason: {reason}");
        }

        /// <summary>
        /// Log structural element intersection detection
        /// </summary>
        public static void LogIntersection(string mepElementType, Autodesk.Revit.DB.ElementId mepElementId, string structuralElementType, Autodesk.Revit.DB.ElementId structuralElementId, string position)
        {
            if (!DebugLogger.IsEnabled) return;
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string logEntry = $"[{timestamp}] {mepElementType}-{structuralElementType.ToUpper()} INTERSECTION - {mepElementType} ID={mepElementId.GetIntegerValue()}, {structuralElementType} ID={structuralElementId.GetIntegerValue()}, Position={position}\n";
            
            try
            {
                if (!IsInitialized)
                    InitializeLogger();
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL] {mepElementType}-{structuralElementType} intersection detected - MEP ID={mepElementId.GetIntegerValue()}, Structural ID={structuralElementId.GetIntegerValue()}");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL] {mepElementType}-{structuralElementType} intersection detected - MEP ID={mepElementId}, Structural ID={structuralElementId}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL-LOG] Failed to write intersection log: {ex.Message}");
            }
        }

        /// <summary>
        /// Log structural type filtering results
        /// </summary>
        public static void LogStructuralTypeCheck(string elementType, Autodesk.Revit.DB.ElementId elementId, bool isStructural, string structuralType = "")
        {
            if (!DebugLogger.IsEnabled) return;
            string action = isStructural ? "STRUCTURAL TYPE CONFIRMED" : "NON-STRUCTURAL TYPE SKIPPED";
            string details = !string.IsNullOrEmpty(structuralType) ? $"Type: {structuralType}" : "";
            LogStructuralElement(elementType, elementId, action, details);
        }

        /// <summary>
        /// Log session summary
        /// </summary>
        public static void LogSummary(string commandName, int totalProcessed, int structuralDetected, int sleevesPlaced, int failures)
        {
            if (!DebugLogger.IsEnabled) return;
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string summary = 
                $"\n[{timestamp}] ===== {commandName.ToUpper()} STRUCTURAL ELEMENTS SUMMARY =====\n" +
                $"Total Elements Processed: {totalProcessed}\n" +
                $"Structural Elements Detected: {structuralDetected}\n" +
                $"Sleeves Successfully Placed: {sleevesPlaced}\n" +
                $"Placement Failures: {failures}\n" +
                $"Success Rate: {(structuralDetected > 0 ? (sleevesPlaced * 100.0 / structuralDetected):0):F1}%\n" +
                $"=========================================\n\n";

            try
            {
                if (!IsInitialized)
                    InitializeLogger();
                // ✅ PERFORMANCE: Safe logging via SafeFileLogger
                SafeFileLogger.SafeAppendText(LogFilePath, summary);
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL] {commandName} summary - Processed: {totalProcessed}, Detected: {structuralDetected}, Placed: {sleevesPlaced}, Failed: {failures}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"[STRUCTURAL-LOG] Failed to write summary: {ex.Message}");
            }
        }

        /// <summary>
        /// Get the current log file path
        /// </summary>
        public static string GetLogFilePath()
        {
            return LogFilePath ?? "Not initialized";
        }

        /// <summary>
        /// Check if logger is properly initialized
        /// </summary>
        public static bool IsLoggerInitialized()
        {
            return IsInitialized && !string.IsNullOrEmpty(LogFilePath) && File.Exists(LogFilePath);
        }
    }
}