using System;
using System.IO;
using System.Reflection;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Safe file logging utility that handles missing directories and prevents crashes.
    /// All file operations are wrapped in try-catch to prevent crashes.
    /// </summary>
    public static class SafeFileLogger
    {
        private static readonly object _lock = new object();
        private static string _logDirectory = null;
        private static bool _logDirectoryInitialized = false;

        /// <summary>
        /// Gets or creates a safe log directory. Uses AppData if project directory doesn't exist.
        /// </summary>
        public static string GetLogDirectory()
        {
            lock (_lock)
            {
                if (_logDirectoryInitialized)
                    return _logDirectory;

                _logDirectory = InitializeLogDirectory();
                _logDirectoryInitialized = true;
                
                // ✅ DIAGNOSTIC: Log the directory location to diagnostic log file
                WriteDiagnosticLogInternal("INIT", $"Log directory initialized: {_logDirectory}, Directory exists: {Directory.Exists(_logDirectory)}", _logDirectory, true);
                
                return _logDirectory;
            }
        }
        
        /// <summary>
        /// Forces initialization of AppData Logs directory (for deployment scenarios)
        /// This ensures the directory exists even if project directory takes priority
        /// </summary>
        public static void EnsureAppDataLogsDirectory()
        {
            try
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings",
                    "Logs"
                );
                
                if (TryCreateDirectory(appDataPath))
                {
                    WriteDiagnosticLogInternal("ENSURE_APPDATA", $"✅ Ensured AppData Logs directory exists: {appDataPath}", appDataPath, true);
                }
                else
                {
                    WriteDiagnosticLogInternal("ENSURE_APPDATA", $"⚠️ Failed to create AppData Logs directory: {appDataPath}", appDataPath, false);
                }
            }
            catch (Exception ex)
            {
                WriteDiagnosticLogInternal("ENSURE_APPDATA", $"Error ensuring AppData Logs directory: {ex.Message}", "unknown", false);
            }
        }

        private static string InitializeLogDirectory()
        {
            // ✅ PRIORITY 1: Use AppData\Roaming for deployment scenarios (preferred for team deployment)
            // This ensures all deployed instances write logs to the same location
            try
            {
                // ✅ CRITICAL FIX: Create base folder first, then Logs subfolder with version tag
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var baseFolder = Path.Combine(appData, "JSE_MEP_Openings");
                
                // Ensure base folder exists first
                if (!Directory.Exists(baseFolder))
                {
                    try
                    {
                        Directory.CreateDirectory(baseFolder);
                        WriteDiagnosticLogInternal("INIT_BASE", $"Created base folder: {baseFolder}", baseFolder, true);
                    }
                    catch (Exception baseEx)
                    {
                        WriteDiagnosticLogInternal("INIT_BASE", $"Cannot create base folder {baseFolder}: {baseEx.Message}", baseFolder, false);
                    }
                }
                
                // ✅ VERSION-SEPARATED LOGS: Append version tag (R2023 or R2024) to separate logs by Revit version
                // This allows testing both versions simultaneously without log conflicts
                string versionTag = VersionInfo.VersionTag; // "R2023" or "R2024"
                string appDataPath = Path.Combine(baseFolder, "Logs", versionTag);

                if (TryCreateDirectory(appDataPath))
                {
                    WriteDiagnosticLogInternal("INIT", $"Using AppData\\Roaming Logs directory: {appDataPath}", appDataPath, true);
                    return appDataPath;
                }
            }
            catch (Exception ex)
            {
                WriteDiagnosticLogInternal("INIT", $"Could not use AppData\\Roaming: {ex.Message}", "unknown", false);
            }

            // ✅ DEPLOYMENT FIX: Priority 2 removed - NEVER use hardcoded project paths
            // All logs now go to AppData (Priority 1) regardless of development vs deployment
            // This ensures deployed users don't see errors about missing C:\JSE_CSharp_Projects paths
            // Development team should check AppData\Roaming\JSE_MEP_Openings\Logs\R2023 for logs

            // Priority 3: Use Temp directory (last resort - always writable) with version tag
            try
            {
                // ✅ VERSION-SEPARATED LOGS: Include version tag in temp path too
                string versionTag = VersionInfo.VersionTag; // "R2023" or "R2024"
                string tempPath = Path.Combine(Path.GetTempPath(), "JSE_MEP_Openings_Logs", versionTag);
                
                if (TryCreateDirectory(tempPath))
                    return tempPath;
            }
            catch
            {
                // If even temp fails, we're in deep trouble
            }

            // Final fallback: Desktop (should always be writable) with version tag
            try
            {
                // ✅ VERSION-SEPARATED LOGS: Include version tag in desktop path too
                string versionTag = VersionInfo.VersionTag; // "R2023" or "R2024"
                string desktopPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "JSE_MEP_Openings_Logs",
                    versionTag
                );
                
                if (TryCreateDirectory(desktopPath))
                    return desktopPath;
            }
            catch
            {
                // Ignore
            }

            // If everything fails, return temp (should never happen)
            return Path.GetTempPath();
        }

        private static bool TryCreateDirectory(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path))
                    return false;

                // Directory.CreateDirectory will create all parent directories if they don't exist
                // This ensures JSE_MEP_Openings is created if it doesn't exist
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    WriteDiagnosticLogInternal("CREATE_DIR", $"Created directory: {path}", path, true);
                }

                // Verify we can write to it
                string testFile = Path.Combine(path, "write_test.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                
                return true;
            }
            catch (Exception ex)
            {
                WriteDiagnosticLogInternal("CREATE_DIR", $"Failed to create directory {path}: {ex.Message}", path, false);
                return false;
            }
        }

        /// <summary>
        /// Internal helper to write diagnostic logs without circular dependency
        /// Uses direct path calculation to avoid calling GetLogDirectory() which might trigger initialization
        /// </summary>
        private static void WriteDiagnosticLogInternal(string operation, string message, string path, bool success)
        {
            string diagnosticLogPath = null;
            Exception lastException = null;
            
            try
            {
                // ✅ FIX: Use direct path calculation to avoid circular dependency with GetLogDirectory()
                // This ensures diagnostic logs can be written even during initialization
                
                // If log directory is already initialized, use it
                if (_logDirectoryInitialized && !string.IsNullOrEmpty(_logDirectory))
                {
                    diagnosticLogPath = Path.Combine(_logDirectory, "safefilelogger_diagnostic.log");
                }
                else
                {
                    // During initialization, use AppData path directly (same logic as InitializeLogDirectory)
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string baseFolder = Path.Combine(appData, "JSE_MEP_Openings");
                    diagnosticLogPath = Path.Combine(baseFolder, "Logs", "safefilelogger_diagnostic.log");
                    
                    // Ensure directory exists
                    string diagnosticDir = Path.GetDirectoryName(diagnosticLogPath);
                    if (!string.IsNullOrEmpty(diagnosticDir) && !Directory.Exists(diagnosticDir))
                    {
                        try
                        {
                            Directory.CreateDirectory(diagnosticDir);
                        }
                        catch (Exception dirEx)
                        {
                            lastException = dirEx;
                            // Try fallback to Desktop
                            string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "JSE_MEP_Openings_Diagnostic.log");
                            try
                            {
                                File.AppendAllText(desktopPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] Directory creation failed: {dirEx.Message}, Operation={operation}\n");
                            }
                            catch { }
                        }
                    }
                }
                
                if (diagnosticLogPath != null)
                {
                    string diagnosticEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] Operation={operation}, Message={message}, Path={path}, Success={success}\n";
                    
                    // Use direct file write for diagnostic log (only exception - needed to debug logging issues)
                    File.AppendAllText(diagnosticLogPath, diagnosticEntry);
                }
            }
            catch (Exception ex)
            {
                lastException = ex;
                // ✅ FIX: Try fallback to Desktop if AppData fails
                try
                {
                    string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "JSE_MEP_Openings_Diagnostic.log");
                    string fallbackEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] ERROR: Failed to write diagnostic log. Operation={operation}, Exception={ex.Message}, Type={ex.GetType().Name}, StackTrace={ex.StackTrace}\n";
                    File.AppendAllText(desktopPath, fallbackEntry);
                }
                catch
                {
                    // If even Desktop write fails, give up silently
                }
            }
        }

        /// <summary>
        /// Writes diagnostic information to a diagnostic log file (always writes, even if main logging fails)
        /// Used for troubleshooting why logs aren't appearing
        /// </summary>
        private static void WriteDiagnosticLog(string fileName, string message, string logPath, bool success)
        {
            string diagnosticLogPath = null;
            Exception lastException = null;
            
            try
            {
                // ✅ FIX: Use direct path calculation to avoid circular dependency with GetLogDirectory()
                
                // If log directory is already initialized, use it
                if (_logDirectoryInitialized && !string.IsNullOrEmpty(_logDirectory))
                {
                    diagnosticLogPath = Path.Combine(_logDirectory, "safefilelogger_diagnostic.log");
                }
                else
                {
                    // During initialization, use AppData path directly
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string baseFolder = Path.Combine(appData, "JSE_MEP_Openings");
                    diagnosticLogPath = Path.Combine(baseFolder, "Logs", "safefilelogger_diagnostic.log");
                    
                    // Ensure directory exists
                    string diagnosticDir = Path.GetDirectoryName(diagnosticLogPath);
                    if (!string.IsNullOrEmpty(diagnosticDir) && !Directory.Exists(diagnosticDir))
                    {
                        try
                        {
                            Directory.CreateDirectory(diagnosticDir);
                        }
                        catch (Exception dirEx)
                        {
                            lastException = dirEx;
                            // Try fallback to Desktop
                            string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "JSE_MEP_Openings_Diagnostic.log");
                            try
                            {
                                File.AppendAllText(desktopPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] Directory creation failed: {dirEx.Message}, fileName={fileName}\n");
                            }
                            catch { }
                        }
                    }
                }
                
                if (diagnosticLogPath != null)
                {
                    string diagnosticEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] fileName={fileName}, messageLength={message?.Length ?? 0}, DeploymentMode={DeploymentConfiguration.DeploymentMode}, LogPath={logPath}, Success={success}, DirectoryExists={Directory.Exists(Path.GetDirectoryName(logPath))}\n";
                    
                    // Use direct file write for diagnostic log (only exception - needed to debug logging issues)
                    File.AppendAllText(diagnosticLogPath, diagnosticEntry);
                }
            }
            catch (Exception ex)
            {
                lastException = ex;
                // ✅ FIX: Try fallback to Desktop if AppData fails
                try
                {
                    string desktopPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "JSE_MEP_Openings_Diagnostic.log");
                    string fallbackEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SAFEFILELOGGER] ERROR: Failed to write diagnostic log. fileName={fileName}, Exception={ex.Message}, Type={ex.GetType().Name}, StackTrace={ex.StackTrace}\n";
                    File.AppendAllText(desktopPath, fallbackEntry);
                }
                catch
                {
                    // If even Desktop write fails, give up silently
                }
            }
        }

        /// <summary>
        /// Safely append text to a log file. Never throws exceptions.
        /// </summary>
        public static void SafeAppendText(string fileName, string message)
        {
            // ✅ DEPLOYMENT MODE: Disable ALL logging in deployment mode
            if (DeploymentConfiguration.DeploymentMode)
            {
                return; // Skip all file writes in deployment mode
            }

            if (ShouldSuppress(fileName))
                return;

            SafeAppendTextAlways(fileName, message);
        }

        /// <summary>
        /// Global suppression logic for verbose diagnostic logs.
        /// </summary>
        private static bool ShouldSuppress(string fileName)
        {
            // ✅ GLOBAL VERBOSE LOGGING CHECK
            // If DisableVerboseLogging is ON, suppress ALL logs except critical ones (like placement_performance.log)
            if (OptimizationFlags.DisableVerboseLogging)
            {
                // Allow specific critical logs even when verbose logging is disabled
                if (string.Equals(fileName, "placement_performance.log", StringComparison.OrdinalIgnoreCase))
                    return false;
                
                // ✅ PERFORMANCE REPORTS: Allow performance logs for all operations (low volume, critical for diagnostics)
                if (fileName != null && fileName.StartsWith("performance_", StringComparison.OrdinalIgnoreCase))
                    return false;

                // ✅ REFRESH LOGS: Always allow main refresh logs (critical for troubleshooting)
                if (fileName != null && fileName.StartsWith("Refresh_", StringComparison.OrdinalIgnoreCase))
                    return false;

                // Allow error logs
                if (string.Equals(fileName, "safefilelogger_errors.log", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(fileName, "safefilelogger_diagnostic.log", StringComparison.OrdinalIgnoreCase))
                    return false;

                // Suppress everything else
                return true;
            }

            if (string.Equals(fileName, "placement_debug.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnablePlacementDebugLog)
                return true;

            if (string.Equals(fileName, "bulk_placement_debug.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableBulkPlacementDebugLog)
                return true;

            if (string.Equals(fileName, "flag_workflow.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableProximityDebugLog)
                return true;

            if (string.Equals(fileName, "batch_v2.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableBatchClusteringDebugLog)
                return true;
            
            if (string.Equals(fileName, "proximity_debug.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableProximityDebugLog)
                return true;

            if (string.Equals(fileName, "cluster_debug.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableClusterSizingDebugLog)
                return true;

            if (string.Equals(fileName, "placement_sizing_debug.log", StringComparison.OrdinalIgnoreCase)
                && !OptimizationFlags.EnableClusterSizingDebugLog)
                return true;

            return false;
        }

        public static void SafeAppendTextAlways(string fileName, string message)
        {
            // ✅ DEPLOYMENT: In production, skip most logging — but always write placement performance log so you can see timing
            if (DeploymentConfiguration.DeploymentMode &&
                !string.Equals(fileName, "placement_performance.log", StringComparison.OrdinalIgnoreCase))
                return;
            
            if (ShouldSuppress(fileName))
                return;

            string logPathFinal = null;
            // bool writeSuccess = false; // FIX: CS0219 - commented to fix critical warning
                
            try
            {
                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(message))
                {
                    WriteDiagnosticLog(fileName ?? "null", message ?? "null", "unknown", false); // Log skipped
                    return;
                }

                string logDir = GetLogDirectory();
                logPathFinal = Path.Combine(logDir, fileName);

                // Ensure directory exists (should already exist, but double-check)
                // Directory.CreateDirectory will create all parent directories if they don't exist
                string directory = Path.GetDirectoryName(logPathFinal);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    try
                    {
                        Directory.CreateDirectory(directory); // Creates parent directories too
                        WriteDiagnosticLog(fileName, $"Created log directory: {directory}", logPathFinal, true);
                    }
                    catch (Exception dirEx)
                    {
                        WriteDiagnosticLog(fileName, $"Failed to create directory {directory}: {dirEx.Message}", logPathFinal, false); // Log failure
                        // Continue - will be caught by outer exception handler
                    }
                }

                // Append with timestamp
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n";
                
                lock (_lock)
                {
                    File.AppendAllText(logPathFinal, logEntry);
                    // writeSuccess = true; // FIX: CS0219 - commented to fix critical warning
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Silently fail - can't write to log (user doesn't have permission)
                WriteDiagnosticLog(fileName, $"No permission to write to log: {fileName}", logPathFinal ?? "unknown", false); // Log permission error
            }
            catch (DirectoryNotFoundException)
            {
                // Directory was deleted or doesn't exist - try to recreate
                try
                {
                    _logDirectoryInitialized = false; // Force re-initialization
                    string logDir = GetLogDirectory();
                    logPathFinal = Path.Combine(logDir, fileName);
                    string directory = Path.GetDirectoryName(logPathFinal);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        // Directory.CreateDirectory creates all parent directories if they don't exist
                        Directory.CreateDirectory(directory);
                        WriteDiagnosticLog(fileName, $"Recreated directory: {directory}", logPathFinal, true);
                        File.AppendAllText(logPathFinal, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
                        // writeSuccess = true; // FIX: CS0219 - commented to fix critical warning
                        WriteDiagnosticLog(fileName, $"SUCCESS after retry: Written to {logPathFinal}", logPathFinal, true); // Log success after retry
                    }
                }
                catch (Exception retryEx)
                {
                    // Final failure - log to diagnostic file
                    WriteDiagnosticLog(fileName, $"Failed to recreate directory and write log: {retryEx.Message}", logPathFinal ?? "unknown", false); // Log failure
                }
            }
            catch (Exception ex)
            {
                // Any other error - log to diagnostic file (don't crash!)
                // ✅ INVESTIGATION: Enhanced error logging to diagnose why logs aren't appearing
                string errorDetails = $"Error writing to {fileName}: {ex.Message}, Exception Type: {ex.GetType().Name}, Stack Trace: {ex.StackTrace}, DeploymentMode: {DeploymentConfiguration.DeploymentMode}, Log Directory: {GetLogDirectory()}, Log Path: {Path.Combine(GetLogDirectory(), fileName)}";
                
                WriteDiagnosticLog(fileName, $"{message} [ERROR: {errorDetails}]", logPathFinal ?? "unknown", false); // Log error
                
                // ✅ INVESTIGATION: Also try to write error to a separate error log file (using direct write as last resort)
                try
                {
                    string errorLogPath = Path.Combine(GetLogDirectory(), "safefilelogger_errors.log");
                    File.AppendAllText(errorLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {errorDetails}\n\n");
                }
                catch
                {
                    // If even error logging fails, just give up
                }
            }
        }

        /// <summary>
        /// Gets the full path to a log file (for reference/debugging).
        /// </summary>
        public static string GetLogFilePath(string fileName)
        {
            try
            {
                return Path.Combine(GetLogDirectory(), fileName);
            }
            catch
            {
                return Path.Combine(Path.GetTempPath(), fileName);
            }
        }

        /// <summary>
        /// Safely reads all text from a log file. Returns empty string if file doesn't exist or error occurs.
        /// </summary>
        public static string SafeReadAllText(string fileName)
        {
            try
            {
                string logPath = GetLogFilePath(fileName);
                if (File.Exists(logPath))
                {
                    return File.ReadAllText(logPath);
                }
            }
            catch
            {
                // Silently return empty string
            }
            return string.Empty;
        }

        /// <summary>
        /// Safely checks if a log file exists.
        /// </summary>
        public static bool LogFileExists(string fileName)
        {
            try
            {
                return File.Exists(GetLogFilePath(fileName));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets information about the log directory (for troubleshooting).
        /// </summary>
        public static string GetLogDirectoryInfo()
        {
            try
            {
                string logDir = GetLogDirectory();
                bool exists = Directory.Exists(logDir);
                bool writable = false;
                
                try
                {
                    string testFile = Path.Combine(logDir, "test.tmp");
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);
                    writable = true;
                }
                catch { }

                // Also check AppData directory status
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings",
                    "Logs"
                );
                bool appDataExists = Directory.Exists(appDataPath);
                
                return $"Log Directory: {logDir}\nExists: {exists}\nWritable: {writable}\n\n" +
                       $"AppData Logs Directory: {appDataPath}\nAppData Exists: {appDataExists}";
            }
            catch (Exception ex)
            {
                return $"Log Directory Info Error: {ex.Message}";
            }
        }
    }
}

