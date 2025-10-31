using System;
using System.IO;
using System.Reflection;

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
                
                // Log the directory location to debug output for troubleshooting
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Log directory initialized: {_logDirectory}");
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Directory exists: {Directory.Exists(_logDirectory)}");
                
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
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] ✅ Ensured AppData Logs directory exists: {appDataPath}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] ⚠️ Failed to create AppData Logs directory: {appDataPath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Error ensuring AppData Logs directory: {ex.Message}");
            }
        }

        private static string InitializeLogDirectory()
        {
            // ✅ PRIORITY 1: Use AppData\Roaming for deployment scenarios (preferred for team deployment)
            // This ensures all deployed instances write logs to the same location
            try
            {
                // ✅ CRITICAL FIX: Create base folder first, then Logs subfolder
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var baseFolder = Path.Combine(appData, "JSE_MEP_Openings");
                
                // Ensure base folder exists first
                if (!Directory.Exists(baseFolder))
                {
                    try
                    {
                        Directory.CreateDirectory(baseFolder);
                        System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Created base folder: {baseFolder}");
                    }
                    catch (Exception baseEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Cannot create base folder {baseFolder}: {baseEx.Message}");
                    }
                }
                
                string appDataPath = Path.Combine(baseFolder, "Logs");

                if (TryCreateDirectory(appDataPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Using AppData\\Roaming Logs directory: {appDataPath}");
                    return appDataPath;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Could not use AppData\\Roaming: {ex.Message}");
            }

            // Priority 2: Try to use project directory (for development only)
            // Only use this if assembly is actually in the project directory
            try
            {
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyLocation);
                
                // Check if assembly is in project directory (development scenario)
                bool isDevelopmentPath = assemblyLocation.Contains(@"JSE_CSharp_Projects\JSE_MEPOPENING_23") ||
                                         assemblyLocation.Contains(@"JSE_RevitAddin_MEP_OPENINGS");
                
                if (isDevelopmentPath && !string.IsNullOrEmpty(assemblyDir))
                {
                    // Navigate up to find project root
                    string projectRoot = assemblyDir;
                    for (int i = 0; i < 5 && projectRoot != null; i++)
                    {
                        // ✅ FIX: Check for "Logs" (plural) not "Log" (singular)
                        if (Directory.Exists(Path.Combine(projectRoot, "Logs")))
                        {
                            string logDir = Path.Combine(projectRoot, "Logs");
                            System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Using project Logs directory (development): {logDir}");
                            return logDir;
                        }
                        // Also check for "Log" (singular) for backward compatibility
                        if (Directory.Exists(Path.Combine(projectRoot, "Log")))
                        {
                            string logDir = Path.Combine(projectRoot, "Log");
                            System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Using project Log directory (development - fallback): {logDir}");
                            return logDir;
                        }
                        projectRoot = Directory.GetParent(projectRoot)?.FullName;
                    }
                    
                    // If project structure found, create Logs directory
                    if (assemblyDir != null)
                    {
                        string logDir = Path.Combine(assemblyDir, "..", "..", "..", "Logs");
                        logDir = Path.GetFullPath(logDir); // Resolve .. paths
                        
                        if (TryCreateDirectory(logDir))
                        {
                            System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Using project Log directory (development): {logDir}");
                            return logDir;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently continue to fallback options
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Could not use project directory: {ex.Message}");
            }

            // Priority 3: Use Temp directory (last resort - always writable)
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), "JSE_MEP_Openings_Logs");
                
                if (TryCreateDirectory(tempPath))
                    return tempPath;
            }
            catch
            {
                // If even temp fails, we're in deep trouble
            }

            // Final fallback: Desktop (should always be writable)
            try
            {
                string desktopPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "JSE_MEP_Openings_Logs"
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
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Created directory: {path}");
                }

                // Verify we can write to it
                string testFile = Path.Combine(path, "write_test.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Failed to create directory {path}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Safely append text to a log file. Never throws exceptions.
        /// </summary>
        public static void SafeAppendText(string fileName, string message)
        {
            // ✅ DEPLOYMENT MODE: Allow only specific logs
            // Allowed: performance.log, Refresh_*.log, refresh_memory_profiling_*.log
            if (DeploymentConfiguration.DeploymentMode)
            {
                bool isPerformance = string.Equals(fileName, "performance.log", StringComparison.OrdinalIgnoreCase);
                bool isRefresh = fileName.StartsWith("Refresh_", StringComparison.OrdinalIgnoreCase) && fileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase);
                bool isRefreshMemory = fileName.StartsWith("refresh_memory_profiling_", StringComparison.OrdinalIgnoreCase) && fileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase);
                bool isSleevePlacement = fileName.StartsWith("sleeve_placement_", StringComparison.OrdinalIgnoreCase) && fileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase);
                if (!isPerformance && !isRefresh && !isRefreshMemory && !isSleevePlacement)
                    return;
            }
                
            try
            {
                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(message))
                    return;

                string logDir = GetLogDirectory();
                string logPath = Path.Combine(logDir, fileName);

                // Ensure directory exists (should already exist, but double-check)
                // Directory.CreateDirectory will create all parent directories if they don't exist
                string directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    try
                    {
                        Directory.CreateDirectory(directory); // Creates parent directories too
                        System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Created log directory: {directory}");
                    }
                    catch (Exception dirEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Failed to create directory {directory}: {dirEx.Message}");
                        // Continue - will be caught by outer exception handler
                    }
                }

                // Append with timestamp
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n";
                
                lock (_lock)
                {
                    File.AppendAllText(logPath, logEntry);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Silently fail - can't write to log (user doesn't have permission)
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] No permission to write to log: {fileName}");
            }
            catch (DirectoryNotFoundException)
            {
                // Directory was deleted or doesn't exist - try to recreate
                try
                {
                    _logDirectoryInitialized = false; // Force re-initialization
                    string logDir = GetLogDirectory();
                    string logPath = Path.Combine(logDir, fileName);
                    string directory = Path.GetDirectoryName(logPath);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        // Directory.CreateDirectory creates all parent directories if they don't exist
                        Directory.CreateDirectory(directory);
                        System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Recreated directory: {directory}");
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
                    }
                }
                catch (Exception retryEx)
                {
                    // Final failure - log to debug only
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Failed to recreate directory and write log {fileName}: {retryEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Any other error - log to debug only (don't crash!)
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Error writing to {fileName}: {ex.Message}");
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

