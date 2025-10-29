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
                return _logDirectory;
            }
        }

        private static string InitializeLogDirectory()
        {
            // Priority 1: Try to use project directory (for development)
            try
            {
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyLocation);
                
                if (!string.IsNullOrEmpty(assemblyDir))
                {
                    // Navigate up to find project root
                    string projectRoot = assemblyDir;
                    for (int i = 0; i < 5 && projectRoot != null; i++)
                    {
                        if (Directory.Exists(Path.Combine(projectRoot, "Log")))
                        {
                            return Path.Combine(projectRoot, "Log");
                        }
                        projectRoot = Directory.GetParent(projectRoot)?.FullName;
                    }
                    
                    // If project structure found, create Log directory
                    if (assemblyDir != null)
                    {
                        string logDir = Path.Combine(assemblyDir, "..", "..", "..", "Log");
                        logDir = Path.GetFullPath(logDir); // Resolve .. paths
                        
                        if (TryCreateDirectory(logDir))
                            return logDir;
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently continue to fallback options
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Could not use project directory: {ex.Message}");
            }

            // Priority 2: Use AppData (always available on all Windows systems)
            try
            {
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings",
                    "Logs"
                );

                if (TryCreateDirectory(appDataPath))
                    return appDataPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Could not use AppData: {ex.Message}");
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

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                // Verify we can write to it
                string testFile = Path.Combine(path, "write_test.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Safely append text to a log file. Never throws exceptions.
        /// </summary>
        public static void SafeAppendText(string fileName, string message)
        {
            try
            {
                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(message))
                    return;

                string logDir = GetLogDirectory();
                string logPath = Path.Combine(logDir, fileName);

                // Ensure directory exists (should already exist, but double-check)
                string directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
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
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
                    }
                }
                catch
                {
                    // Final failure - log to debug only
                    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Failed to write log: {fileName}");
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

                return $"Log Directory: {logDir}\nExists: {exists}\nWritable: {writable}";
            }
            catch (Exception ex)
            {
                return $"Log Directory Info Error: {ex.Message}";
            }
        }
    }
}

