using System;
using System.IO;
using System.Reflection;  // for build timestamp
using System.Runtime.CompilerServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Static logger class for debug messages with support for log levels and source tracking
    /// </summary>
    public static class DebugLogger
    {
    /// <summary>
    /// Set to false to disable all logging globally.
    /// Default enabled here to allow runtime diagnostics for cable tray placement.
    /// Toggle to false if you want to silence logs.
    /// NOTE: DeploymentConfiguration.DeploymentMode automatically disables all logging.
    /// </summary>
    public static bool IsEnabled =false;
    
    /// <summary>
    /// Current service name for logging context
    /// </summary>
    public static string CurrentService = "MainUI";
    // ✅ FIXED: Use SafeFileLogger for deployment-compatible paths (no hardcoded paths)
    private static string LogDir => SafeFileLogger.GetLogDirectory();
    private static string DuctLogFilePath = Path.Combine(LogDir, "ductsleeveplacer.log");
    private static string CableTrayLogFilePath = Path.Combine(LogDir, "cabletraysleeveplacer.log");
    private static string DamperLogFilePath = Path.Combine(LogDir, "dampersleeveplacer.log");
    private static string CombinedSleeveLogFilePath = Path.Combine(LogDir, "combinesleeveplacer.log");
    // ✅ REMOVED: MainUiLogFilePath - Main UI logging removed per user request
    private static string LogFilePath = CableTrayLogFilePath; // Default

    // Single shared writer to avoid repeated open/close per log entry
    private static readonly object _writerLock = new object();
    private static StreamWriter? _writer = null;
    
    // Cache assembly/version info to avoid repeated reflection calls during logging
    private static readonly string _cachedVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
    private static readonly string _cachedAssemblyPath = Assembly.GetExecutingAssembly().Location;

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    /// <summary>
    /// Maximum log file size before rotation (10MB)
    /// </summary>
    private const long MaxLogFileSize = 10 * 1024 * 1024;

    /// <summary>
    /// Check if logging is enabled for the current service
    /// </summary>
    private static bool IsLoggingEnabledForCurrentService()
    {
        // 🔴 FORCE ENABLE for Combined Sleeve Debugging
        if (CurrentService == "CombinedSleeveManual" || CurrentService == "CombinedSleeveAuto")
            return true;
            
        // ⚡ OPTIMIZATION: Disable verbose logging when flag is set (keeps performance logs only)
        if (OptimizationFlags.DisableVerboseLogging)
            return false;
            
        // ✅ DEPLOYMENT MODE: Disable all logging if deployment mode is enabled
        if (DeploymentConfiguration.DeploymentMode)
            return false;
            
        if (!IsEnabled) return false;
        return LoggingConfiguration.IsLoggingEnabled(CurrentService);
    }

    /// <summary>
    /// Check if log file needs rotation and rotate if necessary
    /// </summary>
    private static void CheckAndRotateLogFile()
    {
        try
        {
            if (string.IsNullOrEmpty(LogFilePath) || !File.Exists(LogFilePath))
                return;

            var fileInfo = new FileInfo(LogFilePath);
            if (fileInfo.Length > MaxLogFileSize)
            {
                // Close current writer
                lock (_writerLock)
                {
                    try
                    {
                        if (_writer != null)
                        {
                            _writer.Flush();
                            _writer.Close();
                            _writer.Dispose();
                        }
                    }
                    catch { }
                    _writer = null;
                }

                // Rotate the file
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string rotatedPath = Path.Combine(
                    Path.GetDirectoryName(LogFilePath) ?? LogDir,
                    $"{Path.GetFileNameWithoutExtension(LogFilePath)}_rotated_{timestamp}.log"
                );

                try
                {
                    File.Move(LogFilePath, rotatedPath);
                }
                catch (Exception ex)
                {
                    // If rotation fails, try to delete the old file
                    try { File.Delete(LogFilePath); } catch { }
                }

                // Create new log file
                var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
                string header =
                    $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                    $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {Path.GetFileName(LogFilePath)}\n" +
                    $"Build Version: {_cachedVersion}\n" +
                    $"Build Timestamp: {buildTimestamp}\n" +
                    $"Wrote: {_cachedAssemblyPath}\n" +
                    $"Previous log rotated to: {rotatedPath}\n" +
                    $"====================================================\n";

                EnsureWriterInitialized(LogFilePath, header);
            }
        }
        catch
        {
            // Log rotation should not break logging
        }
    }
    
    /// <summary>
    /// Close all file handles and stop logging
    /// </summary>
    public static void CloseAllLogFiles()
    {
        lock (_writerLock)
        {
            try 
            { 
                if (_writer != null) 
                { 
                    _writer.Flush(); 
                    _writer.Close(); 
                    _writer.Dispose(); 
                } 
            } 
            catch { }
            _writer = null;
        }
    }
        
    /// <summary>
    /// Set the current service context for logging
    /// </summary>
    public static void SetServiceContext(string serviceName)
    {
        CurrentService = serviceName;
    }
    
    /// <summary>
    /// Start a new log file at application startup (default log name)
    /// </summary>
    public static void InitLogFile()
    {
        if (!IsLoggingEnabledForCurrentService()) return;
        InitLogFile("cabletraysleeveplacer");
    }

    /// <summary>
    /// Start a new log file with a custom file name (without extension)
    /// ✅ FIXED: Now OVERWRITES existing log file instead of appending
    /// </summary>
    public static void InitLogFile(string logFileName)
    {
        if (!IsLoggingEnabledForCurrentService()) return;
        try
        {
            // Use the hard-coded log directory and ensure it exists
            string logDir = LogDir;
            if (!Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);
            // If logFileName has an extension, use as is; otherwise, add .log
            LogFilePath = Path.Combine(logDir, logFileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase) ? logFileName : logFileName + ".log");
            // Include build/version information
            var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
            string header =
                $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                $"Build Version: {_cachedVersion}\n" +
                $"Build Timestamp: {buildTimestamp}\n" +
                $"Wrote: {_cachedAssemblyPath}\n" +
                $"====================================================\n";
            // ✅ FIXED: Initialize writer with overwrite=true to clear old logs
            EnsureWriterInitialized(LogFilePath, header, overwrite: true);
        }
        catch (Exception ex)
        {
            // Log to default log if custom log creation fails - ✅ FIXED: Use SafeFileLogger for deployment-compatible path
            string fallbackLog = SafeFileLogger.GetLogFilePath("cabletraysleeveplacer.log");
            string msg = $"[LOGGER ERROR] Could not create custom log file '{logFileName}': {ex.Message}\n{ex.StackTrace}\n";
            try { File.AppendAllText(fallbackLog, msg); } catch { /* ignore */ }
        }
    }

    /// <summary>
    /// Start a new log file with a custom file name and build timestamp
    /// </summary>
    public static void InitCustomLogFile(string logFileName)
    {
        if (!IsLoggingEnabledForCurrentService()) return;
        try
        {
            // Set log file path with build timestamp under the hard-coded log directory
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            if (!Directory.Exists(LogDir)) Directory.CreateDirectory(LogDir);
            LogFilePath = Path.Combine(LogDir, $"{logFileName}_{timestamp}.log");

            var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
            string header =
                $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                $"Build Version: {_cachedVersion}\n" +
                $"Build Timestamp: {buildTimestamp}\n" +
                $"Wrote: {_cachedAssemblyPath}\n" +
                $"====================================================\n";
            EnsureWriterInitialized(LogFilePath, header);
        }
        catch
        {
            // Silently fail - we don't want logging to break the application
        }
    }

    /// <summary>
    /// Start a new log file with a custom file name and overwrite existing content
    /// </summary>
    public static void InitCustomLogFileOverwrite(string logFileName)
    {
        // ✅ DEPLOYMENT MODE: Check is handled by IsLoggingEnabledForCurrentService() - no duplicate check needed
        if (!IsLoggingEnabledForCurrentService())
            return;
        
        // ✅ FIX: Use SafeFileLogger instead of direct File.AppendAllText() to follow coding standards
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        SafeFileLogger.SafeAppendText("logger_debug.txt", $"[{DateTime.Now}] InitCustomLogFileOverwrite called with: {logFileName} (timestamp: {timestamp})\n");

        if (!IsEnabled)
        {
            SafeFileLogger.SafeAppendText("logger_debug.txt", $"[{DateTime.Now}] DebugLogger.IsEnabled = false\n");
            return;
        }

        try
        {
            // Use timestamped log files to avoid overwriting
            string logDir = SafeFileLogger.GetLogDirectory();
            LogFilePath = Path.Combine(logDir, $"{logFileName}_{timestamp}.log");

            // Ensure directory exists
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // Include build/version information
            var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
            string header =
                $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                $"Build Version: {_cachedVersion}\n" +
                $"Build Timestamp: {buildTimestamp}\n" +
                $"Wrote: {_cachedAssemblyPath}\n" +
                $"Log Path: {LogFilePath}\n" +
                $"====================================================\n";

            EnsureWriterInitialized(LogFilePath, header, overwrite: true);

            // Test log to verify the custom file was created
            Info($"Custom log file initialized: {LogFilePath}");
        }
        catch (Exception ex)
        {
            // ✅ FIX: Use SafeFileLogger instead of direct File.AppendAllText() to follow coding standards
            SafeFileLogger.SafeAppendText("logger_debug.txt", $"[{DateTime.Now}] ERROR in InitCustomLogFileOverwrite: {ex.Message}\n{ex.StackTrace}\n");

            // Log to fallback log file with timestamp
            try
            {
                string fallbackTimestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string fallbackLogDir = SafeFileLogger.GetLogDirectory();
                string fallbackLog = Path.Combine(fallbackLogDir, $"fallback_debug_{fallbackTimestamp}.log");
                SafeFileLogger.SafeAppendText($"fallback_debug_{fallbackTimestamp}.log", $"[{DateTime.Now}] ERROR initializing custom log '{logFileName}': {ex.Message}\n{ex.StackTrace}\n");
            }
            catch (Exception fallbackEx)
            {
                SafeFileLogger.SafeAppendText("logger_debug.txt", $"[{DateTime.Now}] Fallback logging also failed: {fallbackEx.Message}\n");
            }
        }
    }

    /// <summary>
    /// Initialize log file using an absolute path (full filename). Creates directory if needed.
    /// ✅ FIXED: Now OVERWRITES existing log file instead of appending
    /// </summary>
    public static void InitAbsoluteLogFile(string absoluteFilePath)
    {
        if (!IsLoggingEnabledForCurrentService()) return;
        try
        {
            var logDir = Path.GetDirectoryName(absoluteFilePath);
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
                Directory.CreateDirectory(logDir);

            LogFilePath = absoluteFilePath;
            var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
            string header =
                $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {Path.GetFileName(absoluteFilePath)}\n" +
                $"Build Version: {_cachedVersion}\n" +
                $"Build Timestamp: {buildTimestamp}\n" +
                $"Wrote: {_cachedAssemblyPath}\n" +
                $"====================================================\n";
            // ✅ FIXED: Initialize writer with overwrite=true to clear old logs
            EnsureWriterInitialized(LogFilePath, header, overwrite: true);
        }
        catch
        {
            // Silently fail - we don't want logging to break the application
        }
    }

    private static void EnsureWriterInitialized(string path, string header, bool overwrite = false)
    {
        try
        {
            lock (_writerLock)
            {
                if (_writer != null)
                {
                    // If already pointing to same file, nothing to do
                    if (string.Equals(_writer?.BaseStream is FileStream fs ? fs.Name : null, path, StringComparison.OrdinalIgnoreCase))
                        return;
                    // Close existing writer
                    try { if (_writer != null) { _writer.Flush(); _writer.Close(); _writer.Dispose(); } } catch { }
                    _writer = null;
                }

                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

                var fileMode = overwrite ? FileMode.Create : FileMode.Append;
                var fsNew = new FileStream(path, fileMode, FileAccess.Write, FileShare.ReadWrite);
                _writer = new StreamWriter(fsNew) { AutoFlush = true };
                if (!overwrite)
                    _writer.Write(header);
                else
                    _writer.Write(header);
            }
        }
        catch
        {
            // swallow - logging must not throw
        }
    }

        public static void SetDuctLogFile()
        {
            LogFilePath = DuctLogFilePath;
        }
        public static void SetCableTrayLogFile()
        {
            LogFilePath = CableTrayLogFilePath;
        }
        public static void SetDamperLogFile()
        {
            LogFilePath = DamperLogFilePath;
        }
        public static void SetCombinedSleeveLogFile()
        {
            LogFilePath = CombinedSleeveLogFilePath;
        }

        /// <summary>
        /// Standard log method for backward compatibility
        /// </summary>
        public static void Log(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsLoggingEnabledForCurrentService()) return;
            Log(LogLevel.Info, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Enhanced log method with log level and source tracking
        /// </summary>
        public static void Log(
            LogLevel level,
            string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsLoggingEnabledForCurrentService()) return;
            try
            {
                // Check if log rotation is needed before writing
                CheckAndRotateLogFile();

                // Get the class name from the source file path
                string className = Path.GetFileNameWithoutExtension(sourceFile);
                // Format the log entry with timestamp, version, level, class and line
                string levelText = level.ToString().ToUpper();
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [v{_cachedVersion}] [{levelText}] [{className}:{lineNumber}] {message}{Environment.NewLine}";
                lock (_writerLock)
                {
                    if (_writer == null)
                    {
                        // Try initialize minimal writer
                        try { EnsureWriterInitialized(LogFilePath, ""); } catch { }
                    }
                    try { _writer?.Write(logEntry); } catch { }
                }
            }
            catch
            {
                // Logging failures should not break execution
            }
        }

        /// <summary>
        /// Log a warning message
        /// </summary>
        public static void Warning(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsLoggingEnabledForCurrentService()) return;
            Log(LogLevel.Warning, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Log an error message
        /// </summary>
        public static void Error(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsLoggingEnabledForCurrentService()) return;
            Log(LogLevel.Error, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Log a debug message
        /// </summary>
        public static void Debug(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            Log(LogLevel.Debug, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Log a critical error message
        /// </summary>
        public static void Critical(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            Log(LogLevel.Error, $"CRITICAL: {message}", sourceFile, lineNumber);
        }

        /// <summary>
        /// Log an informational message
        /// </summary>
        public static void Info(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            Log(LogLevel.Info, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Check if the log file is writable
        /// </summary>
        public static bool IsLogFileWritable()
        {
            try
            {
                using (FileStream fs = new FileStream(LogFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    return fs.CanWrite;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Clean up old log files (keep only last 10 files per type)
        /// </summary>
        public static void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(LogDir))
                    return;

                var logFiles = Directory.GetFiles(LogDir, "*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                // Keep only the 10 most recent files
                var filesToDelete = logFiles.Skip(10).ToList();

                foreach (var file in filesToDelete)
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        // Log deletion failure but continue
                        System.Diagnostics.Debug.WriteLine($"Failed to delete old log file {file.Name}: {ex.Message}");
                    }
                }

                if (filesToDelete.Any())
                {
                    System.Diagnostics.Debug.WriteLine($"Cleaned up {filesToDelete.Count} old log files");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during log cleanup: {ex.Message}");
            }
        }
    }
}
