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
    /// </summary>
    public static bool IsEnabled = false;
    // Always log to the hard-coded Log directory requested by the user
    private static readonly string LogDir = "C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23\\Log";
    private static string DuctLogFilePath = Path.Combine(LogDir, "ductsleeveplacer.log");
    private static string CableTrayLogFilePath = Path.Combine(LogDir, "cabletraysleeveplacer.log");
    private static string DamperLogFilePath = Path.Combine(LogDir, "dampersleeveplacer.log");
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
        /// Start a new log file at application startup (default log name)
        /// </summary>
        public static void InitLogFile()
        {
            if (!IsEnabled) return;
            InitLogFile("cabletraysleeveplacer");
        }

        /// <summary>
        /// Start a new log file with a custom file name (without extension)
        /// </summary>
        public static void InitLogFile(string logFileName)
        {
            if (!IsEnabled) return;
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
                // Initialize writer
                EnsureWriterInitialized(LogFilePath, header);
            }
            catch (Exception ex)
            {
                // Log to default log if custom log creation fails
                string fallbackLog = Path.Combine("c:\\JSE_CSharp_Projects\\JSE_RevitAddin_MEP_OPENINGS\\JSE_RevitAddin_MEP_OPENINGS\\Logs", "cabletraysleeveplacer.log");
                string msg = $"[LOGGER ERROR] Could not create custom log file '{logFileName}': {ex.Message}\n{ex.StackTrace}\n";
                try { File.AppendAllText(fallbackLog, msg); } catch { /* ignore */ }
            }
        }

        /// <summary>
        /// Start a new log file with a custom file name and build timestamp
        /// </summary>
        public static void InitCustomLogFile(string logFileName)
        {
            if (!IsEnabled) return;
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
            if (!IsEnabled) return;
            try
            {
                // Set log file path under the hard-coded log directory (overwrite mode)
                if (!Directory.Exists(LogDir)) Directory.CreateDirectory(LogDir);
                LogFilePath = Path.Combine(LogDir, $"{logFileName}.log");

                // Include build/version information
                var buildTimestamp = File.GetLastWriteTime(_cachedAssemblyPath).ToString("o");
                string header =
                    $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                    $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                    $"Build Version: {_cachedVersion}\n" +
                    $"Build Timestamp: {buildTimestamp}\n" +
                    $"Wrote: {_cachedAssemblyPath}\n" +
                    $"====================================================\n";
                EnsureWriterInitialized(LogFilePath, header, overwrite: true);
            }
            catch
            {
                // Silently fail - we don't want logging to break the application
            }
        }

        /// <summary>
        /// Initialize log file using an absolute path (full filename). Creates directory if needed.
        /// </summary>
        public static void InitAbsoluteLogFile(string absoluteFilePath)
        {
            if (!IsEnabled) return;
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
                EnsureWriterInitialized(LogFilePath, header);
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
                    var fsNew = new FileStream(path, fileMode, FileAccess.Write, FileShare.Read);
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

        /// <summary>
        /// Standard log method for backward compatibility
        /// </summary>
        public static void Log(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsEnabled) return;
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
            if (!IsEnabled) return;
            try
            {
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
            if (!IsEnabled) return;
            Log(LogLevel.Warning, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Log an error message
        /// </summary>
        public static void Error(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsEnabled) return;
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
                using (FileStream fs = new FileStream(LogFilePath, FileMode.Append, FileAccess.Write))
                {
                    return fs.CanWrite;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
