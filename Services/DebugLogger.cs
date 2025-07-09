// (moved to correct location below)
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
        /// Log a message to a custom log file in the user's Documents folder (does not affect the main log file)
        /// </summary>
        public static void LogToFile(string fileName, string message)
        {
            try
            {
                string customLogPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);
                string version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [v{version}] {message}{Environment.NewLine}";
                File.AppendAllText(customLogPath, logEntry);
            }
            catch
            {
                // Logging failures should not break execution
            }
        }
        private static string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "RevitAddin_Debug.log");

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
            InitLogFile("RevitAddin_Debug");
        }

        /// <summary>
        /// Start a new log file with a custom file name (without extension)
        /// </summary>
        public static void InitLogFile(string logFileName)
        {
            try
            {
                // Set log file path
                LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), logFileName + ".log");
                // Include build/version information
                var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
                var buildTimestamp = File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location).ToString("o");
                string header =
                    $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                    $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                    $"Build Version: {version}\n" +
                    $"Build Timestamp: {buildTimestamp}\n" +
                    $"====================================================\n";
                File.WriteAllText(LogFilePath, header);
            }
            catch
            {
                // Silently fail - we don't want logging to break the application
            }
        }

        /// <summary>
        /// Start a new log file with a custom file name and build timestamp
        /// </summary>
        public static void InitCustomLogFile(string logFileName)
        {
            try
            {
                // Set log file path with build timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"{logFileName}_{timestamp}.log");

                // Include build/version information
                var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
                var buildTimestamp = File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location).ToString("o");
                string header =
                    $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                    $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                    $"Build Version: {version}\n" +
                    $"Build Timestamp: {buildTimestamp}\n" +
                    $"====================================================\n";
                File.WriteAllText(LogFilePath, header);
            }
            catch
            {
                // Silently fail - we don't want logging to break the application
            }
        }

        /// <summary>
        /// Start a new log file and overwrite existing content (backward compatibility)
        /// </summary>
        public static void InitLogFileOverwrite(string logFileName)
        {
            InitCustomLogFileOverwrite(logFileName);
        }

        /// <summary>
        /// Start a new log file with a custom file name and overwrite existing content
        /// </summary>
        public static void InitCustomLogFileOverwrite(string logFileName)
        {
            try
            {
                // Set log file path
                LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"{logFileName}.log");

                // Include build/version information
                var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
                var buildTimestamp = File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location).ToString("o");
                string header =
                    $"===== NEW LOG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n" +
                    $"JSE_RevitAddin_MEP_OPENINGS Debug Log: {logFileName}\n" +
                    $"Build Version: {version}\n" +
                    $"Build Timestamp: {buildTimestamp}\n" +
                    $"====================================================\n";
                File.WriteAllText(LogFilePath, header); // Overwrite existing content
            }
            catch
            {
                // Silently fail - we don't want logging to break the application
            }
        }

        /// <summary>
        /// Standard log method for backward compatibility
        /// </summary>
        public static void Log(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
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
            try
            {
                // Get the class name from the source file path
                string className = Path.GetFileNameWithoutExtension(sourceFile);
                // Get assembly version for this entry
                string version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "n/a";
                // Format the log entry with timestamp, version, level, class and line
                string levelText = level.ToString().ToUpper();
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [v{version}] [{levelText}] [{className}:{lineNumber}] {message}{Environment.NewLine}";
                File.AppendAllText(LogFilePath, logEntry);
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
            Log(LogLevel.Warning, message, sourceFile, lineNumber);
        }

        /// <summary>
        /// Log an error message
        /// </summary>
        public static void Error(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
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
        /// Log a detailed message (backward compatibility)
        /// </summary>
        public static void LogDetailed(string message,
            [CallerFilePath] string sourceFile = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            Log(LogLevel.Debug, $"DETAILED: {message}", sourceFile, lineNumber);
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
