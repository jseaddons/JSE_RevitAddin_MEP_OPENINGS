using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Dedicated logger for UI persistence debugging
    /// Writes to: Log/R2023/R2023/ui_persistence_debug.log
    /// </summary>
    public static class UIPersistenceDebugLogger
    {
        private static readonly string LogDirectory = Path.Combine(
            Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "",
            "Log", "R2023", "R2023"
        );

        private static readonly string LogFilePath = Path.Combine(LogDirectory, "ui_persistence_debug.log");

        static UIPersistenceDebugLogger()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }
            }
            catch { /* Ignore directory creation errors */ }
        }

        public static void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        public static void LogWarning(string message)
        {
            WriteLog("WARN", message);
        }

        public static void LogError(string message)
        {
            WriteLog("ERROR", message);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string logEntry = $"[{timestamp}] [{level}] {message}";
                
                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently fail if logging fails
            }
        }

        /// <summary>
        /// Clear the log file (useful for starting fresh)
        /// </summary>
        public static void ClearLog()
        {
            try
            {
                if (File.Exists(LogFilePath))
                {
                    File.Delete(LogFilePath);
                }
            }
            catch { /* Ignore */ }
        }
    }
}
