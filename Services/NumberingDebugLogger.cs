using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Specialized logger for debugging Numbering/Prefix functionality.
    /// Writes to numbering_debug.log as requested by user.
    /// </summary>
    public static class NumberingDebugLogger
    {
        private static readonly string LogFileName = "numbering_debug.log";

        /// <summary>
        /// Logs a high-level step.
        /// </summary>
        public static void LogStep(string message)
        {
            LogInternal($"[STEP] {message}");
        }

        /// <summary>
        /// Logs specific data points.
        /// </summary>
        public static void LogData(string key, object value)
        {
            LogInternal($"[DATA] {key}: {value}");
        }

        /// <summary>
        /// Logs an error with optional exception details.
        /// </summary>
        public static void LogError(string message, Exception ex = null)
        {
            string error = $"[ERROR] {message}";
            if (ex != null)
            {
                error += $"\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
            }
            LogInternal(error);
        }

        /// <summary>
        /// Logs informational messages.
        /// </summary>
        public static void LogInfo(string message)
        {
            LogInternal($"[INFO] {message}");
        }

        private static void LogInternal(string message)
        {
            try
            {
                // FORCE log to file regardless of deployment configuration
                SafeFileLogger.SafeAppendTextAlways(LogFileName, message);
                
                // Also write to Debug for VS
                System.Diagnostics.Debug.WriteLine($"[NUMBERING_DEBUG] {message}");
            }
            catch
            {
                // Silently fail to avoid crashing the app
            }
        }
    }
}
