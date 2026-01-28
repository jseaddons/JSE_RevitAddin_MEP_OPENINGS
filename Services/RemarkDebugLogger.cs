using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Specialized logger for debugging Remark Selected functionality.
    /// Writes to remark_debug.log even in deployment mode if needed.
    /// </summary>
    public static class RemarkDebugLogger
    {
        private static readonly string LogFileName = "parameter_ui.log";

        /// <summary>
        /// Logs a high-level step in the remarking process.
        /// </summary>
        public static void LogStep(string message)
        {
            LogInternal($"[STEP] {message}");
        }

        /// <summary>
        /// Logs specific selection data (checkbox states, prefix values).
        /// </summary>
        public static void LogSelection(string key, object value)
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
                // We use SafeAppendTextAlways because we WANT these logs for debugging 
                // even if the general DeploymentMode says no.
                SafeFileLogger.SafeAppendTextAlways(LogFileName, message);
                
                // Also write to standard debug output for Visual Studio attached sessions
                System.Diagnostics.Debug.WriteLine($"[REMARK_DEBUG] {message}");
            }
            catch
            {
                // Silently fail if logging fails
            }
        }
    }
}
