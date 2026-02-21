using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{    /// <summary>
     /// Logger for fire damper debugging, writes to a separate log file.
     /// </summary>
    public static class DamperLogger
    {
    private static string LogFileName => "RevitAddin_DamperDebug.log";

        public static void InitLogFile()
        {
            if (!DebugLogger.IsEnabled) return;
            try
            {
                // Clear existing log file and write header
                var header = $"=== Fire Damper Debug Log Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n";
                SafeFileLogger.SafeAppendTextAlways(LogFileName, header);
            }
            catch
            {
                // Logging failures should not break execution
            }
        }

        public static void Log(string message)
        {
            if (!DebugLogger.IsEnabled) return;
            try
            {
                // ✅ PERFORMANCE: Safe logging via SafeFileLogger
                SafeFileLogger.SafeAppendText(LogFileName, message);
            }
            catch
            {
                // Logging failures should not break execution
            }
        }
    }
}
