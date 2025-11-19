using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{    /// <summary>
     /// Logger for fire damper debugging, writes to a separate log file.
     /// </summary>
    public static class DamperLogger
    {
    // ✅ FIXED: Use SafeFileLogger for deployment-compatible paths (no hardcoded paths)
    private static string LogFilePath => SafeFileLogger.GetLogFilePath("RevitAddin_DamperDebug.log");

        public static void InitLogFile()
        {
            if (!DebugLogger.IsEnabled) return;
            try
            {
                // Clear existing log file and write header
                var header = $"=== Fire Damper Debug Log Started at {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={Environment.NewLine}";
                File.WriteAllText(LogFilePath, header);
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
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logEntry = $"[{timestamp}] {message}{Environment.NewLine}";
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(LogFilePath, logEntry);
                }
            }
            catch
            {
                // Logging failures should not break execution
            }
        }
    }
}
