using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Logging
{
    /// <summary>
    /// Debug logger specifically for remark operations
    /// </summary>
    public static class RemarkDebugLogger
    {
        private static readonly string LogFileName = "remark_debug.log";
        private static readonly string LogDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "JSE_RevitAddin_Logs"
        );

        static RemarkDebugLogger()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }
            }
            catch { }
        }

        public static void Info(string message)
        {
            Log("INFO", message);
        }

        public static void LogInfo(string message)
        {
            Log("INFO", message);
        }

        public static void LogStep(string message)
        {
            Log("STEP", message);
        }

        public static void LogError(string message)
        {
            Log("ERROR", message);
        }

        public static void Warning(string message)
        {
            Log("WARNING", message);
        }

        public static void Error(string message)
        {
            Log("ERROR", message);
        }

        public static void LogError(string message, Exception ex)
        {
            Log("ERROR", $"{message}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
        }

        private static void Log(string level, string message)
        {
            if (DeploymentConfiguration.DeploymentMode) return;

            try
            {
                string logPath = Path.Combine(LogDirectory, LogFileName);
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }
            catch { }
        }
    }
}
