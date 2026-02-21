using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class DebugLogger
    {
        public static bool IsEnabled = false;
        public static string CurrentService = "MainUI";

        public enum LogLevel { Debug, Info, Warning, Error }

        private static bool IsLoggingEnabledForCurrentService()
        {
            if (CurrentService == "CombinedSleeveManual" || CurrentService == "CombinedSleeveAuto")
                return true;
                
            if (OptimizationFlags.DisableVerboseLogging)
                return false;
                
            if (DeploymentConfiguration.DeploymentMode)
                return false;
                
            return IsEnabled;
        }

        public static void Log(LogLevel level, string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0)
        {
            if (!IsLoggingEnabledForCurrentService()) return;

            string className = Path.GetFileNameWithoutExtension(sourceFile);
            string levelText = level.ToString().ToUpper();
            string logEntry = $"[{levelText}] [{className}:{lineNumber}] {message}";

            string fileName = GetLogFileNameForService();
            SafeFileLogger.SafeAppendTextAlways(fileName, logEntry);
        }

        public static void Log(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Log(LogLevel.Info, message, sourceFile, lineNumber);

        private static string GetLogFileNameForService()
        {
            switch (CurrentService)
            {
                case "Duct": return "ductsleeveplacer.log";
                case "CableTray": return "cabletraysleeveplacer.log";
                case "Damper": return "dampersleeveplacer.log";
                case "CombinedSleeveManual":
                case "CombinedSleeveAuto": return "combinesleeveplacer.log";
                default: return "refresh.log";
            }
        }

        public static void Info(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Log(LogLevel.Info, message, sourceFile, lineNumber);

        public static void Warning(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Log(LogLevel.Warning, message, sourceFile, lineNumber);

        public static void Error(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Log(LogLevel.Error, message, sourceFile, lineNumber);

        public static void Debug(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Log(LogLevel.Debug, message, sourceFile, lineNumber);

        public static void Critical(string message, [CallerFilePath] string sourceFile = "", [CallerLineNumber] int lineNumber = 0) => 
            Error($"CRITICAL: {message}", sourceFile, lineNumber);

        public static void InitLogFile(string logFileName) { }
        public static void InitLogFile() { }
        public static void CloseAllLogFiles() { }
        public static void SetServiceContext(string serviceName) => CurrentService = serviceName;
        public static void CleanupOldLogs() { }
        
        public static void SetDuctLogFile() { CurrentService = "Duct"; }
        public static void SetCableTrayLogFile() { CurrentService = "CableTray"; }
        public static void SetDamperLogFile() { CurrentService = "Damper"; }
        public static void SetCombinedSleeveLogFile() { CurrentService = "CombinedSleeveManual"; }
    }
}
