using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{    /// <summary>
     /// Logger for fire damper debugging, writes to a separate log file.
     /// </summary>
    public static class DamperLogger
    {
    // Match the project's Log directory used by DebugLogger for consistency
    private static readonly string LogFilePath = Path.Combine("C:", "JSE_CSharp_Projects", "JSE_MEPOPENING_23", "Log", "RevitAddin_DamperDebug.log");

        public static void InitLogFile()
        {
            // Logging disabled - no action needed
            return;
        }

        public static void Log(string message)
        {
            // Logging disabled - no action needed
            return;
        }
    }
}
