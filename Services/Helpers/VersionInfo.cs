using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Reflection;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Helpers
{
    /// <summary>
    /// Centralized Revit version info abstraction using compile-time constants.
    /// Keeps runtime checks lightweight and avoids scattering #if blocks.
    /// Extend by adding new compile constants in the csproj when supporting further versions.
    /// </summary>
    public static class VersionInfo
    {
#if REVIT2024_OR_GREATER
        public const int CurrentMajor = 2024;
        public const bool Is2024Plus = true;
#else
        public const int CurrentMajor = 2023;
        public const bool Is2024Plus = false;
#endif
        /// <summary>
        /// Returns true if current build targets at least the specified Revit major version.
        /// </summary>
        public static bool IsAtLeast(int major) => CurrentMajor >= major;

        /// <summary>
        /// Convenience: True if this build should activate behaviors for 2023 or older branch.
        /// </summary>
        public static bool Is2023Branch => !Is2024Plus;

        /// <summary>
        /// Returns a string identifier for logging or diagnostics.
        /// </summary>
        public static string VersionTag => $"R{CurrentMajor}";
        
        /// <summary>
        /// Gets the build timestamp from the assembly's last write time.
        /// Useful for logging when the add-in was built/deployed.
        /// </summary>
        /// <returns>DateTime of when the assembly was last compiled</returns>
        public static DateTime GetBuildTimestamp()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var assemblyPath = assembly.Location;
                
                if (!string.IsNullOrEmpty(assemblyPath) && File.Exists(assemblyPath))
                {
                    var fileInfo = new FileInfo(assemblyPath);
                    return fileInfo.LastWriteTime;
                }
                
                return DateTime.Now;
            }
            catch
            {
                return DateTime.Now;
            }
        }

        /// <summary>
        /// Gets the build timestamp as a formatted string.
        /// </summary>
        /// <returns>Formatted timestamp (yyyy-MM-dd HH:mm:ss)</returns>
        public static string GetBuildTimestampString()
        {
            return GetBuildTimestamp().ToString("yyyy-MM-dd HH:mm:ss");
        }
    
    }
}
