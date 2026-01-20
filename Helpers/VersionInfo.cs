using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
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
        /// Gets the build timestamp based on the assembly's last write time.
        /// </summary>
        public static string GetBuildTimestamp()
        {
            try
            {
                var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var fileInfo = new System.IO.FileInfo(assemblyLocation);
                return fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch
            {
                return "Unknown";
            }
        }
    }
}
