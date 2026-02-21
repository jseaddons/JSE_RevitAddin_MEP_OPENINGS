using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// ⚠️⚠️⚠️ PROTECTED CODE: CRITICAL VERSION DETECTION LOGIC ⚠️⚠️⚠️
    /// DO NOT MODIFY, SIMPLIFY, OR REMOVE THIS LOGIC.
    /// This file is the central source of truth for Revit versioning across the entire project.
    /// Incorrect changes here will break logging, database paths, and API compatibility.
    /// 
    /// Centralized Revit version info abstraction using compile-time constants.
    /// Keeps runtime checks lightweight and avoids scattering #if blocks.
    /// </summary>
    public static class VersionInfo
    {
        // ✅ EXHAUSTIVE VERSION DETECTION
#if REVIT2026
        public const int CurrentMajor = 2026;
        public const bool Is2024Plus = true;
        public const bool Is2025Plus = true;
        public const bool Is2026Plus = true;
#elif REVIT2025
        public const int CurrentMajor = 2025;
        public const bool Is2024Plus = true;
        public const bool Is2025Plus = true;
        public const bool Is2026Plus = false;
#elif REVIT2024
        public const int CurrentMajor = 2024;
        public const bool Is2024Plus = true;
        public const bool Is2025Plus = false;
        public const bool Is2026Plus = false;
#elif REVIT2023
        public const int CurrentMajor = 2023;
        public const bool Is2024Plus = false;
        public const bool Is2025Plus = false;
        public const bool Is2026Plus = false;
#else
        // Fallback for design-time or generic builds
        public const int CurrentMajor = 2025;
        public const bool Is2024Plus = true;
        public const bool Is2025Plus = true;
        public const bool Is2026Plus = false;
#endif

        /// <summary>
        /// ⚠️ PROTECTED: Returns true if current build targets at least the specified Revit major version.
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
