using System;
using System.Runtime.InteropServices;
using Serilog;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class NativeLibraryLoader
    {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        /// <summary>
        /// Explicitly load a native library from a specific path.
        /// Useful for resolving native dependencies in Revit Addin Manager.
        /// </summary>
        public static bool LoadNativeLibrary(string libraryPath)
        {
            try
            {
                if (!System.IO.File.Exists(libraryPath))
                {
                    Log.Warning("[NativeLoader] ⚠️ Cannot load library, file not found: {Path}", libraryPath);
                    return false;
                }

                IntPtr handle = LoadLibrary(libraryPath);
                if (handle == IntPtr.Zero)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    Log.Error("[NativeLoader] ❌ Failed to load library: {Path}. Error Code: {Error}", libraryPath, errorCode);
                    return false;
                }

                Log.Information("[NativeLoader] ✅ Successfully loaded native library: {Path}", libraryPath);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[NativeLoader] ❌ Exception while loading native library: {Path}", libraryPath);
                return false;
            }
        }
    }
}
