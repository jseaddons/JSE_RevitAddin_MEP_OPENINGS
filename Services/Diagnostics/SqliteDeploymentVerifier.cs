using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Serilog;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Diagnostics
{
    /// <summary>
    /// Verifies SQLite deployment (managed + native) and logs actionable diagnostics.
    /// Suppresses false "missing dependency" warnings when assembly is already merged or loaded.
    /// </summary>
    public static class SqliteDeploymentVerifier
    {
        private static bool _verified;
        public static void Verify(string? additionalLogFile = null)
        {
            if (_verified) return; // Run once per session
            _verified = true;

            string executionDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            
#if NET8_0_OR_GREATER
            string managedName = "Microsoft.Data.Sqlite.dll";
            string nativeName = "e_sqlite3.dll";
            string assemblyName = "Microsoft.Data.Sqlite";
#else
            string managedName = "System.Data.SQLite.dll";
            string nativeName = "SQLite.Interop.dll"; // Usually in x64 subfolder
            string assemblyName = "System.Data.SQLite";
#endif

            string managedPath = Path.Combine(executionDir, managedName);
            string nativePathDirect = Path.Combine(executionDir, nativeName);
            string nativePathX64 = Path.Combine(executionDir, "x64", nativeName);

            bool assemblyLoaded = AppDomain.CurrentDomain
                .GetAssemblies()
                .Any(a => string.Equals(a.GetName().Name, assemblyName, StringComparison.OrdinalIgnoreCase));

            // Attempt explicit load if not already loaded and file exists
            if (!assemblyLoaded && File.Exists(managedPath))
            {
                try
                {
                    Assembly.LoadFrom(managedPath);
                    assemblyLoaded = true;
                    Log.Information("[SQLite] ✅ Explicitly loaded managed provider from {Path}", managedPath);
                }
                catch (Exception loadEx)
                {
                    Log.Warning(loadEx, "[SQLite] ⚠️ Failed to explicitly load managed provider {Path}", managedPath);
                }
            }

            bool managedPresent = File.Exists(managedPath);
            bool nativePresent = File.Exists(nativePathDirect) || File.Exists(nativePathX64);

            // Log summary
            string summary = $"ExecutionDir={executionDir}; ManagedPresent={managedPresent}; NativePresent={nativePresent}; AssemblyLoaded={assemblyLoaded}; ManagedPath={managedPath}; NativePathDirect={nativePathDirect}";
            Log.Information("[SQLite] Deployment Summary: {Summary}", summary);

            if (additionalLogFile != null)
            {
                try
                {
                    File.AppendAllText(additionalLogFile, $"[{DateTime.Now}] [SQLite] Deployment Summary: {summary}\n");
                }
                catch { /* ignore file logging failures */ }
            }

            // Actionable guidance if something is missing
            if (!managedPresent && !assemblyLoaded)
            {
                Log.Warning("[SQLite] ❌ Managed provider missing ({ManagedName}) AND not loaded.", managedName);
            }
            else if (!nativePresent)
            {
                Log.Warning("[SQLite] ❌ Native interop DLL missing. Expected {NativeName} in root or x64 subfolder.", nativeName);
            }
            else
            {
                Log.Information("[SQLite] ✅ All required managed/native components available ({ManagedName}, {NativeName}).", managedName, nativeName);
            }
        }
    }
}
