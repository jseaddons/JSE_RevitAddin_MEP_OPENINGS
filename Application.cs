using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Security;
using Nice3point.Revit.Toolkit.External;
using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

using JSE_RevitAddin_MEP_OPENINGS.Services;
namespace JSE_RevitAddin_MEP_OPENINGS
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            // IMMEDIATE LOGGING - Create file as soon as add-in loads (Build: 2025-11-20 14:30)
            try
            {
                // ✅ CONFIGURATION: Use AppData\Roaming (consistent with SafeFileLogger)
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
                Directory.CreateDirectory(logDir);
                string startupLogPath = Path.Combine(logDir, "addin_startup.log");
                
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ========================================\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ADD-IN STARTUP BUILD 2025-11-20 14:30\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ========================================\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Assembly: {System.Reflection.Assembly.GetExecutingAssembly().Location}\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Process: {System.Diagnostics.Process.GetCurrentProcess().ProcessName}\n");
            }
            catch (Exception startupEx)
            {
                // Try alternative location if main log fails
                try
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), "addin_startup.log");
                    File.AppendAllText(tempPath, $"[{DateTime.Now}] ADD-IN STARTUP FAILED: {startupEx.Message}\n");
                }
                catch { }
            }

            // Initialize logging first so we can capture any startup failures
            try
            {
                // ✅ CONFIGURATION: Use AppData\Roaming (consistent with SafeFileLogger)
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
                string startupLogPath = Path.Combine(logDir, "addin_startup.log");

                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] About to CreateLogger()\n");
                CreateLogger();
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] CreateLogger() DONE\n");
                
                // ✅ CRITICAL: Copy native SQLite DLL to temporary execution directory
                // Revit copies the add-in DLL to a temp directory but doesn't copy native DLLs
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] About to CopyNativeSqliteDllToExecutionDirectory()\n");
                CopyNativeSqliteDllToExecutionDirectory();
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] CopyNativeSqliteDllToExecutionDirectory() DONE\n");

                // ✅ VERIFY SQLITE DEPLOYMENT (managed + native)
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] About to SqliteDeploymentVerifier.Verify()\n");
                Services.Diagnostics.SqliteDeploymentVerifier.Verify(startupLogPath);
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] SqliteDeploymentVerifier.Verify() DONE\n");
                
                // Initialize optimization services (Phase 1 - Foundation)
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] About to InitializeOptimizationServices()\n");
                InitializeOptimizationServices();
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] InitializeOptimizationServices() DONE\n");

                // Validate license before startup - Simple JSE domain check
                bool licensed = true;
                try
                {
                    licensed = LicenseValidator.ValidateLicense() && LicenseValidator.ValidateHardwareId();
                }
                catch (Exception lex)
                {
                    // Log license validation exceptions but do not throw — allow ribbon creation for diagnostics
                    Log.Error(lex, "License validation threw an exception");
                    TaskDialog.Show("License Warning", "License validation encountered an error: " + lex.Message);
                    licensed = false;
                }

                if (!licensed)
                {
                    // Inform the user but continue startup to avoid causing Revit to report an external command failure
                    TaskDialog.Show("License Warning",
                        "This add-in may not be licensed for this machine. The ribbon will still be created for testing and diagnostics.");
                }

                CreateRibbon();
            }
            catch (Exception ex)
            {
                // Catch any unexpected startup errors, log them and prevent them from propagating to Revit
                try { Log.Fatal(ex, "Unhandled exception during OnStartup"); } catch { }
                try { TaskDialog.Show("Startup Error", "An error occurred initializing the add-in: " + ex.Message); } catch { }
            }
        }

        public override void OnShutdown()
        {
            Log.CloseAndFlush();
        }

        private void CreateRibbon()
        {
            // ✅ CONFIGURATION: Use AppData\Roaming path for ribbon logs (consistent with SafeFileLogger)
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
            string ribbonLogPath = Path.Combine(logDir, "ribbon_creation.log");
            
            // Log ribbon creation start
            try
            {
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] CreateRibbon() started\n");
            }
            catch { }

            var panel = Application.CreatePanel("Commands", "JSE_RevitAddin_MEP_OPENINGS");
            if (panel == null)
            {
                // Panel creation failed for some environment; skip ribbon creation to avoid null references.
                try
                {
                    File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Panel creation FAILED - null panel returned\n");
                }
                catch { }
                return;
            }

            try
            {
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Panel created successfully\n");
            }
            catch { }

            // Note: Execute and Test Profile buttons removed - JSE Openings handles everything
            // The "JSE Openings" button provides access to profile setup, profile management, and main UI

        // Main JSE Openings Command - handles profile setup, management, and main UI
        var button5 = panel.AddPushButton<TestProfileManagementCommand>("JSE Openings");
        button5.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
        button5.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

        // ✅ NEW: Update XML Command - updates XML after manual cluster sleeve adjustments
        var buttonUpdateXml = panel.AddPushButton<UpdateXmlCommand>("Update XML");
        buttonUpdateXml.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
        buttonUpdateXml.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
        buttonUpdateXml.ToolTip = "Update XML files after manual cluster sleeve changes. ⚠️ Expensive operation - use only after resizing/joining sleeves.";

        // ✅ V2: Parameter Service Command V2 - standalone parameter transfer functionality (NEW VERSION)
        var button6 = panel.AddPushButton<TestParameterServiceDialogV2Command>("Parameter Service");
        button6.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
        button6.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            try
            {
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] TestProfileManagementCommand button added to ribbon\n");
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Ribbon creation COMPLETED\n");
            }
            catch { }

            // Deleted commands removed from ribbon: DeletePipeSleevesCommand, GetSleeveSummaryCommand
        }


        private static void CreateLogger()
        {
            const string outputTemplate = "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}";

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Debug(LogEventLevel.Debug, outputTemplate)
                .MinimumLevel.Debug()
                .CreateLogger();

            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                var exception = (Exception)args.ExceptionObject;
                Log.Fatal(exception, "Domain unhandled exception");
            };
        }
        
        /// <summary>
        /// Initialize optimization services for Phase 1 Foundation optimizations
        /// </summary>
        private void InitializeOptimizationServices()
        {
            try
            {
                // Load optimization flags from configuration
                OptimizationFlags.LoadFromConfiguration();
                
                // Initialize cache invalidation monitor
                if (OptimizationFlags.UseCacheInvalidation)
                {
                    DebugLogger.Info("[Application] Initialized Cache Invalidation Monitor");
                }
                
                // Initialize memory management service
                if (OptimizationFlags.UseMemoryManagement)
                {
                    DebugLogger.Info("[Application] Initialized Memory Management Service");
                }
                
                // Initialize smart tolerance service
                if (OptimizationFlags.UseSmartTolerance)
                {
                    DebugLogger.Info("[Application] Initialized Smart Tolerance Service");
                }
                
                // Log optimization status
                DebugLogger.Info($"[Application] Optimization Services Initialized:\n{OptimizationFlags.GetOptimizationStatus()}");
                
                // Log memory statistics
                var memoryStats = MemoryManagementService.GetMemoryStatistics();
                DebugLogger.Info($"[Application] Initial Memory Statistics: {memoryStats}");
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[Application] Error initializing optimization services: {ex.Message}");
                // Fallback to safe defaults
                OptimizationFlags.ResetToSafeDefaults();
            }
        }
        
        /// <summary>
        /// Copy e_sqlite3.dll to the temporary execution directory where Revit loads the add-in
        /// Revit creates a temp directory like: C:\Users\...\AppData\Local\Temp\RevitAddins\JSE_RevitAddin_MEP_OPENINGS-Executing-{timestamp}\
        /// The native DLL must be in the same directory as the executing DLL for SQLite to work
        /// </summary>
        private static void CopyNativeSqliteDllToExecutionDirectory()
        {
            try
            {
                var executingAssembly = Assembly.GetExecutingAssembly();
                var executionDirectory = Path.GetDirectoryName(executingAssembly.Location);
                if (string.IsNullOrEmpty(executionDirectory))
                {
                    Log.Warning("[SQLite] Execution directory is null; cannot verify native dependencies.");
                    return;
                }

                Log.Debug($"[SQLite] Execution directory: {executionDirectory}");

                var dependencies = new (string RelativePath, bool PreferX64)[]
                {
                    ("System.Data.SQLite.dll", false),
                    (Path.Combine("x64", "SQLite.Interop.dll"), true)
                };

                foreach (var dependency in dependencies)
                {
                    var targetPath = Path.Combine(executionDirectory, dependency.RelativePath);
                    if (File.Exists(targetPath))
                    {
                        Log.Debug($"[SQLite] Dependency already present: {targetPath}");
                        continue;
                    }

                    var sourcePath = LocateDependency(dependency.RelativePath, dependency.PreferX64);
                    if (sourcePath == null)
                    {
                        Log.Warning($"[SQLite] ⚠️ Unable to locate dependency '{dependency.RelativePath}'. SQLite may fail to load.");
                        continue;
                    }

                    try
                    {
                        var destinationDirectory = Path.GetDirectoryName(targetPath);
                        if (!string.IsNullOrEmpty(destinationDirectory))
                        {
                            Directory.CreateDirectory(destinationDirectory);
                        }

                        File.Copy(sourcePath, targetPath, overwrite: true);
                        Log.Information($"[SQLite] ✅ Copied dependency: {sourcePath} -> {targetPath}");
                    }
                    catch (Exception copyEx)
                    {
                        Log.Error(copyEx, $"[SQLite] ❌ Failed to copy dependency '{dependency.RelativePath}'");
                    }
                }
            }
            catch (Exception ex)
            {
                // Don't fail startup if DLL copy fails - SQLite will handle the error gracefully
                Log.Warning(ex, "[SQLite] Error during native DLL copy operation");
            }

            static string? LocateDependency(string relativePath, bool preferX64)
            {
                var fileName = Path.GetFileName(relativePath);
                var candidatePaths = new List<string>();

                void AddPath(string? path)
                {
                    if (!string.IsNullOrEmpty(path))
                    {
                        candidatePaths.Add(path);
                    }
                }

                // Supported versions pruned to 2023 & 2024 (2025 temporarily removed)
                var revitVersions = new[] { "2023", "2024" };
                foreach (var version in revitVersions)
                {
                    var appDataDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        @"Autodesk\Revit\Addins", version);
                    AddPath(Path.Combine(appDataDir, fileName));
                    AddPath(Path.Combine(appDataDir, "x64", fileName));

                    var programDataDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                        @"Autodesk\Revit\Addins", version);
                    AddPath(Path.Combine(programDataDir, fileName));
                    AddPath(Path.Combine(programDataDir, "x64", fileName));
                }

                try
                {
                    var addinManifests = new[]
                    {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2024\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2023\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"Autodesk\Revit\Addins\2024\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"Autodesk\Revit\Addins\2023\JSE_RevitAddin_MEP_OPENINGS.addin")
                    };

                    foreach (var manifest in addinManifests)
                    {
                        if (!File.Exists(manifest)) continue;

                        var addinXml = XDocument.Load(manifest);
                        var assemblyElement = addinXml.Descendants("Assembly").FirstOrDefault();
                        if (assemblyElement == null) continue;

                        var assemblyPath = assemblyElement.Value;
                        if (!Path.IsPathRooted(assemblyPath) || !File.Exists(assemblyPath)) continue;

                        var assemblyDir = Path.GetDirectoryName(assemblyPath);
                        AddPath(Path.Combine(assemblyDir ?? string.Empty, fileName));
                        AddPath(Path.Combine(assemblyDir ?? string.Empty, "x64", fileName));
                    }
                }
                catch (Exception manifestEx)
                {
                    Log.Debug($"[SQLite] Error parsing .addin manifests: {manifestEx.Message}");
                }

                var nugetRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    @".nuget\packages\system.data.sqlite.core\1.0.118.0");

                if (relativePath.EndsWith("System.Data.SQLite.dll", StringComparison.OrdinalIgnoreCase))
                {
                    AddPath(Path.Combine(nugetRoot, @"lib\net46\System.Data.SQLite.dll"));
                    AddPath(Path.Combine(nugetRoot, @"build\net46\System.Data.SQLite.dll"));
                }
                else if (relativePath.EndsWith("SQLite.Interop.dll", StringComparison.OrdinalIgnoreCase))
                {
                    AddPath(Path.Combine(nugetRoot, @"runtimes\win-x64\native\SQLite.Interop.dll"));
                }

                // .addin manifest lookup – handles deployments where files live in a custom subfolder
                try
                {
                    var manifestDirectories = new[]
                    {
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2024"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2023"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"Autodesk\Revit\Addins\2024"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"Autodesk\Revit\Addins\2023")
                    };

                    foreach (var manifestDir in manifestDirectories)
                    {
                        if (!Directory.Exists(manifestDir)) continue;

                        foreach (var manifest in Directory.EnumerateFiles(manifestDir, "JSE_RevitAddin_MEP_OPENINGS.addin", SearchOption.AllDirectories))
                        {
                            try
                            {
                                var addinXml = XDocument.Load(manifest);
                                var assemblyElement = addinXml.Descendants("Assembly").FirstOrDefault();
                                if (assemblyElement == null) continue;

                                var assemblyPath = assemblyElement.Value;
                                if (!Path.IsPathRooted(assemblyPath) || !File.Exists(assemblyPath)) continue;

                                var assemblyDir = Path.GetDirectoryName(assemblyPath);
                                AddPath(Path.Combine(assemblyDir ?? string.Empty, fileName));
                                AddPath(Path.Combine(assemblyDir ?? string.Empty, "x64", fileName));
                            }
                            catch (Exception manifestEx)
                            {
                                Log.Debug($"[SQLite] Error parsing manifest '{manifest}': {manifestEx.Message}");
                            }
                        }
                    }
                }
                catch (Exception manifestSearchEx)
                {
                    Log.Debug($"[SQLite] Error searching manifests for dependency locations: {manifestSearchEx.Message}");
                }

                string? fallback = null;
                foreach (var path in candidatePaths)
                {
                    if (!File.Exists(path)) continue;

                    if (preferX64 && path.IndexOf("x64", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return path;
                    }

                    if (fallback == null)
                    {
                        fallback = path;
                    }
                }

                return fallback;
            }
        }
    }
}
