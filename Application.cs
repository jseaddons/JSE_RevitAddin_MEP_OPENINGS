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
using System.Text.Json;
using System.Xml.Linq;

using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
namespace JSE_RevitAddin_MEP_OPENINGS
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public class UsedImplicitlyAttribute : Attribute { }

    public class Application : ExternalApplication
    {
        private static bool _ribbonCreated = false;

        public override void OnStartup()
        {
            try
            {
                // 1. Essentials - Store app reference and subscribe to events (fast)
                m_uiApp = this.Application;
                this.Application.ViewActivated += OnViewActivated;

                // 2. Logging - Basic setup
                CreateLogger();

                // 3. Background - All heavy I/O, SQLite, License, etc.
                System.Threading.Tasks.Task.Run(() => {
                    try {
                        InitializeAppBackground();
                    } catch { }
                });

                // 4. UI - Create Ribbon (once)
                if (!_ribbonCreated)
                {
                    CreateRibbon();
                    _ribbonCreated = true;
                }
            }
            catch (Exception ex)
            {
                // Last resort fallback - don't let it crash Revit
                try { System.Diagnostics.Debug.WriteLine($"Startup Critical Failure: {ex.Message}"); } catch { }
            }
        }


        private void InitializeAppBackground()
        {
            try
            {
                // 1. License Check
                try {
                    LicenseValidator.ValidateLicense(true);
                } catch { }

                // 2. Logging and Build Info (Quietly)
                try
                {
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
                    Directory.CreateDirectory(logDir);
                    
                    string startupLogPath = Path.Combine(logDir, "addin_startup.log");
                    File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Starting background initialization...\n");
                    
                    MasterSwitch.DiagnosticLogging = false; // Ensure deployment mode
                } catch { }

                // 3. SQLite Support
                CopyNativeSqliteDllToExecutionDirectory();
#if NET8_0_OR_GREATER
                SQLitePCL.Batteries.Init();
#endif
                
                // 4. Flags and Phase 1 optimizations
                InitializeOptimizationServices();
            }
            catch { }
        }


        public override void OnShutdown()
        {
            // Unsubscribe from events
            try
            {
                if (m_uiApp != null)
                {
                    m_uiApp.ViewActivated -= OnViewActivated;
                }
            }
            catch { }

            Log.CloseAndFlush();
        }
        
        private UIControlledApplication m_uiApp;

        /// <summary>
        /// Event handler to update ApplicationProfileService context when view changes
        /// This ensures we always have the correct project path for the active document
        /// </summary>
        private void OnViewActivated(object sender, Autodesk.Revit.UI.Events.ViewActivatedEventArgs e)
        {
            try
            {
                if (e.Document != null && !e.Document.IsFamilyDocument)
                {
                    // Force update the profile service context using Document overload
                    // (internally resolves standardized AppData path via ProjectPathService)
                    ApplicationProfileService.Instance.UpdateForCurrentDocument(e.Document);
                }
            }
            catch (Exception ex)
            {
                // Fail silently but safely - don't crash Revit
                System.Diagnostics.Debug.WriteLine($"Error in OnViewActivated: {ex.Message}");
            }
        }

        private void CreateRibbon()
        {
            try
            {
                string tabName = "JSE_RevitAddin_MEP_OPENINGS";
                try
                {
                    this.Application.CreateRibbonTab(tabName);
                }
                catch { /* Tab might already exist */ }

                RibbonPanel panel = this.Application.CreateRibbonPanel(tabName, "Commands");
                
                if (panel == null) return;

                string assemblyPath = Assembly.GetExecutingAssembly().Location;


            // 1. Main JSE Openings Command
            var button5Data = new PushButtonData("cmdJseOpenings", "JSE Openings", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.TestProfileManagementCommand");
            var button5 = panel.AddItem(button5Data) as PushButton;
            if (button5 != null)
            {
                button5.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                button5.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                button5.ToolTip = "JSE MEP Openings - Version 3.0";
            }

            // 2. Combined Sleeve Manager
            var buttonCombinedData = new PushButtonData("cmdCombinedSleeve", "Combined\nSleeve", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.CombinedSleeveCommand");
            var buttonCombined = panel.AddItem(buttonCombinedData) as PushButton;
            if (buttonCombined != null)
            {
                buttonCombined.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                buttonCombined.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                buttonCombined.ToolTip = "Manage combined sleeves (Manual Join).";
            }

            // 3. Parameter Service (External Project: JSE_Parameter_Service)
            // This integrates the Parameter Service from the separate project into the same ribbon panel
            var paramServiceDll = Path.Combine(Path.GetDirectoryName(assemblyPath), "JSE_Parameter_Service.dll");
            if (File.Exists(paramServiceDll))
            {
                var btnParamService = new PushButtonData(
                    "cmdParameterService",
                    "Parameter\nService",
                    paramServiceDll,
                    "JSE_Parameter_Service.Commands.TestParameterServiceDialogV2Command");
                btnParamService.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btnParamService.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                btnParamService.ToolTip = "Open Parameter Service (from JSE_Parameter_Service project).";
                panel.AddItem(btnParamService);
            }
            else
            {
                // Optional: Add a placeholder button that informs user about the missing Parameter Service
                var btnParamServiceMissing = new PushButtonData(
                    "cmdParameterServiceMissing",
                    "Parameter\nService",
                    assemblyPath,
                    "JSE_RevitAddin_MEP_OPENINGS.Commands.ShowParameterServiceMissingCommand");
                btnParamServiceMissing.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btnParamServiceMissing.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                btnParamServiceMissing.ToolTip = "Parameter Service not found. Please build JSE_Parameter_Service project.";
                btnParamServiceMissing.AvailabilityClassName = "JSE_RevitAddin_MEP_OPENINGS.Commands.ParameterServiceMissingAvailability";
                panel.AddItem(btnParamServiceMissing);
            }

            // 4. Update DB Command
            var buttonUpdateDbData = new PushButtonData(
                "cmdUpdateDb",
                "Update DB",
                assemblyPath,
                "JSE_RevitAddin_MEP_OPENINGS.Commands.UpdateDbCommand");
            var buttonUpdateDb = panel.AddItem(buttonUpdateDbData) as PushButton;
            if (buttonUpdateDb != null)
            {
                buttonUpdateDb.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                buttonUpdateDb.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                buttonUpdateDb.ToolTip = "Synchronize Database with Model.";
            }

            // 5. BW Tagging - launches JSE_T1 SMART Tagging
            var jseT1DllPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Autodesk\Revit\Addins",
                this.Application.ControlledApplication.VersionNumber,
                @"JSE_T1\RevitAddIn1.dll");
            var bwTaggingButton = panel.AddItem(new PushButtonData(
                "BWTaggingMEPOpenings",
                "BW Tagging\nMEP Openings",
                jseT1DllPath,
                "JSE_T1.Commands.BWTaggingMEPOpeningsCommand"
            )) as PushButton;
            if (bwTaggingButton != null)
            {
                bwTaggingButton.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                bwTaggingButton.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                bwTaggingButton.ToolTip = "ver 1.  for sleeve openings only";
            }

            // 6. DIAGNOSTIC TOOLS GROUP
            var pulldownData = new PulldownButtonData("DiagnosticTools", "Diagnostic\nTools");
            pulldownData.ToolTip = "Access diagnostic and recovery tools.";
            var pulldown = panel.AddItem(pulldownData) as PulldownButton;
            if (pulldown != null)
            {
                pulldown.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                pulldown.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
                
                // 5.1 Toggle Diagnostic
                var btn1Data = new PushButtonData("cmdToggleDiagnostic", "Toggle Diagnostic", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.ToggleDiagnosticCommand");
                var btn1 = pulldown.AddPushButton(btn1Data);
                btn1.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btn1.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

                // 5.2 Diagnostic Status
                var btn2Data = new PushButtonData("cmdDiagnosticStatus", "Diagnostic Status", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.DiagnosticStatusCommand");
                var btn2 = pulldown.AddPushButton(btn2Data);
                btn2.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btn2.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

                // 5.3 Reset Flags (Box)
                var btn3Data = new PushButtonData("cmdResetFlagsBox", "Reset Flags (Box)", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.ResetFlagsInSectionBoxCommand");
                var btn3 = pulldown.AddPushButton(btn3Data);
                btn3.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btn3.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

                // 5.4 Clear All DB
                var btn4Data = new PushButtonData("cmdClearAllSleeveDb", "Clear All Sleeve DB", assemblyPath, "JSE_RevitAddin_MEP_OPENINGS.Commands.ClearAllSleeveDbTablesCommand");
                var btn4 = pulldown.AddPushButton(btn4Data);
                btn4.Image = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
                btn4.LargeImage = GetImageSource("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");
            }


                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string ribbonLogPath = Path.Combine(appData, "JSE_MEP_Openings", "Logs", "ribbon_creation.log");
                try { File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Ribbon creation COMPLETED\n"); } catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ribbon Creation Failure: {ex.Message}");
            }
        }

        private System.Windows.Media.Imaging.BitmapImage GetImageSource(string path)
        {
            try
            {
                return new System.Windows.Media.Imaging.BitmapImage(new Uri(path, UriKind.RelativeOrAbsolute));
            }
            catch
            {
                return null;
            }
        }


        private static void CreateLogger()
        {
            const string outputTemplate = "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}";

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Debug(LogEventLevel.Debug, outputTemplate)
                .MinimumLevel.Debug()
                .CreateLogger();
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
                
                // Configure logging: disable verbose logs in deployment or when diagnostics are off
                OptimizationFlags.DisableVerboseLogging = DeploymentConfiguration.DeploymentMode 
                    || !OptimizationFlags.UseDiagnosticMode;
                
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
#if NET8_0_OR_GREATER
                    ("Microsoft.Data.Sqlite.dll", false),
                    ("e_sqlite3.dll", true)
#else
                    ("System.Data.SQLite.dll", false),
                    (Path.Combine("x64", "SQLite.Interop.dll"), true)
#endif
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
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        candidatePaths.Add(path);
                    }
                }

                // 1. Check current assembly directory first (MOST LIKELY)
                var executingAssembly = Assembly.GetExecutingAssembly();
                var executionDirectory = Path.GetDirectoryName(executingAssembly.Location);
                if (!string.IsNullOrEmpty(executionDirectory))
                {
                    AddPath(Path.Combine(executionDirectory, fileName));
                    AddPath(Path.Combine(executionDirectory, "x64", fileName));
                }

                // 2. Check AppData Addins folder for current version
                try
                {
                    var revitVersion = Assembly.GetExecutingAssembly().FullName?.Contains("2023") == true ? "2023" :
                                      Assembly.GetExecutingAssembly().FullName?.Contains("2024") == true ? "2024" :
                                      Assembly.GetExecutingAssembly().FullName?.Contains("2025") == true ? "2025" :
                                      Assembly.GetExecutingAssembly().FullName?.Contains("2026") == true ? "2026" : null;

                    if (revitVersion != null)
                    {
                        var appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins", revitVersion);
                        AddPath(Path.Combine(appDataDir, fileName));
                        AddPath(Path.Combine(appDataDir, "x64", fileName));
                    }
                }
                catch { }

                string? fallback = null;
                foreach (var path in candidatePaths)
                {
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
