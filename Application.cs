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
        public override void OnStartup()
        {
            // ✅ DIAGNOSTIC SWITCH: Set diagnostic logging here
            // Change this value to true/false to enable/disable diagnostic logging
            // true  = Full diagnostic logging ON, Deployment mode OFF (SLOW - 9 zones/sec)
            // false = Full diagnostic logging OFF, Deployment mode ON (FAST - 30 zones/sec)
            // 
            // ⚠️ TO CHECK STATUS: Click "Diagnostic Status" button in Revit ribbon
            // ⚠️ TO TOGGLE: Click "Toggle Diagnostic" button in Revit ribbon
            MasterSwitch.DiagnosticLogging = true; // ⬅️ CHANGE THIS VALUE: true = ON (SLOW), false = OFF (FAST) - ✅ DIAGNOSTIC MODE ON for Bottom of Opening investigation
            
            MasterSwitch.DiagnosticLogging = true; // ⬅️ CHANGE THIS VALUE: true = ON (SLOW), false = OFF (FAST) - ✅ DIAGNOSTIC MODE ON for Bottom of Opening investigation
            
            // Store UIApp reference for event subscription
            // FIX: Use 'this.Application' property from ExternalApplication base class
            m_uiApp = this.Application;
            try
            {
                this.Application.ViewActivated += OnViewActivated;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to subscribe to ViewActivated: {ex.Message}");
            }
            
            // ✅ VERIFY SETTING: Double-check that the values were actually set
            System.Diagnostics.Debug.WriteLine($"[Application.OnStartup] ✅ AFTER MasterSwitch.DiagnosticLogging = true:");
            System.Diagnostics.Debug.WriteLine($"[Application.OnStartup]   - MasterSwitch.DiagnosticLogging = {MasterSwitch.DiagnosticLogging}");
            System.Diagnostics.Debug.WriteLine($"[Application.OnStartup]   - DeploymentConfiguration.DeploymentMode = {DeploymentConfiguration.DeploymentMode} (should be FALSE)");
            System.Diagnostics.Debug.WriteLine($"[Application.OnStartup]   - OptimizationFlags.UseDiagnosticMode = {OptimizationFlags.UseDiagnosticMode} (should be TRUE)");
            
            // ✅ LOG STARTUP STATUS: Always log diagnostic status at startup so you know the current state
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
                Directory.CreateDirectory(logDir);
                string startupLogPath = Path.Combine(logDir, "addin_startup.log");
                File.AppendAllText(startupLogPath, 
                    $"[{DateTime.Now}] 🔌 DIAGNOSTIC STATUS: DiagnosticLogging={(MasterSwitch.DiagnosticLogging ? "ON (SLOW)" : "OFF (FAST)")}, " +
                    $"DeploymentMode={DeploymentConfiguration.DeploymentMode}, " +
                    $"UseDiagnosticMode={OptimizationFlags.UseDiagnosticMode}\n");
            }
            catch { }
            // IMMEDIATE LOGGING - Create file as soon as add-in loads (Build: 2025-11-20 14:30)
            try
            {
                // ✅ CONFIGURATION: Use AppData\Roaming (consistent with SafeFileLogger)
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs");
                Directory.CreateDirectory(logDir);
                string startupLogPath = Path.Combine(logDir, "addin_startup.log");
                string r2023Dir = Path.Combine(logDir, "R2023");
                Directory.CreateDirectory(r2023Dir);
                string buildStampPath = Path.Combine(r2023Dir, "build_stamp.log");
                
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ========================================\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ADD-IN STARTUP BUILD 2025-11-20 14:30\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ========================================\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Assembly: {System.Reflection.Assembly.GetExecutingAssembly().Location}\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Process: {System.Diagnostics.Process.GetCurrentProcess().ProcessName}\n");

                // Write a hard build stamp with assembly info to R2023 folder
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                string asmLoc = asm.Location;
                DateTime asmWrite = File.Exists(asmLoc) ? File.GetLastWriteTime(asmLoc) : DateTime.MinValue;
                string asmVersion = asm.GetName().Version?.ToString() ?? "unknown";
                File.AppendAllText(buildStampPath, $"[{DateTime.Now:HH:mm:ss.fff}] *** STARTUP BUILD STAMP ***\n");
                File.AppendAllText(buildStampPath, $"AssemblyLocation={asmLoc}\n");
                File.AppendAllText(buildStampPath, $"AssemblyLastWrite={asmWrite:yyyy-MM-dd HH:mm:ss}\n");
                File.AppendAllText(buildStampPath, $"AssemblyVersion={asmVersion}\n");
                File.AppendAllText(buildStampPath, $"StartupPID={System.Diagnostics.Process.GetCurrentProcess().Id}\n\n");

                // Show a quick prompt with build timestamp on every run
                try
                {
                    var td = new TaskDialog("JSE Openings — Build Info");
                    td.MainInstruction = "Add-in Loaded";
                    td.MainContent =
                        $"Build Timestamp: {asmWrite:yyyy-MM-dd HH:mm:ss}\n" +
                        $"Version: {asmVersion}\n" +
                        $"Assembly: {asmLoc}";
                    td.CommonButtons = TaskDialogCommonButtons.Close;
                    td.Show();
                }
                catch { }
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
                
#if NET8_0_OR_GREATER
                // ✅ CRITICAL FOR REVIT 2025+ (.NET 8): Initialize SQLitePCL provider
                try
                {
                    File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] Initializing SQLitePCL batteries...\n");
                    SQLitePCL.Batteries.Init();
                    File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] SQLitePCL batteries initialized.\n");
                }
                catch (Exception initEx)
                {
                    File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] ❌ Error initializing SQLitePCL: {initEx.Message}\n");
                }
#endif

                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] About to CopyNativeSqliteDllToExecutionDirectory()\n");
                CopyNativeSqliteDllToExecutionDirectory();
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] CopyNativeSqliteDllToExecutionDirectory() DONE\n");

                // ✅ CRITICAL: Explicitly load SQLite native DLLs to satisfy providers
                // This is especially needed for Addin Manager where the search path might be incorrect.
                try
                {
                    var asmLoc = Assembly.GetExecutingAssembly().Location;
                    var asmDir = Path.GetDirectoryName(asmLoc);
                    if (!string.IsNullOrEmpty(asmDir))
                    {
                        // Check multiple potential locations
                        string[] candidatePaths = {
                            Path.Combine(asmDir, "SQLite.Interop.dll"),
                            Path.Combine(asmDir, "x64", "SQLite.Interop.dll"),
                            Path.Combine(asmDir, "e_sqlite3.dll") // Native DLL for net8.0 (Microsoft.Data.Sqlite)
                        };

                        foreach (var path in candidatePaths)
                        {
                            if (File.Exists(path))
                            {
                                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] Attempting LoadLibrary: {path}\n");
                                bool loaded = NativeLibraryLoader.LoadNativeLibrary(path);
                                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] LoadLibrary result: {loaded}\n");
                            }
                        }
                    }
                }
                catch (Exception loadEx)
                {
                    File.AppendAllText(startupLogPath, $"[{DateTime.Now}] [SQLite] ❌ Error during explicit LoadLibrary: {loadEx.Message}\n");
                }

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

            string tabName = "JSE_RevitAddin_MEP_OPENINGS";
            try
            {
                this.Application.CreateRibbonTab(tabName);
            }
            catch { /* Tab might already exist */ }

            RibbonPanel panel = this.Application.CreateRibbonPanel(tabName, "Commands");
            
            if (panel == null)
            {
                try { File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Panel creation FAILED\n"); } catch { }
                return;
            }

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

            try { File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Ribbon creation COMPLETED\n"); } catch { }
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
                    if (!string.IsNullOrEmpty(path))
                    {
                        candidatePaths.Add(path);
                    }
                }

                // Supported versions: 2023 - 2026
                var revitVersions = new[] { "2023", "2024", "2025", "2026" };
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
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2026\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2025\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2024\JSE_RevitAddin_MEP_OPENINGS.addin"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2023\JSE_RevitAddin_MEP_OPENINGS.addin")
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
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2026"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2025"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2024"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Autodesk\Revit\Addins\2023")
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
