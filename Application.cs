using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Security;
using Nice3point.Revit.Toolkit.External;
using Serilog;
using Serilog.Events;
using System.IO;

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
            // IMMEDIATE LOGGING - Create file as soon as add-in loads
            try
            {
                string startupLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\addin_startup.log";
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] ADD-IN STARTUP: OnStartup() called\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Assembly: {System.Reflection.Assembly.GetExecutingAssembly().Location}\n");
                File.AppendAllText(startupLogPath, $"[{DateTime.Now}] Process: {System.Diagnostics.Process.GetCurrentProcess().ProcessName}\n");
            }
            catch (Exception startupEx)
            {
                // Try alternative location if main log fails
                try
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\temp\addin_startup.log", $"[{DateTime.Now}] ADD-IN STARTUP FAILED: {startupEx.Message}\n");
                }
                catch { }
            }

            // Initialize logging first so we can capture any startup failures
            try
            {
                CreateLogger();

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
            // Log ribbon creation start
            try
            {
                string ribbonLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\ribbon_creation.log";
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] CreateRibbon() started\n");
            }
            catch { }

            var panel = Application.CreatePanel("Commands", "JSE_RevitAddin_MEP_OPENINGS");
            if (panel == null)
            {
                // Panel creation failed for some environment; skip ribbon creation to avoid null references.
                try
                {
                    string ribbonLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\ribbon_creation.log";
                    File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Panel creation FAILED - null panel returned\n");
                }
                catch { }
                return;
            }

            try
            {
                string ribbonLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\ribbon_creation.log";
                File.AppendAllText(ribbonLogPath, $"[{DateTime.Now}] Panel created successfully\n");
            }
            catch { }

            var button1 = panel.AddPushButton<StartupCommand>("Execute");
            button1.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button1.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            var button2 = panel.AddPushButton<OpeningsPLaceCommand>("Openings");
            button2.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button2.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            var button3 = panel.AddPushButton<MarkParameterAddValue>("Mark Parameters");
            button3.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button3.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            // ClusterMergeCommand removed from ribbon (command deprecated)

            panel.AddSeparator(); // This adds a visual gap

            // Test Profile System Command
            var button4 = panel.AddPushButton<TestProfileSystemCommand>("Test Profile");
            button4.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button4.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            // TEST: Profile Management Command
            var button5 = panel.AddPushButton<TestProfileManagementCommand>("Profile Manager");
            button5.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button5.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            try
            {
                string ribbonLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\ribbon_creation.log";
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
    }
}
