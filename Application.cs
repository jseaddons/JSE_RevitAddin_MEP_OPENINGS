using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Security;
using Nice3point.Revit.Toolkit.External;
using Serilog;
using Serilog.Events;

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
            var panel = Application.CreatePanel("Commands", "JSE_RevitAddin_MEP_OPENINGS");
            if (panel == null)
            {
                // Panel creation failed for some environment; skip ribbon creation to avoid null references.
                return;
            }

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

            // Test Profile Management Command
            var button5 = panel.AddPushButton<TestProfileManagementCommand>("Profile Manager");
            button5.SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png");
            button5.SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

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