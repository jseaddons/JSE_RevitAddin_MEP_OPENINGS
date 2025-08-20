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
            // Validate license before startup - Simple JSE domain check
            if (!LicenseValidator.ValidateLicense() || !LicenseValidator.ValidateHardwareId())
            {
                TaskDialog.Show("License Error",
                    "This add-in is licensed for JSE domain computers only. " +
                    $"Current domain: {Environment.UserDomainName}");
                return;
            }

            CreateLogger();
            CreateRibbon();
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