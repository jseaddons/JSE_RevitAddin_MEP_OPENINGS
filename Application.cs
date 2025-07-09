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

            panel.AddPushButton<StartupCommand>("Execute")
                .SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            panel.AddPushButton<OpeningsPLaceCommand>("Openings")
                .SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

            panel.AddSeparator(); // This adds a visual gap

            panel.AddPushButton<MarkParameterAddValue>("Mark Parameters")
                .SetImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/JSE_RevitAddin_MEP_OPENINGS;component/Resources/Icons/RibbonIcon32.png");

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