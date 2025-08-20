using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Security;
using Nice3point.Revit.Toolkit.External;

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
            // Silent domain validation - user should not be aware of this check
            var domainValidation = DomainValidator.ValidateDomain();
            if (!domainValidation.IsValid)
            {
                // Silently fail - no error dialogs, no user notification
                // Just return without loading the add-in
                return;
            }

            // Domain validation successful - no logging needed
            CreateLogger();
            CreateRibbon();
        }

        public override void OnShutdown()
        {
            // Logging disabled - no cleanup needed
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

            // Deleted commands removed from ribbon: DeletePipeSleeveCommand, GetSleeveSummaryCommand
        }


        private static void CreateLogger()
        {
            // Logging disabled - no logger needed
            // AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            // {
            //     // Exception logging disabled
            // };
        }
    }
}