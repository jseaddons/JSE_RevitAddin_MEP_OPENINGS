using Autodesk.Revit.Attributes;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using Nice3point.Revit.Toolkit.External;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class StartupCommand : ExternalCommand
    {
        public override void Execute()
        {
            DebugLogger.InitLogFile();
            DebugLogger.Log("JSE_RevitAddin_MEP_OPENINGS starting up...");

            // Ensure we're on the UI thread for WPF operations
            if (System.Windows.Application.Current == null)
            {
                // Create WPF Application if it doesn't exist (required for Revit add-ins)
                new System.Windows.Application();
            }

            var dialog = new MainDialog();
            DebugLogger.Log("Opening MainDialog window (modeless)");
            dialog.Show(); // Show modelessly to avoid blocking ExternalEvents
            DebugLogger.Log("MainDialog opened, command completed (modeless dialog)");
        }
    }
}
