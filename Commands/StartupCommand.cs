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
            // Initialize the debug log file
            DebugLogger.InitLogFile();
            DebugLogger.Log("JSE_RevitAddin_MEP_OPENINGS starting up...");

            var viewModel = new JSE_RevitAddin_MEP_OPENINGSViewModel(ExternalCommandData);
            var dialog = new JSE_RevitAddin_MEP_OPENINGS.Views.JSE_RevitAddin_MEP_OPENINGSDialog(viewModel);
            DebugLogger.Log("Opening main dialog window");
            dialog.ShowDialog();
            DebugLogger.Log("Dialog closed, command completed");
        }
    }
}
