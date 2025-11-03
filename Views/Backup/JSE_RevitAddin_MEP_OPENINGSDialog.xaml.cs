using System.Windows;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    public partial class JSE_RevitAddin_MEP_OPENINGSDialog : Window
    {
        public JSE_RevitAddin_MEP_OPENINGSDialog(JSE_RevitAddin_MEP_OPENINGSViewModel viewModel)
        {
            DataContext = viewModel;
            // Host the user control in the window
            Content = new JSE_RevitAddin_MEP_OPENINGSView(viewModel);
        }
    }
}
