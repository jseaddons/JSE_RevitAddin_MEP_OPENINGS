using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    public sealed partial class JSE_RevitAddin_MEP_OPENINGSView
    {
        public JSE_RevitAddin_MEP_OPENINGSView(JSE_RevitAddin_MEP_OPENINGSViewModel viewModel)
        {
            DataContext = viewModel;
            InitializeComponent();
        }
    }
}
