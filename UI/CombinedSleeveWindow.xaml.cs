using System;
using System.Windows;

namespace JSE_RevitAddin_MEP_OPENINGS.UI
{
    public partial class CombinedSleeveWindow : Window
    {
        public CombinedSleeveWindow(CombinedSleeveViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            
            // Wire up visibility actions
            viewModel.HideRequest = () => this.Hide();
            viewModel.ShowRequest = () => this.Show();
        }
    }
}
