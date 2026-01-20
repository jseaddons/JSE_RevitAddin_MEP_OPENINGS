using System;
using System.Windows;
using System.ComponentModel;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    public sealed partial class MainDialog : Window
    {
        private MainDialogViewModel? _viewModel;

        public MainDialog()
        {
            InitializeComponent();
            _viewModel = new MainDialogViewModel();
            DataContext = _viewModel;

            // Configure as modeless dialog (non-blocking)
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.ShowInTaskbar = false; // Don't show in taskbar as it's modeless
        }

        public MainDialog(MainDialogViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? new MainDialogViewModel();
            DataContext = _viewModel;

            // Configure as modeless dialog (non-blocking)
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.ShowInTaskbar = false; // Don't show in taskbar as it's modeless
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // Dispose ViewModel when dialog is closing
            if (_viewModel != null)
            {
                _viewModel.Dispose();
                _viewModel = null;
            }

            base.OnClosing(e);
        }
    }
}


