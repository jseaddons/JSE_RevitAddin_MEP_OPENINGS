using System;
using System.Windows;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Interaction logic for ProfileManagementDialog.xaml
    /// </summary>
    public partial class ProfileManagementDialog : Window
    {
        private readonly ProfileManagementViewModel _viewModel;

        public ProfileManagementDialog(ProfileManagementViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            // Subscribe to events
            _viewModel.DialogClosed += OnDialogClosed;
        }

        private void OnDialogClosed(object? sender, EventArgs e)
        {
            DialogResult = true;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Unsubscribe from events
            if (_viewModel != null)
            {
                _viewModel.DialogClosed -= OnDialogClosed;
            }

            base.OnClosed(e);
        }
    }
}
