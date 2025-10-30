using System;
using System.Windows;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.ViewModels;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Interaction logic for ProfileSetupDialog.xaml
    /// </summary>
    public partial class ProfileSetupDialog : Window
    {
        private readonly object _viewModel; // Use object to support both ViewModels

        public ProfileSetupDialog(ProfileSetupViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            // Subscribe to events
            viewModel.ProfileCreated += OnProfileCreated;
            viewModel.DialogClosed += OnDialogClosed;
        }

        public ProfileSetupDialog(ProfileSetupViewModelSafe viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            // Subscribe to events
            viewModel.ProfileCreated += OnProfileCreated;
            viewModel.DialogClosed += OnDialogClosed;
        }

        /// <summary>
        /// Gets the created profile
        /// </summary>
        public UserProfile? CreatedProfile { get; private set; }

        /// <summary>
        /// Gets whether the dialog was completed successfully
        /// </summary>
        public bool IsCompleted { get; private set; }

        private void OnProfileCreated(object? sender, ProfileCreatedEventArgs e)
        {
            CreatedProfile = e.Profile;
            IsCompleted = true;
            DialogResult = true;
            Close();
        }

        private void OnDialogClosed(object? sender, EventArgs e)
        {
            IsCompleted = false;
            DialogResult = false;
            Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Unsubscribe from events based on ViewModel type
            if (_viewModel is ProfileSetupViewModel originalVM)
            {
                originalVM.ProfileCreated -= OnProfileCreated;
                originalVM.DialogClosed -= OnDialogClosed;
                originalVM.Dispose();
            }
            else if (_viewModel is ProfileSetupViewModelSafe safeVM)
            {
                safeVM.ProfileCreated -= OnProfileCreated;
                safeVM.DialogClosed -= OnDialogClosed;
                safeVM.Dispose();
            }

            base.OnClosed(e);
        }
    }
}
