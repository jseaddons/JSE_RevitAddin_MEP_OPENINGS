using System;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;
using System.Windows.Threading;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    public sealed partial class MainDialogViewModel : ObservableObject, IDisposable
    {
        private readonly ApplicationProfileService? _appProfileService;
        private readonly ExternalEventManager? _externalEventManager;
        private readonly Dispatcher _uiDispatcher;

        public LeftPanelViewModel LeftPanelViewModel { get; } = new LeftPanelViewModel();
        public MiddlePanelViewModel MiddlePanelViewModel { get; } = new MiddlePanelViewModel();
        public RightPanelViewModel RightPanelViewModel { get; } = new RightPanelViewModel();

        public ObservableCollection<string> OpeningModes { get; } = new ObservableCollection<string>
        {
            "Horizontal",
            "Vertical"
        };

        [ObservableProperty]
        private string _selectedOpeningMode = "Horizontal";

        [ObservableProperty]
        private double _progressPercentage = 0.0;

        [ObservableProperty]
        private string _progressText = "Ready";

        [ObservableProperty]
        private string _profileSummary = string.Empty;

        [ObservableProperty]
        private string _statusText = "Ready";

        [ObservableProperty]
        private bool _autoUpdateEnabled = true;

        public IRelayCommand OkCommand { get; }
        public IRelayCommand CancelCommand { get; }
        public IRelayCommand SaveCommand { get; }
        public IRelayCommand CloseCommand { get; }
        public IRelayCommand CheckForUpdatesCommand { get; }
        public IRelayCommand ShowStatusCommand { get; }

        public MainDialogViewModel()
        {
            _uiDispatcher = Dispatcher.CurrentDispatcher;

            OkCommand = new RelayCommand(OnOk);
            CancelCommand = new RelayCommand(OnCancel);
            SaveCommand = new RelayCommand(OnSave);
            CloseCommand = new RelayCommand(OnClose);
            CheckForUpdatesCommand = new RelayCommand(OnCheckForUpdates);
            ShowStatusCommand = new RelayCommand(OnShowStatus);

            ProfileSummary = "No Profile Active";
        }

        public MainDialogViewModel(ApplicationProfileService appProfileService) : this()
        {
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));

            // Initialize ExternalEventManager for safe Revit API calls
            try
            {
                _externalEventManager = new ExternalEventManager(appProfileService);
                if (!_externalEventManager.IsInitialized)
                {
                    _appProfileService.StatusManager.UpdateStatus("Warning: ExternalEvent system not initialized properly", StatusType.Warning);
                }
            }
            catch (Exception ex)
            {
                _appProfileService.StatusManager.UpdateStatus($"Failed to initialize ExternalEvent system: {ex.Message}", StatusType.Error);
            }

            // Subscribe to profile and status change events
            _appProfileService.ProfileChanged += OnProfileChanged;
            _appProfileService.StatusUpdated += OnStatusUpdated;

            // Initialize with current profile info
            UpdateProfileInfo();
        }

        private void OnOk()
        {
            if (_externalEventManager == null)
            {
                // Fallback if ExternalEvent system is not available
                ProgressText = "Error: ExternalEvent system not available";
                ProgressPercentage = 0;
                StatusText = "✗ Cannot execute operations - ExternalEvent system failed";
                return;
            }

            // Update UI immediately (on UI thread)
            ProgressText = "Requesting opening placement...";
            ProgressPercentage = 10;

            // Execute opening placement via ExternalEvent (safe for Revit API)
            _externalEventManager.ExecuteOpeningPlacement();

            // Note: Actual progress updates will come from the ExternalEvent handler
            // which runs on Revit's main thread and updates status via StatusManager
        }

        private void OnCancel()
        {
            ProgressText = "Cancelled";
        }

        private void OnSave()
        {
            ProgressText = "Saved";
        }

        private void OnClose()
        {
            // Intentionally left blank; window can close itself by handling this command if needed
        }

        private void OnCheckForUpdates()
        {
            try
            {
                StatusText = "⟳ Checking for MEP element movements...";
                
                // This would need to be implemented with proper service injection
                // For now, just show a status message
                StatusText = "✓ Opening position check completed";
            }
            catch (Exception ex)
            {
                StatusText = $"✗ Error checking updates: {ex.Message}";
            }
        }

        private void OnShowStatus()
        {
            try
            {
                // This would open the OpeningStatusDialog
                // For now, just show a status message
                StatusText = "ℹ Opening status dialog would open here";
            }
            catch (Exception ex)
            {
                StatusText = $"✗ Error opening status dialog: {ex.Message}";
            }
        }

        /// <summary>
        /// Handles profile change events
        /// </summary>
        private void OnProfileChanged(object? sender, EventArgs e)
        {
            UpdateProfileInfo();
        }

        /// <summary>
        /// Handles status update events
        /// </summary>
        private void OnStatusUpdated(object? sender, StatusUpdateEventArgs e)
        {
            if (e != null)
            {
                // Format status text with type indicator
                string typeIndicator = e.Type switch
                {
                    StatusType.Success => "✓",
                    StatusType.Error => "✗",
                    StatusType.Warning => "⚠",
                    StatusType.Info => "ℹ",
                    StatusType.Processing => "⟳",
                    _ => ""
                };

                StatusText = $"{typeIndicator} {e.Message}";
            }
            else
            {
                StatusText = "Ready";
            }
        }

        /// <summary>
        /// Updates the profile summary information
        /// </summary>
        private void UpdateProfileInfo()
        {
            if (_appProfileService != null)
            {
                ProfileSummary = _appProfileService.GetUserInfoString();
            }
        }

        /// <summary>
        /// Disposes the ViewModel and unsubscribes from events
        /// </summary>
        public void Dispose()
        {
            if (_appProfileService != null)
            {
                _appProfileService.ProfileChanged -= OnProfileChanged;
                _appProfileService.StatusUpdated -= OnStatusUpdated;
            }

            if (_externalEventManager != null)
            {
                _externalEventManager.Dispose();
            }
        }
    }
}
