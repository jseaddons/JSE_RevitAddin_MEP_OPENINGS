using System;
using Autodesk.Revit.DB;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    /// <summary>
    /// ViewModel for the opening status dialog
    /// </summary>
    public partial class OpeningStatusDialogViewModel : ObservableObject
    {
        private readonly OpeningTrackingService _trackingService;
        private readonly LinkedFileReloadService _linkedFileService;
        private readonly StatusManager _statusManager;
        private Document? _document;

        public OpeningStatusPanelViewModel StatusPanelViewModel { get; }

        [ObservableProperty]
        private string _autoUpdateStatus = "Enabled";

        [ObservableProperty]
        private string _lastCheckTime = "Never";

        [ObservableProperty]
        private bool _isCheckingForUpdates;

        public IRelayCommand RefreshCommand { get; }
        public IRelayCommand CheckForUpdatesCommand { get; }
        public IRelayCommand ForceCheckCommand { get; }

        public OpeningStatusDialogViewModel(OpeningTrackingService trackingService, 
                                          LinkedFileReloadService linkedFileService,
                                          StatusManager statusManager)
        {
            _trackingService = trackingService;
            _linkedFileService = linkedFileService;
            _statusManager = statusManager;

            StatusPanelViewModel = new OpeningStatusPanelViewModel(trackingService, statusManager);

            RefreshCommand = new RelayCommand(OnRefresh);
            CheckForUpdatesCommand = new RelayCommand(OnCheckForUpdates, CanCheckForUpdates);
            ForceCheckCommand = new RelayCommand(OnForceCheck, CanCheckForUpdates);

            // Subscribe to events
            _linkedFileService.LinkedFileReloaded += OnLinkedFileReloaded;
            _statusManager.StatusUpdated += OnStatusUpdated;

            // Initialize
            UpdateLastCheckTime();
        }

        /// <summary>
        /// Sets the document for the dialog
        /// </summary>
        public void SetDocument(Document document)
        {
            _document = document;
            StatusPanelViewModel.SetDocument(document);
        }

        private void OnRefresh()
        {
            StatusPanelViewModel.RefreshCommand.Execute(null);
            UpdateLastCheckTime();
        }

        private void OnCheckForUpdates()
        {
            if (_document == null) return;

            try
            {
                IsCheckingForUpdates = true;
                _statusManager.UpdateStatus("Checking for MEP element movements...", StatusType.Processing);
                
                _trackingService.UpdateOpeningsForMovedMepElements(_document);
                _linkedFileService.CheckForLinkedFileUpdates(_document);
                
                _statusManager.UpdateStatus("Opening position check completed", StatusType.Success);
                UpdateLastCheckTime();
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error checking updates: {ex.Message}", StatusType.Error);
            }
            finally
            {
                IsCheckingForUpdates = false;
            }
        }

        private void OnForceCheck()
        {
            if (_document == null) return;

            try
            {
                IsCheckingForUpdates = true;
                _statusManager.UpdateStatus("Force checking for updates...", StatusType.Processing);
                
                _linkedFileService.ForceCheckForUpdates(_document);
                _trackingService.UpdateOpeningsForMovedMepElements(_document);
                
                _statusManager.UpdateStatus("Force check completed", StatusType.Success);
                UpdateLastCheckTime();
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error in force check: {ex.Message}", StatusType.Error);
            }
            finally
            {
                IsCheckingForUpdates = false;
            }
        }

        private bool CanCheckForUpdates()
        {
            return !IsCheckingForUpdates && _document != null;
        }

        private void OnLinkedFileReloaded(object? sender, LinkedFileReloadedEventArgs e)
        {
            // Update the UI on the main thread
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                StatusPanelViewModel.RefreshCommand.Execute(null);
                UpdateLastCheckTime();
            });
        }

        private void OnStatusUpdated(object? sender, StatusUpdateEventArgs e)
        {
            // Update the UI on the main thread
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                // You could update a status message here if needed
            });
        }

        private void UpdateLastCheckTime()
        {
            LastCheckTime = DateTime.Now.ToString("HH:mm:ss");
        }

        /// <summary>
        /// Gets the current status summary
        /// </summary>
        public string GetStatusSummary()
        {
            return StatusPanelViewModel.GetStatusSummary();
        }

        /// <summary>
        /// Checks if there are any issues that need attention
        /// </summary>
        public bool HasIssues()
        {
            return StatusPanelViewModel.HasOpeningsNeedingAttention();
        }

        /// <summary>
        /// Gets the overall status color
        /// </summary>
        public string GetOverallStatusColor()
        {
            return StatusPanelViewModel.GetStatusColor();
        }
    }
}
