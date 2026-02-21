using System;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    /// <summary>
    /// ViewModel for the opening status panel
    /// </summary>
    public partial class OpeningStatusPanelViewModel : ObservableObject
    {
        private readonly OpeningTrackingService _trackingService;
        private readonly StatusManager _statusManager;
        private Document? _document;

        public ObservableCollection<OpeningStatusItem> OpeningStatuses { get; } = new();

        [ObservableProperty]
        private int _totalOpenings;

        [ObservableProperty]
        private int _activeOpenings;

        [ObservableProperty]
        private int _movedOpenings;

        [ObservableProperty]
        private int _errorOpenings;

        [ObservableProperty]
        private bool _isRefreshing;

        public IRelayCommand UpdateOpeningCommand { get; }
        public IRelayCommand RefreshCommand { get; }
        public IRelayCommand ClearAllCommand { get; }

        public OpeningStatusPanelViewModel(OpeningTrackingService trackingService, StatusManager statusManager)
        {
            _trackingService = trackingService;
            _statusManager = statusManager;

            UpdateOpeningCommand = new RelayCommand<OpeningStatusItem>(OnUpdateOpening);
            RefreshCommand = new RelayCommand(OnRefresh, CanRefresh);
            ClearAllCommand = new RelayCommand(OnClearAll);

            // Subscribe to status updates
            _trackingService.OpeningStatusUpdated += OnOpeningStatusUpdated;
            _trackingService.OpeningMoved += OnOpeningMoved;
        }

        /// <summary>
        /// Sets the document for getting additional element information
        /// </summary>
        public void SetDocument(Document document)
        {
            _document = document;
            RefreshStatusList();
        }

        private void OnRefresh()
        {
            RefreshStatusList();
        }

        private bool CanRefresh()
        {
            return !IsRefreshing;
        }

        private void OnUpdateOpening(OpeningStatusItem? item)
        {
            if (item == null || _document == null) return;

            try
            {
                _statusManager.UpdateStatus($"Updating opening {item.OpeningId.GetIntegerValue()}...", StatusType.Processing);
                _trackingService.UpdateSingleOpening(item.OpeningId, _document);
                _statusManager.UpdateStatus($"Updated opening {item.OpeningId.GetIntegerValue()}", StatusType.Success);
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error updating opening: {ex.Message}", StatusType.Error);
            }
        }

        private void OnClearAll()
        {
            try
            {
                _trackingService.ClearAll();
                RefreshStatusList();
                _statusManager.UpdateStatus("Cleared all tracked openings", StatusType.Info);
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error clearing openings: {ex.Message}", StatusType.Error);
            }
        }

        private void RefreshStatusList()
        {
            try
            {
                IsRefreshing = true;
                OpeningStatuses.Clear();

                var statuses = _trackingService.GetAllStatuses();
                foreach (var status in statuses)
                {
                    var item = new OpeningStatusItem(status, _document);
                    OpeningStatuses.Add(item);
                }

                UpdateCounts();
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error refreshing status list: {ex.Message}", StatusType.Error);
            }
            finally
            {
                IsRefreshing = false;
            }
        }

        private void UpdateCounts()
        {
            TotalOpenings = OpeningStatuses.Count;
            ActiveOpenings = OpeningStatuses.Count(s => s.Status == "Active");
            MovedOpenings = OpeningStatuses.Count(s => s.Status == "Moved");
            ErrorOpenings = OpeningStatuses.Count(s => s.Status == "Error");
        }

        private void OnOpeningStatusUpdated(object? sender, OpeningStatusUpdatedEventArgs e)
        {
            // Update the UI on the main thread
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                RefreshStatusList();
            });
        }

        private void OnOpeningMoved(object? sender, OpeningMovedEventArgs e)
        {
            // Update the UI on the main thread
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                RefreshStatusList();
            });
        }

        /// <summary>
        /// Gets a summary of the current status
        /// </summary>
        public string GetStatusSummary()
        {
            if (TotalOpenings == 0)
                return "No openings tracked";

            return $"Total: {TotalOpenings}, Active: {ActiveOpenings}, Moved: {MovedOpenings}, Errors: {ErrorOpenings}";
        }

        /// <summary>
        /// Checks if there are any openings that need attention
        /// </summary>
        public bool HasOpeningsNeedingAttention()
        {
            return MovedOpenings > 0 || ErrorOpenings > 0;
        }

        /// <summary>
        /// Gets the status color based on current state
        /// </summary>
        public string GetStatusColor()
        {
            if (ErrorOpenings > 0) return "Red";
            if (MovedOpenings > 0) return "Orange";
            if (ActiveOpenings > 0) return "Green";
            return "Gray";
        }
    }
}
