using System;
using System.Threading;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Manages ExternalEvents for safe Revit API execution from WPF UI
    /// Addresses Building Coder issues: Modal Dialog Blocking, UI Thread Management, ExternalEvent Timing
    /// </summary>
    public class ExternalEventManager : IDisposable
    {
        private readonly ApplicationProfileService _profileService;

        // External event handlers (singleton pattern to avoid multiple instances)
        private OpeningPlacementExternalEventHandler? _openingPlacementHandler;
        private ParameterSyncExternalEventHandler? _parameterSyncHandler;
        private StatusUpdateExternalEventHandler? _statusUpdateHandler;

        // External events
        private ExternalEvent? _openingPlacementEvent;
        private ExternalEvent? _parameterSyncEvent;
        private ExternalEvent? _statusUpdateEvent;

        // Thread safety
        private readonly object _lock = new object();

        public ExternalEventManager(ApplicationProfileService profileService)
        {
            _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
            InitializeExternalEvents();
        }

        /// <summary>
        /// Initialize all external events on startup
        /// </summary>
        private void InitializeExternalEvents()
        {
            lock (_lock)
            {
                // Create handlers
                _openingPlacementHandler = new OpeningPlacementExternalEventHandler(_profileService);
                _parameterSyncHandler = new ParameterSyncExternalEventHandler(_profileService);
                _statusUpdateHandler = new StatusUpdateExternalEventHandler(_profileService);

                // Create external events
                _openingPlacementEvent = ExternalEvent.Create(_openingPlacementHandler);
                _parameterSyncEvent = ExternalEvent.Create(_parameterSyncHandler);
                _statusUpdateEvent = ExternalEvent.Create(_statusUpdateHandler);
            }
        }

        /// <summary>
        /// Execute opening placement operation via ExternalEvent
        /// </summary>
        public void ExecuteOpeningPlacement()
        {
            if (_openingPlacementEvent != null)
            {
                // Update UI status immediately (on UI thread)
                _profileService.StatusManager.UpdateStatus("Requesting MEP opening placement...", StatusType.Info);

                // Raise external event (will execute on Revit thread)
                var result = _openingPlacementEvent.Raise();

                if (result == ExternalEventRequest.Accepted)
                {
                    // Event accepted, will execute asynchronously
                    _profileService.StatusManager.UpdateStatus("Opening placement request queued", StatusType.Processing);
                }
                else
                {
                    _profileService.StatusManager.UpdateStatus("Failed to queue opening placement", StatusType.Warning);
                }
            }
        }

        /// <summary>
        /// Execute parameter synchronization via ExternalEvent
        /// </summary>
        public void ExecuteParameterSync()
        {
            if (_parameterSyncEvent != null)
            {
                _profileService.StatusManager.UpdateStatus("Requesting parameter synchronization...", StatusType.Info);

                var result = _parameterSyncEvent.Raise();

                if (result == ExternalEventRequest.Accepted)
                {
                    _profileService.StatusManager.UpdateStatus("Parameter sync request queued", StatusType.Processing);
                }
                else
                {
                    _profileService.StatusManager.UpdateStatus("Failed to queue parameter sync", StatusType.Warning);
                }
            }
        }

        /// <summary>
        /// Update status via ExternalEvent (for thread safety)
        /// </summary>
        public void UpdateStatusViaEvent(string message, Models.StatusType statusType)
        {
            if (_statusUpdateHandler != null && _statusUpdateEvent != null)
            {
                _statusUpdateHandler.UpdateStatus(message, statusType);

                var result = _statusUpdateEvent.Raise();

                if (result != ExternalEventRequest.Accepted)
                {
                    // Fallback to direct update if ExternalEvent fails
                    _profileService.StatusManager.UpdateStatus(message, statusType);
                }
            }
        }

        /// <summary>
        /// Check if external events are properly initialized
        /// </summary>
        public bool IsInitialized => _openingPlacementEvent != null && _parameterSyncEvent != null && _statusUpdateEvent != null;

        /// <summary>
        /// Dispose of external events and handlers
        /// </summary>
        public void Dispose()
        {
            lock (_lock)
            {
                _openingPlacementEvent?.Dispose();
                _parameterSyncEvent?.Dispose();
                _statusUpdateEvent?.Dispose();

                _openingPlacementEvent = null;
                _parameterSyncEvent = null;
                _statusUpdateEvent = null;

                _openingPlacementHandler = null;
                _parameterSyncHandler = null;
                _statusUpdateHandler = null;
            }
        }
    }
}
