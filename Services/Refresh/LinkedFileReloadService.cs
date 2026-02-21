using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for detecting linked file reloads and triggering opening position updates
    /// </summary>
    public class LinkedFileReloadService
    {
        private readonly OpeningTrackingService _trackingService;
        private readonly StatusManager _statusManager;
        private Document? _lastDocument;
        private DateTime _lastCheckTime;
        private readonly Dictionary<string, DateTime> _linkedFileLastModified;

        public event EventHandler<LinkedFileReloadedEventArgs>? LinkedFileReloaded;

        public LinkedFileReloadService(OpeningTrackingService trackingService, StatusManager statusManager)
        {
            _trackingService = trackingService;
            _statusManager = statusManager;
            _lastCheckTime = DateTime.Now;
            _linkedFileLastModified = new Dictionary<string, DateTime>();
        }

        /// <summary>
        /// Checks for linked file reloads and updates openings if needed
        /// </summary>
        public void CheckForLinkedFileUpdates(Document doc)
        {
            if (doc == null) return;

            try
            {
                // Check if this is a different document or if enough time has passed
                if (_lastDocument != doc || DateTime.Now - _lastCheckTime > TimeSpan.FromMinutes(1))
                {
                    _lastDocument = doc;
                    _lastCheckTime = DateTime.Now;

                    // Check for linked file changes
                    var linkedFilesChanged = CheckLinkedFilesChanged(doc);

                    if (linkedFilesChanged)
                    {
                        _statusManager.UpdateStatus("Linked files reloaded - updating opening positions", StatusType.Info);
                        _trackingService.UpdateOpeningsForMovedMepElements(doc);
                        
                        LinkedFileReloaded?.Invoke(this, new LinkedFileReloadedEventArgs(doc));
                    }
                }
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error checking linked files: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Forces a check for linked file updates regardless of timing
        /// </summary>
        public void ForceCheckForUpdates(Document doc)
        {
            if (doc == null) return;

            try
            {
                _statusManager.UpdateStatus("Force checking for linked file updates...", StatusType.Processing);
                
                var linkedFilesChanged = CheckLinkedFilesChanged(doc);
                
                if (linkedFilesChanged)
                {
                    _statusManager.UpdateStatus("Linked files changed - updating opening positions", StatusType.Info);
                    _trackingService.UpdateOpeningsForMovedMepElements(doc);
                    
                    LinkedFileReloaded?.Invoke(this, new LinkedFileReloadedEventArgs(doc));
                }
                else
                {
                    _statusManager.UpdateStatus("No linked file changes detected", StatusType.Info);
                }
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error force checking linked files: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Gets information about all linked files in the document
        /// </summary>
        public List<LinkedFileReloadInfo> GetLinkedFileInfo(Document doc)
        {
            var linkedFiles = new List<LinkedFileReloadInfo>();

            try
            {
                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var link in linkInstances)
                {
                    var linkDoc = link.GetLinkDocument();
                    var info = new LinkedFileReloadInfo
                    {
                        LinkInstanceId = link.Id,
                        LinkName = link.Name,
                        IsLoaded = linkDoc != null,
                        LastModified = linkDoc != null ? GetDocumentLastModified(linkDoc) : DateTime.MinValue,
                        FilePath = GetLinkFilePath(link)
                    };
                    
                    linkedFiles.Add(info);
                }
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error getting linked file info: {ex.Message}", StatusType.Warning);
            }

            return linkedFiles;
        }

        private bool CheckLinkedFilesChanged(Document doc)
        {
            try
            {
                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                bool hasChanges = false;

                foreach (var link in linkInstances)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        var linkKey = $"{doc.Title}_{link.Id.GetIntegerValue()}";
                        var currentModified = GetDocumentLastModified(linkDoc);
                        
                        if (_linkedFileLastModified.TryGetValue(linkKey, out var lastModified))
                        {
                            if (currentModified > lastModified)
                            {
                                hasChanges = true;
                                _statusManager.UpdateStatus($"Linked file '{link.Name}' has been updated", StatusType.Info);
                            }
                        }
                        else
                        {
                            // First time seeing this link
                            hasChanges = true;
                        }
                        
                        _linkedFileLastModified[linkKey] = currentModified;
                    }
                }

                return hasChanges;
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error checking linked files: {ex.Message}", StatusType.Warning);
                return false;
            }
        }

        private DateTime GetDocumentLastModified(Document doc)
        {
            try
            {
                // Try to get the document's last modified time
                // This is a simplified approach - you might need to implement
                // a more sophisticated way to detect document changes
                return DateTime.Now;
            }
            catch
            {
                return DateTime.Now;
            }
        }

        private string GetLinkFilePath(RevitLinkInstance link)
        {
            try
            {
                // Get the file path of the linked document
                var linkDoc = link.GetLinkDocument();
                return linkDoc?.PathName ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// Resets the tracking state (useful for testing or manual resets)
        /// </summary>
        public void ResetTracking()
        {
            _linkedFileLastModified.Clear();
            _lastDocument = null;
            _lastCheckTime = DateTime.Now;
            _statusManager.UpdateStatus("Linked file tracking reset", StatusType.Info);
        }
    }

    /// <summary>
    /// Information about a linked file for reload tracking
    /// </summary>
    public class LinkedFileReloadInfo
    {
        public ElementId LinkInstanceId { get; set; }
        public string LinkName { get; set; } = string.Empty;
        public bool IsLoaded { get; set; }
        public DateTime LastModified { get; set; }
        public string FilePath { get; set; } = string.Empty;
    }

    /// <summary>
    /// Event arguments for linked file reload events
    /// </summary>
    public class LinkedFileReloadedEventArgs : EventArgs
    {
        public Document Document { get; }
        public DateTime Timestamp { get; }

        public LinkedFileReloadedEventArgs(Document document)
        {
            Document = document;
            Timestamp = DateTime.Now;
        }
    }
}
