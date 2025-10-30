using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for tracking opening positions relative to MEP elements and updating them when MEP elements move
    /// </summary>
    public class OpeningTrackingService
    {
        private readonly Dictionary<ElementId, OpeningStatus> _openingStatuses;
        private readonly StatusManager _statusManager;
        private readonly object _lockObject = new object();

        public event EventHandler<OpeningStatusUpdatedEventArgs>? OpeningStatusUpdated;
        public event EventHandler<OpeningMovedEventArgs>? OpeningMoved;

        public OpeningTrackingService(StatusManager statusManager)
        {
            _openingStatuses = new Dictionary<ElementId, OpeningStatus>();
            _statusManager = statusManager;
        }

        /// <summary>
        /// Gets all tracked opening statuses
        /// </summary>
        public List<OpeningStatus> GetAllStatuses()
        {
            lock (_lockObject)
            {
                return _openingStatuses.Values.ToList();
            }
        }

        /// <summary>
        /// Gets a specific opening status by opening ID
        /// </summary>
        public OpeningStatus? GetStatus(ElementId openingId)
        {
            lock (_lockObject)
            {
                return _openingStatuses.TryGetValue(openingId, out var status) ? status : null;
            }
        }

        /// <summary>
        /// Registers an opening with its associated MEP element for tracking
        /// </summary>
        public void RegisterOpening(OpeningRequirement opening)
        {
            if (opening.ElementId == null || opening.WallId == null || opening.Location == null)
            {
                _statusManager.UpdateStatus("Cannot register opening: missing required data", StatusType.Warning);
                return;
            }

            try
            {
                // Get the actual opening element ID from the requirement
                var openingElementId = GetOpeningElementId(opening);
                if (openingElementId == null)
                {
                    _statusManager.UpdateStatus("Cannot find opening element in document", StatusType.Warning);
                    return;
                }

                var status = new OpeningStatus(
                    openingElementId,
                    opening.ElementId!,
                    opening.WallId!,
                    opening.Location!,
                    opening.MepType ?? "Unknown",
                    opening.Width,
                    opening.Height
                );

                lock (_lockObject)
                {
                    _openingStatuses[status.OpeningId] = status;
                }

                _statusManager.UpdateStatus($"Registered opening {status.OpeningId.IntegerValue} for {opening.MepType}", StatusType.Info);
                
                OpeningStatusUpdated?.Invoke(this, new OpeningStatusUpdatedEventArgs(status));
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error registering opening: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Registers an opening using existing opening element and MEP element
        /// </summary>
        public void RegisterOpening(ElementId openingId, ElementId mepElementId, ElementId wallId, Document doc)
        {
            try
            {
                var opening = doc.GetElement(openingId);
                var mepElement = doc.GetElement(mepElementId);
                
                if (opening == null || mepElement == null)
                {
                    _statusManager.UpdateStatus("Cannot register opening: element not found", StatusType.Warning);
                    return;
                }

                var location = GetElementLocation(opening);
                var mepType = GetMepElementType(mepElement);
                var dimensions = GetOpeningDimensions(opening);

                var status = new OpeningStatus(
                    openingId,
                    mepElementId,
                    wallId,
                    location,
                    mepType,
                    dimensions.Width,
                    dimensions.Height
                );

                lock (_lockObject)
                {
                    _openingStatuses[status.OpeningId] = status;
                }

                _statusManager.UpdateStatus($"Registered opening {status.OpeningId.IntegerValue} for {mepType}", StatusType.Info);
                
                OpeningStatusUpdated?.Invoke(this, new OpeningStatusUpdatedEventArgs(status));
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Error registering opening: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Updates opening positions when MEP elements move
        /// </summary>
        public void UpdateOpeningsForMovedMepElements(Document doc)
        {
            var updatedCount = 0;
            var errorCount = 0;

            var settings = ApplicationProfileService.Instance.GetCurrentSettings();

            lock (_lockObject)
            {
                foreach (var status in _openingStatuses.Values.Where(s => s.IsAutoUpdateEnabled))
                {
                    try
                    {
                        var mepElement = doc.GetElement(status.MepElementId);
                        if (mepElement == null)
                        {
                            status.MarkAsError("MEP element not found");
                            errorCount++;
                            continue;
                        }

                        var newLocation = GetMepElementLocation(mepElement);
                        if (newLocation == null)
                        {
                            status.MarkAsError("Cannot get MEP element location");
                            errorCount++;
                            continue;
                        }

                        var opening = doc.GetElement(status.OpeningId);
                        if (opening == null)
                        {
                            status.MarkAsError("Opening element not found");
                            errorCount++;
                            continue;
                        }

                        var newDimensions = GetOpeningDimensions(opening);

                        // Check if MEP element moved significantly
                        var locationDelta = status.CurrentLocation.DistanceTo(newLocation);
                        var dimensionDeltaWidth = Math.Abs(status.Width - newDimensions.Width);
                        var dimensionDeltaHeight = Math.Abs(status.Height - newDimensions.Height);

                        bool hasChanged = false;

                        if (locationDelta > settings.LocationChangeThreshold)
                        {
                            hasChanged = true;
                        }

                        if (dimensionDeltaWidth > settings.DimensionChangeThreshold || dimensionDeltaHeight > settings.DimensionChangeThreshold)
                        {
                            hasChanged = true;
                        }

                        if (hasChanged)
                        {
                            if (settings.ResetApprovalStatus)
                            {
                                status.ApprovalStatus = "Pending";
                            }

                            UpdateOpeningPosition(doc, status, newLocation);
                            status.UpdateLocation(newLocation);
                            status.Width = newDimensions.Width;
                            status.Height = newDimensions.Height;
                            updatedCount++;
                            
                            OpeningMoved?.Invoke(this, new OpeningMovedEventArgs(status, locationDelta));
                        }
                    }
                    catch (Exception ex)
                    {
                        status.MarkAsError(ex.Message);
                        errorCount++;
                        _statusManager.UpdateStatus($"Error updating opening {status.OpeningId}: {ex.Message}", StatusType.Error);
                    }
                }
            }

            if (updatedCount > 0)
            {
                _statusManager.UpdateStatus($"Updated {updatedCount} opening positions", StatusType.Success);
            }
            
            if (errorCount > 0)
            {
                _statusManager.UpdateStatus($"Encountered {errorCount} errors during update", StatusType.Warning);
            }
        }

        /// <summary>
        /// Updates a single opening position
        /// </summary>
        public void UpdateSingleOpening(ElementId openingId, Document doc)
        {
            lock (_lockObject)
            {
                if (!_openingStatuses.TryGetValue(openingId, out var status))
                {
                    _statusManager.UpdateStatus($"Opening {openingId.IntegerValue} not found in tracking", StatusType.Warning);
                    return;
                }

                try
                {
                    var mepElement = doc.GetElement(status.MepElementId);
                    if (mepElement == null)
                    {
                        status.MarkAsError("MEP element not found");
                        return;
                    }

                    var newLocation = GetMepElementLocation(mepElement);
                    if (newLocation == null)
                    {
                        status.MarkAsError("Cannot get MEP element location");
                        return;
                    }

                    UpdateOpeningPosition(doc, status, newLocation);
                    status.UpdateLocation(newLocation);
                    
                    _statusManager.UpdateStatus($"Updated opening {status.OpeningId.IntegerValue}", StatusType.Success);
                }
                catch (Exception ex)
                {
                    status.MarkAsError(ex.Message);
                    _statusManager.UpdateStatus($"Error updating opening {status.OpeningId}: {ex.Message}", StatusType.Error);
                }
            }
        }

        /// <summary>
        /// Removes an opening from tracking
        /// </summary>
        public void UnregisterOpening(ElementId openingId)
        {
            lock (_lockObject)
            {
                if (_openingStatuses.Remove(openingId))
                {
                    _statusManager.UpdateStatus($"Unregistered opening {openingId.IntegerValue}", StatusType.Info);
                }
            }
        }

        /// <summary>
        /// Clears all tracked openings
        /// </summary>
        public void ClearAll()
        {
            lock (_lockObject)
            {
                var count = _openingStatuses.Count;
                _openingStatuses.Clear();
                _statusManager.UpdateStatus($"Cleared {count} tracked openings", StatusType.Info);
            }
        }

        private void UpdateOpeningPosition(Document doc, OpeningStatus status, XYZ newMepLocation)
        {
            using (var transaction = new Transaction(doc, "Update Opening Position"))
            {
                transaction.Start();

                var opening = doc.GetElement(status.OpeningId);
                if (opening != null)
                {
                    // Calculate the movement vector
                    var movementVector = newMepLocation - status.CurrentLocation;
                    
                    // Move the opening to match the new MEP element position
                    ElementTransformUtils.MoveElement(doc, status.OpeningId, movementVector);
                }

                transaction.Commit();
            }
        }

        private ElementId? GetOpeningElementId(OpeningRequirement opening)
        {
            // This method needs to be implemented based on how you create openings
            // For now, we'll assume the opening requirement has the element ID
            // You may need to modify this based on your opening creation process
            return opening.ElementId;
        }

        private XYZ GetElementLocation(Element element)
        {
            var location = element.Location;
            
            if (location is LocationPoint point)
            {
                return point.Point;
            }
            else if (location is LocationCurve curve)
            {
                return curve.Curve.GetEndPoint(0);
            }
            
            return XYZ.Zero;
        }

        private XYZ? GetMepElementLocation(Element mepElement)
        {
            var location = mepElement.Location;
            
            if (location is LocationPoint point)
            {
                return point.Point;
            }
            else if (location is LocationCurve curve)
            {
                // For linear elements, use the start point
                return curve.Curve.GetEndPoint(0);
            }
            
            return null;
        }

        private string GetMepElementType(Element mepElement)
        {
            if (mepElement is MEPCurve)
            {
                var category = mepElement.Category?.Name ?? "Unknown";
                return category switch
                {
                    var name when name.Contains("Pipe") => "Pipe",
                    var name when name.Contains("Duct") => "Duct",
                    var name when name.Contains("Cable Tray") => "CableTray",
                    var name when name.Contains("Conduit") => "Conduit",
                    _ => category
                };
            }
            
            return mepElement.Category?.Name ?? "Unknown";
        }

        private (double Width, double Height) GetOpeningDimensions(Element opening)
        {
            try
            {
                var widthParam = opening.LookupParameter("Width");
                var heightParam = opening.LookupParameter("Height");
                
                var width = widthParam?.AsDouble() * 304.8 ?? 0; // Convert feet to mm
                var height = heightParam?.AsDouble() * 304.8 ?? 0; // Convert feet to mm
                
                return (width, height);
            }
            catch
            {
                return (0, 0);
            }
        }
    }

    /// <summary>
    /// Event arguments for opening status updates
    /// </summary>
    public class OpeningStatusUpdatedEventArgs : EventArgs
    {
        public OpeningStatus Status { get; }
        public DateTime Timestamp { get; }

        public OpeningStatusUpdatedEventArgs(OpeningStatus status)
        {
            Status = status;
            Timestamp = DateTime.Now;
        }
    }

    /// <summary>
    /// Event arguments for opening movement events
    /// </summary>
    public class OpeningMovedEventArgs : EventArgs
    {
        public OpeningStatus Status { get; }
        public double MovementDistance { get; }
        public DateTime Timestamp { get; }

        public OpeningMovedEventArgs(OpeningStatus status, double movementDistance)
        {
            Status = status;
            MovementDistance = movementDistance;
            Timestamp = DateTime.Now;
        }
    }
}
