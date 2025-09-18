using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the status and tracking information for an opening relative to its MEP element
    /// </summary>
    public class OpeningStatus : INotifyPropertyChanged
    {
        private ElementId _openingId;
        private ElementId _mepElementId;
        private ElementId _wallId;
        private XYZ _originalLocation;
        private XYZ _currentLocation;
        private string _mepType = string.Empty;
        private DateTime _lastUpdated;
        private bool _isAutoUpdateEnabled = true;
        private string _status = "Active";
        private double _width;
        private double _height;
        private string _approvalStatus = "Approved";

        public ElementId OpeningId
        {
            get => _openingId;
            set => SetProperty(ref _openingId, value);
        }

        public ElementId MepElementId
        {
            get => _mepElementId;
            set => SetProperty(ref _mepElementId, value);
        }

        public ElementId WallId
        {
            get => _wallId;
            set => SetProperty(ref _wallId, value);
        }

        public XYZ OriginalLocation
        {
            get => _originalLocation;
            set => SetProperty(ref _originalLocation, value);
        }

        public XYZ CurrentLocation
        {
            get => _currentLocation;
            set => SetProperty(ref _currentLocation, value);
        }

        public string MepType
        {
            get => _mepType;
            set => SetProperty(ref _mepType, value);
        }

        public DateTime LastUpdated
        {
            get => _lastUpdated;
            set => SetProperty(ref _lastUpdated, value);
        }

        public bool IsAutoUpdateEnabled
        {
            get => _isAutoUpdateEnabled;
            set => SetProperty(ref _isAutoUpdateEnabled, value);
        }

        public string Status
        {
            get => _status;
            set => SetProperty(ref _status, value);
        }

        public double Width
        {
            get => _width;
            set => SetProperty(ref _width, value);
        }

        public double Height
        {
            get => _height;
            set => SetProperty(ref _height, value);
        }

        public string ApprovalStatus
        {
            get => _approvalStatus;
            set => SetProperty(ref _approvalStatus, value);
        }

        /// <summary>
        /// Gets the movement distance from original position
        /// </summary>
        public double MovementDistance => OriginalLocation.DistanceTo(CurrentLocation);

        /// <summary>
        /// Gets a formatted movement distance string
        /// </summary>
        public string MovementDistanceText => MovementDistance > 0.01 ? $"{MovementDistance:F2} mm" : "No movement";

        /// <summary>
        /// Gets the status icon based on current status
        /// </summary>
        public string StatusIcon => Status switch
        {
            "Active" => "✓",
            "Moved" => "↗",
            "Error" => "✗",
            "Disabled" => "⏸",
            _ => "?"
        };

        /// <summary>
        /// Gets a formatted timestamp string
        /// </summary>
        public string FormattedLastUpdated => LastUpdated.ToString("HH:mm:ss");

        public OpeningStatus()
        {
            LastUpdated = DateTime.Now;
        }

        public OpeningStatus(ElementId openingId, ElementId mepElementId, ElementId wallId, XYZ location, string mepType, double width, double height)
        {
            OpeningId = openingId;
            MepElementId = mepElementId;
            WallId = wallId;
            OriginalLocation = location;
            CurrentLocation = location;
            MepType = mepType;
            Width = width;
            Height = height;
            LastUpdated = DateTime.Now;
            Status = "Active";
        }

        /// <summary>
        /// Creates an OpeningStatus from an OpeningRequirement
        /// </summary>
        public static OpeningStatus FromOpeningRequirement(OpeningRequirement requirement, ElementId openingElementId)
        {
            if (requirement.ElementId == null || requirement.WallId == null || requirement.Location == null)
            {
                throw new ArgumentException("OpeningRequirement is missing required data");
            }

            return new OpeningStatus(
                openingElementId,
                requirement.ElementId!,
                requirement.WallId!,
                requirement.Location!,
                requirement.MepType ?? "Unknown",
                requirement.Width,
                requirement.Height
            );
        }

        /// <summary>
        /// Updates the current location and status
        /// </summary>
        public void UpdateLocation(XYZ newLocation)
        {
            CurrentLocation = newLocation;
            LastUpdated = DateTime.Now;
            
            // Update status based on movement
            if (MovementDistance > 0.01)
            {
                Status = "Moved";
            }
            else
            {
                Status = "Active";
            }
        }

        /// <summary>
        /// Marks the opening as having an error
        /// </summary>
        public void MarkAsError(string errorMessage = "")
        {
            Status = "Error";
            LastUpdated = DateTime.Now;
        }

        /// <summary>
        /// Resets the opening to active status
        /// </summary>
        public void ResetToActive()
        {
            Status = "Active";
            LastUpdated = DateTime.Now;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}