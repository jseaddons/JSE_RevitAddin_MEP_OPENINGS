using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using CommunityToolkit.Mvvm.ComponentModel;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    /// <summary>
    /// ViewModel for displaying opening status information in the UI
    /// </summary>
    public class OpeningStatusItem : ObservableObject
    {
        private readonly OpeningStatus _status;
        private readonly Document? _document;

        public ElementId OpeningId => _status.OpeningId;
        public ElementId MepElementId => _status.MepElementId;
        public ElementId WallId => _status.WallId;

        // Opening Details
        public string OpeningIdText => $"Opening ID: {_status.OpeningId.IntegerValue}";
        public string OpeningDimensions => $"{_status.Width:F1} × {_status.Height:F1} mm";
        public string OpeningLocation => $"X:{_status.CurrentLocation.X:F2}, Y:{_status.CurrentLocation.Y:F2}, Z:{_status.CurrentLocation.Z:F2}";
        public string OpeningFamilyName { get; }

        // MEP Element Details
        public string MepElementIdText => $"MEP ID: {_status.MepElementId.IntegerValue}";
        public string MepElementType => _status.MepType;
        public string MepElementSize { get; }
        public string MepElementLocation { get; }

        // Wall Details
        public string WallIdText => $"Wall ID: {_status.WallId.IntegerValue}";
        public string WallType { get; }

        // Status Information
        public string Status => _status.Status;
        public string StatusIcon => _status.StatusIcon;
        public string LastUpdatedText => _status.FormattedLastUpdated;
        public bool CanUpdate => _status.Status != "Active";
        public string MovementDistance => _status.MovementDistanceText;

        public OpeningStatusItem(OpeningStatus status, Document? document = null)
        {
            _status = status;
            _document = document;

            // Get additional information from document if available
            OpeningFamilyName = GetOpeningFamilyName();
            MepElementSize = GetMepElementSize();
            MepElementLocation = GetMepElementLocation();
            WallType = GetWallType();

            // Subscribe to status changes
            _status.PropertyChanged += OnStatusPropertyChanged;
        }

        private void OnStatusPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            // Forward property changes to the UI
            OnPropertyChanged(e.PropertyName);
        }

        private string GetOpeningFamilyName()
        {
            if (_document == null) return "Unknown";

            try
            {
                var opening = _document.GetElement(_status.OpeningId) as FamilyInstance;
                return opening?.Symbol?.Family?.Name ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        private string GetMepElementSize()
        {
            if (_document == null) return "Unknown size";

            try
            {
                var element = _document.GetElement(_status.MepElementId);

                if (element is MEPCurve mepCurve)
                {
                    var category = mepCurve.Category?.Name ?? "Unknown";
                    
                    if (category.Contains("Pipe"))
                    {
                        var diameter = mepCurve.LookupParameter("Diameter")?.AsDouble() ?? 0;
                        return $"Ø{diameter * 304.8:F1} mm"; // Convert feet to mm
                    }
                    else if (category.Contains("Duct"))
                    {
                        var width = mepCurve.LookupParameter("Width")?.AsDouble() ?? 0;
                        var height = mepCurve.LookupParameter("Height")?.AsDouble() ?? 0;
                        return $"{width * 304.8:F1} × {height * 304.8:F1} mm";
                    }
                    else if (category.Contains("Cable Tray"))
                    {
                        var width = mepCurve.LookupParameter("Width")?.AsDouble() ?? 0;
                        var height = mepCurve.LookupParameter("Height")?.AsDouble() ?? 0;
                        return $"{width * 304.8:F1} × {height * 304.8:F1} mm";
                    }
                }

                return "Unknown size";
            }
            catch
            {
                return "Unknown size";
            }
        }

        private string GetMepElementLocation()
        {
            if (_document == null) return "Unknown location";

            try
            {
                var element = _document.GetElement(_status.MepElementId);
                var location = element?.Location;

                if (location is LocationPoint point)
                {
                    var xyz = point.Point;
                    return $"X:{xyz.X:F2}, Y:{xyz.Y:F2}, Z:{xyz.Z:F2}";
                }
                else if (location is LocationCurve curve)
                {
                    var start = curve.Curve.GetEndPoint(0);
                    return $"Start: X:{start.X:F2}, Y:{start.Y:F2}, Z:{start.Z:F2}";
                }

                return "Unknown location";
            }
            catch
            {
                return "Unknown location";
            }
        }

        private string GetWallType()
        {
            if (_document == null) return "Unknown wall";

            try
            {
                var wall = _document.GetElement(_status.WallId) as Wall;
                return wall?.WallType?.Name ?? "Unknown wall";
            }
            catch
            {
                return "Unknown wall";
            }
        }

        /// <summary>
        /// Gets the underlying OpeningStatus object
        /// </summary>
        public OpeningStatus GetStatus() => _status;
    }
}
