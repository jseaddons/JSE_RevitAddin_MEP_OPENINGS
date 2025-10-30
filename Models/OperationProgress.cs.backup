using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the progress of an operation
    /// </summary>
    public class OperationProgress : INotifyPropertyChanged
    {
        private string _operationName = string.Empty;
        private DateTime _startTime;
        private DateTime? _endTime;
        private DateTime _lastUpdated;
        private int _currentStep;
        private int _totalSteps;
        private string _currentStepDescription = string.Empty;
        private bool _isCompleted;
        private bool _success;

        public string OperationName
        {
            get => _operationName;
            set => SetProperty(ref _operationName, value);
        }

        public DateTime StartTime
        {
            get => _startTime;
            set => SetProperty(ref _startTime, value);
        }

        public DateTime? EndTime
        {
            get => _endTime;
            set => SetProperty(ref _endTime, value);
        }

        public DateTime LastUpdated
        {
            get => _lastUpdated;
            set => SetProperty(ref _lastUpdated, value);
        }

        public int CurrentStep
        {
            get => _currentStep;
            set => SetProperty(ref _currentStep, value);
        }

        public int TotalSteps
        {
            get => _totalSteps;
            set => SetProperty(ref _totalSteps, value);
        }

        public string CurrentStepDescription
        {
            get => _currentStepDescription;
            set => SetProperty(ref _currentStepDescription, value);
        }

        public bool IsCompleted
        {
            get => _isCompleted;
            set => SetProperty(ref _isCompleted, value);
        }

        public bool Success
        {
            get => _success;
            set => SetProperty(ref _success, value);
        }

        /// <summary>
        /// Gets the progress percentage (0-100)
        /// </summary>
        public int ProgressPercentage => TotalSteps > 0 ? (int)((double)CurrentStep / TotalSteps * 100) : 0;

        /// <summary>
        /// Gets the duration of the operation
        /// </summary>
        public TimeSpan? Duration => EndTime.HasValue ? EndTime.Value - StartTime : DateTime.Now - StartTime;

        /// <summary>
        /// Gets a formatted duration string
        /// </summary>
        public string FormattedDuration
        {
            get
            {
                var duration = Duration;
                if (duration.HasValue)
                {
                    if (duration.Value.TotalSeconds < 60)
                        return $"{duration.Value.TotalSeconds:F1}s";
                    else if (duration.Value.TotalMinutes < 60)
                        return $"{duration.Value.TotalMinutes:F1}m";
                    else
                        return $"{duration.Value.TotalHours:F1}h";
                }
                return "0s";
            }
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
