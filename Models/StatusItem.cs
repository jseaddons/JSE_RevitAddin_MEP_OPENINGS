using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a status item in the status history
    /// </summary>
    public class StatusItem : INotifyPropertyChanged
    {
        private DateTime _timestamp;
        private string _message = string.Empty;
        private StatusType _type = StatusType.Info;
        private string _operation = string.Empty;

        public DateTime Timestamp
        {
            get => _timestamp;
            set => SetProperty(ref _timestamp, value);
        }

        public string Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        public StatusType Type
        {
            get => _type;
            set => SetProperty(ref _type, value);
        }

        public string Operation
        {
            get => _operation;
            set => SetProperty(ref _operation, value);
        }

        /// <summary>
        /// Gets a formatted timestamp string
        /// </summary>
        public string FormattedTimestamp => Timestamp.ToString("HH:mm:ss");

        /// <summary>
        /// Gets the status icon based on type
        /// </summary>
        public string StatusIcon => Type switch
        {
            StatusType.Success => "✓",
            StatusType.Warning => "⚠",
            StatusType.Error => "✗",
            StatusType.Processing => "⟳",
            StatusType.Info => "ℹ",
            _ => "•"
        };

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

    /// <summary>
    /// Status types for different message categories
    /// </summary>
    public enum StatusType
    {
        Info,
        Success,
        Warning,
        Error,
        Processing
    }
}
