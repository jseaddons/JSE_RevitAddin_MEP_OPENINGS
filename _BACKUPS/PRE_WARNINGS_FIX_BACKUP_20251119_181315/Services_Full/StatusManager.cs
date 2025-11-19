using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing status updates and progress tracking
    /// </summary>
    public class StatusManager : INotifyPropertyChanged
    {
        private readonly object _lockObject = new object();
        private readonly ObservableCollection<StatusItem> _statusHistory;
        private readonly Dictionary<string, OperationProgress> _operationProgress;
        
        private string _currentStatus = "Ready";
        private StatusType _statusType = StatusType.Info;
        private int _overallProgressPercentage = 0;
        private string _progressMessage = "";
        private bool _isOperationRunning = false;
        private string _currentOperation = "";

        public event EventHandler<StatusUpdateEventArgs>? StatusUpdated;
        public event EventHandler<ProgressUpdateEventArgs>? ProgressUpdated;
        public event EventHandler<OperationCompletedEventArgs>? OperationCompleted;

        public StatusManager()
        {
            _statusHistory = new ObservableCollection<StatusItem>();
            _operationProgress = new Dictionary<string, OperationProgress>();
            StatusHistory = new ReadOnlyObservableCollection<StatusItem>(_statusHistory);
        }

        public string CurrentStatus
        {
            get => _currentStatus;
            private set
            {
                if (_currentStatus != value)
                {
                    _currentStatus = value;
                    OnPropertyChanged();
                }
            }
        }

        public StatusType StatusType
        {
            get => _statusType;
            private set
            {
                if (_statusType != value)
                {
                    _statusType = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(StatusColor));
                    OnPropertyChanged(nameof(StatusBackgroundColor));
                }
            }
        }

        public int OverallProgressPercentage
        {
            get => _overallProgressPercentage;
            private set
            {
                if (_overallProgressPercentage != value)
                {
                    _overallProgressPercentage = value;
                    OnPropertyChanged();
                }
            }
        }

        public string ProgressMessage
        {
            get => _progressMessage;
            private set
            {
                if (_progressMessage != value)
                {
                    _progressMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsOperationRunning
        {
            get => _isOperationRunning;
            private set
            {
                if (_isOperationRunning != value)
                {
                    _isOperationRunning = value;
                    OnPropertyChanged();
                }
            }
        }

        public ReadOnlyObservableCollection<StatusItem> StatusHistory { get; }

        public System.Windows.Media.Brush StatusColor
        {
            get
            {
                return StatusType switch
                {
                    StatusType.Success => System.Windows.Media.Brushes.Green,
                    StatusType.Warning => System.Windows.Media.Brushes.Orange,
                    StatusType.Error => System.Windows.Media.Brushes.Red,
                    StatusType.Processing => System.Windows.Media.Brushes.Blue,
                    StatusType.Info => System.Windows.Media.Brushes.DarkBlue,
                    _ => System.Windows.Media.Brushes.Black
                };
            }
        }

        public System.Windows.Media.Brush StatusBackgroundColor
        {
            get
            {
                return StatusType switch
                {
                    StatusType.Success => new SolidColorBrush(System.Windows.Media.Color.FromRgb(163, 190, 140)), // Green
                    StatusType.Warning => new SolidColorBrush(System.Windows.Media.Color.FromRgb(235, 203, 139)), // Yellow
                    StatusType.Error => new SolidColorBrush(System.Windows.Media.Color.FromRgb(191, 97, 106)),   // Red
                    StatusType.Processing => new SolidColorBrush(System.Windows.Media.Color.FromRgb(136, 192, 208)), // Blue
                    StatusType.Info => new SolidColorBrush(System.Windows.Media.Color.FromRgb(94, 129, 172)),    // Dark Blue
                    _ => new SolidColorBrush(System.Windows.Media.Color.FromRgb(216, 222, 233)) // Light Gray
                };
            }
        }

        /// <summary>
        /// Starts a new operation
        /// </summary>
        public void StartOperation(string operationName)
        {
            lock (_lockObject)
            {
                _currentOperation = operationName;
                IsOperationRunning = true;
                
                _operationProgress[operationName] = new OperationProgress
                {
                    OperationName = operationName,
                    StartTime = DateTime.Now,
                    CurrentStep = 0,
                    TotalSteps = 0
                };
                
                UpdateStatus($"Starting {operationName}...", StatusType.Processing);
            }
        }

        /// <summary>
        /// Updates progress for an operation
        /// </summary>
        public void UpdateProgress(string operationName, int current, int total, string stepDescription = "")
        {
            lock (_lockObject)
            {
                if (_operationProgress.TryGetValue(operationName, out var progress))
                {
                    progress.CurrentStep = current;
                    progress.TotalSteps = total;
                    progress.CurrentStepDescription = stepDescription;
                    progress.LastUpdated = DateTime.Now;
                    
                    // Calculate overall progress
                    CalculateOverallProgress();
                    
                    ProgressMessage = string.IsNullOrEmpty(stepDescription) 
                        ? $"{operationName}: {current}/{total}" 
                        : $"{stepDescription} ({current}/{total})";
                    
                    OnPropertyChanged(nameof(ProgressMessage));
                    OnPropertyChanged(nameof(OverallProgressPercentage));
                    
                    ProgressUpdated?.Invoke(this, new ProgressUpdateEventArgs
                    {
                        OperationName = operationName,
                        Current = current,
                        Total = total,
                        Percentage = OverallProgressPercentage,
                        StepDescription = stepDescription
                    });
                }
            }
        }

        /// <summary>
        /// Completes an operation
        /// </summary>
        public void CompleteOperation(string operationName, bool success, string message = "")
        {
            lock (_lockObject)
            {
                if (_operationProgress.TryGetValue(operationName, out var progress))
                {
                    progress.IsCompleted = true;
                    progress.EndTime = DateTime.Now;
                    progress.Success = success;
                    
                    var statusType = success ? StatusType.Success : StatusType.Error;
                    var statusMessage = string.IsNullOrEmpty(message) 
                        ? $"{(success ? "Completed" : "Failed")} {operationName}" 
                        : message;
                    
                    UpdateStatus(statusMessage, statusType);
                    
                    OperationCompleted?.Invoke(this, new OperationCompletedEventArgs
                    {
                        OperationName = operationName,
                        Success = success,
                        Message = statusMessage,
                        Duration = (progress.EndTime ?? DateTime.Now) - progress.StartTime
                    });
                }
                
                // Check if all operations are complete
                if (_operationProgress.Values.All(p => p.IsCompleted))
                {
                    IsOperationRunning = false;
                    _currentOperation = "";
                }
            }
        }

        /// <summary>
        /// Updates the current status
        /// </summary>
        public void UpdateStatus(string message, StatusType type)
        {
            lock (_lockObject)
            {
                CurrentStatus = message;
                StatusType = type;
                
                var statusItem = new StatusItem
                {
                    Timestamp = DateTime.Now,
                    Message = message,
                    Type = type,
                    Operation = _currentOperation
                };
                
                _statusHistory.Add(statusItem);
                
                // Limit history size to prevent memory issues
                if (_statusHistory.Count > 1000)
                {
                    _statusHistory.RemoveAt(0);
                }
                
                OnPropertyChanged(nameof(CurrentStatus));
                OnPropertyChanged(nameof(StatusType));
                
                StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                    message, 
                    type, 
                    DateTime.Now, 
                    _currentOperation));
            }
        }

        /// <summary>
        /// Calculates overall progress across all operations
        /// </summary>
        private void CalculateOverallProgress()
        {
            if (_operationProgress.Count == 0)
            {
                OverallProgressPercentage = 0;
                return;
            }
            
            var totalProgress = _operationProgress.Values
                .Where(p => !p.IsCompleted)
                .Sum(p => p.TotalSteps > 0 ? (double)p.CurrentStep / p.TotalSteps : 0);
            
            OverallProgressPercentage = (int)(totalProgress / _operationProgress.Count * 100);
        }

        /// <summary>
        /// Clears the status history
        /// </summary>
        public void ClearHistory()
        {
            lock (_lockObject)
            {
                _statusHistory.Clear();
                _operationProgress.Clear();
                CurrentStatus = "Ready";
                StatusType = StatusType.Info;
                OverallProgressPercentage = 0;
                ProgressMessage = "";
                IsOperationRunning = false;
                _currentOperation = "";
            }
        }

        /// <summary>
        /// Gets the current operation progress
        /// </summary>
        public OperationProgress? GetCurrentOperationProgress()
        {
            lock (_lockObject)
            {
                return _operationProgress.Values.FirstOrDefault(p => !p.IsCompleted);
            }
        }

        /// <summary>
        /// Gets all operation progress
        /// </summary>
        public IReadOnlyList<OperationProgress> GetAllOperationProgress()
        {
            lock (_lockObject)
            {
                return _operationProgress.Values.ToList().AsReadOnly();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
