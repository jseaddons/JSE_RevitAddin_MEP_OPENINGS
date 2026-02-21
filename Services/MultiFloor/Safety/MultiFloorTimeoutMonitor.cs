using System;
using System.Diagnostics;
using System.Threading;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor.Safety
{
    /// <summary>
    /// Timeout monitoring for multi-floor operations.
    /// Prevents indefinite hanging by tracking per-floor and total processing time.
    /// </summary>
    public class MultiFloorTimeoutMonitor : IDisposable
    {
        private readonly Stopwatch _stopwatch;
        private readonly int _floorTimeoutMs;
        private readonly int _totalTimeoutMs;
        private readonly Action<string> _logger;
        private string _currentFloor;
        private Stopwatch _floorStopwatch;
        
        /// <summary>
        /// Event raised when a floor timeout occurs
        /// </summary>
        public event EventHandler<TimeoutEventArgs> FloorTimeout;
        
        /// <summary>
        /// Event raised when total timeout occurs
        /// </summary>
        public event EventHandler<TimeoutEventArgs> TotalTimeout;
        
        public MultiFloorTimeoutMonitor(
            int floorTimeoutMinutes = 1, 
            int totalTimeoutMinutes = 10,
            Action<string> logger = null)
        {
            _floorTimeoutMs = floorTimeoutMinutes * 60 * 1000;
            _totalTimeoutMs = totalTimeoutMinutes * 60 * 1000;
            _logger = logger ?? (msg => DebugLogger.Info(msg));
            _stopwatch = new Stopwatch();
            _floorStopwatch = new Stopwatch();
        }
        
        /// <summary>
        /// Start monitoring the entire multi-floor operation
        /// </summary>
        public void StartOperation()
        {
            _stopwatch.Restart();
            _logger?.Invoke($"[TimeoutMonitor] Started: Floor limit = {_floorTimeoutMs/60000}min, Total limit = {_totalTimeoutMs/60000}min");
        }
        
        /// <summary>
        /// Start monitoring a specific floor
        /// </summary>
        public void StartFloor(string floorName)
        {
            _currentFloor = floorName;
            _floorStopwatch.Restart();
            _logger?.Invoke($"[TimeoutMonitor] Started floor: {floorName}");
        }
        
        /// <summary>
        /// Stop monitoring current floor
        /// </summary>
        public void EndFloor(string floorName)
        {
            _floorStopwatch.Stop();
            var elapsed = _floorStopwatch.Elapsed;
            _logger?.Invoke($"[TimeoutMonitor] Completed floor: {floorName} in {elapsed.TotalSeconds:F1}s");
        }
        
        /// <summary>
        /// Check if current floor has exceeded timeout
        /// </summary>
        public bool CheckFloorTimeout()
        {
            if (!_floorStopwatch.IsRunning)
                return false;
                
            if (_floorStopwatch.ElapsedMilliseconds > _floorTimeoutMs)
            {
                _floorStopwatch.Stop();
                var message = $"Floor '{_currentFloor}' exceeded {_floorTimeoutMs/60000} minute timeout";
                _logger?.Invoke($"[TimeoutMonitor] ⏱ TIMEOUT: {message}");
                
                FloorTimeout?.Invoke(this, new TimeoutEventArgs
                {
                    FloorName = _currentFloor,
                    ElapsedMs = _floorStopwatch.ElapsedMilliseconds,
                    LimitMs = _floorTimeoutMs,
                    Message = message,
                    IsTotalTimeout = false
                });
                
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// Check if total operation has exceeded timeout
        /// </summary>
        public bool CheckTotalTimeout()
        {
            if (!_stopwatch.IsRunning)
                return false;
                
            if (_stopwatch.ElapsedMilliseconds > _totalTimeoutMs)
            {
                var message = $"Total operation exceeded {_totalTimeoutMs/60000} minute timeout";
                _logger?.Invoke($"[TimeoutMonitor] ⏱ TOTAL TIMEOUT: {message}");
                
                TotalTimeout?.Invoke(this, new TimeoutEventArgs
                {
                    FloorName = _currentFloor,
                    ElapsedMs = _stopwatch.ElapsedMilliseconds,
                    LimitMs = _totalTimeoutMs,
                    Message = message,
                    IsTotalTimeout = true
                });
                
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// Get current operation statistics
        /// </summary>
        public TimeoutStats GetStats()
        {
            return new TimeoutStats
            {
                TotalElapsedMs = _stopwatch.ElapsedMilliseconds,
                CurrentFloorElapsedMs = _floorStopwatch.ElapsedMilliseconds,
                CurrentFloor = _currentFloor,
                FloorTimeoutMs = _floorTimeoutMs,
                TotalTimeoutMs = _totalTimeoutMs,
                IsFloorTimedOut = _floorStopwatch.ElapsedMilliseconds > _floorTimeoutMs,
                IsTotalTimedOut = _stopwatch.ElapsedMilliseconds > _totalTimeoutMs
            };
        }
        
        /// <summary>
        /// Create a cancellation token that triggers on timeout
        /// </summary>
        public CancellationTokenSource CreateFloorCancellationToken()
        {
            return new CancellationTokenSource(_floorTimeoutMs);
        }
        
        public void Dispose()
        {
            _stopwatch?.Stop();
            _floorStopwatch?.Stop();
        }
    }
    
    /// <summary>
    /// Event args for timeout events
    /// </summary>
    public class TimeoutEventArgs : EventArgs
    {
        public string FloorName { get; set; }
        public long ElapsedMs { get; set; }
        public long LimitMs { get; set; }
        public string Message { get; set; }
        public bool IsTotalTimeout { get; set; }
    }
    
    /// <summary>
    /// Statistics for timeout monitoring
    /// </summary>
    public class TimeoutStats
    {
        public long TotalElapsedMs { get; set; }
        public long CurrentFloorElapsedMs { get; set; }
        public string CurrentFloor { get; set; }
        public long FloorTimeoutMs { get; set; }
        public long TotalTimeoutMs { get; set; }
        public bool IsFloorTimedOut { get; set; }
        public bool IsTotalTimedOut { get; set; }
        
        public TimeSpan TotalElapsed => TimeSpan.FromMilliseconds(TotalElapsedMs);
        public TimeSpan CurrentFloorElapsed => TimeSpan.FromMilliseconds(CurrentFloorElapsedMs);
        public TimeSpan FloorTimeout => TimeSpan.FromMilliseconds(FloorTimeoutMs);
        public TimeSpan TotalTimeout => TimeSpan.FromMilliseconds(TotalTimeoutMs);
    }
}
