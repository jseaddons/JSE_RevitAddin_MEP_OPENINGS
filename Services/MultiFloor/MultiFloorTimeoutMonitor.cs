using System;
using System.Diagnostics;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Monitors timeout for multi-floor operations to prevent Revit hangs
    /// Tracks both per-floor and total operation timeouts
    /// </summary>
    public class MultiFloorTimeoutMonitor : IDisposable
    {
        private readonly Stopwatch _stopwatch;
        private readonly TimeSpan _floorTimeout;
        private readonly TimeSpan _totalTimeout;
        private readonly Action<string> _logger;
        private bool _disposed;

        public MultiFloorTimeoutMonitor(
            int floorTimeoutMinutes = 5,
            int totalTimeoutMinutes = 15,
            Action<string> logger = null)
        {
            _floorTimeout = TimeSpan.FromMinutes(floorTimeoutMinutes);
            _totalTimeout = TimeSpan.FromMinutes(totalTimeoutMinutes);
            _logger = logger ?? (msg => { });
            _stopwatch = new Stopwatch();
        }

        public void StartOperation()
        {
            _stopwatch.Start();
            _logger($"[TIMEOUT-MONITOR] Started (floor: {_floorTimeout.TotalMinutes}min, total: {_totalTimeout.TotalMinutes}min)");
        }

        public bool CheckFloorTimeout(string floorName = null)
        {
            if (!_stopwatch.IsRunning) return false;
            
            if (_stopwatch.Elapsed > _floorTimeout)
            {
                var msg = floorName != null 
                    ? $"[TIMEOUT-MONITOR] ⏱️ FLOOR TIMEOUT: {floorName} exceeded {_floorTimeout.TotalMinutes} minutes"
                    : $"[TIMEOUT-MONITOR] ⏱️ FLOOR TIMEOUT: Exceeded {_floorTimeout.TotalMinutes} minutes";
                _logger(msg);
                return true;
            }
            return false;
        }

        public bool CheckTotalTimeout()
        {
            if (!_stopwatch.IsRunning) return false;
            
            if (_stopwatch.Elapsed > _totalTimeout)
            {
                _logger($"[TIMEOUT-MONITOR] ⏱️ TOTAL TIMEOUT: Exceeded {_totalTimeout.TotalMinutes} minutes");
                return true;
            }
            return false;
        }

        public TimeSpan Elapsed => _stopwatch.Elapsed;

        public void Dispose()
        {
            if (!_disposed)
            {
                _stopwatch?.Stop();
                _disposed = true;
            }
        }
    }
}
