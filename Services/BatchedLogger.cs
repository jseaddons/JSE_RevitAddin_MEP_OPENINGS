using System;
using System.Text;
using System.Threading;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ MEMORY OPTIMIZATION: Batches log messages to reduce string allocations and file I/O.
    /// Flushes periodically (every N messages or time-based) and on critical operations.
    /// </summary>
    public class BatchedLogger
    {
        private readonly Action<string> _originalLogger;
        private readonly StringBuilder _buffer;
        private readonly object _bufferLock;
        private readonly int _batchSize;
        private readonly TimeSpan _flushInterval;
        private DateTime _lastFlush;
        private int _messageCount;
        private readonly string _logFileName;

        /// <summary>
        /// Creates a batched logger wrapper around an existing logging action
        /// </summary>
        /// <param name="originalLogger">The underlying logger to write to</param>
        /// <param name="batchSize">Flush after N messages (default: 50)</param>
        /// <param name="flushIntervalSeconds">Flush after N seconds (default: 5 seconds)</param>
        /// <param name="logFileName">Optional: Log file name for error reporting</param>
        public BatchedLogger(Action<string> originalLogger, int batchSize = 50, int flushIntervalSeconds = 5, string logFileName = null)
        {
            _originalLogger = originalLogger ?? throw new ArgumentNullException(nameof(originalLogger));
            _buffer = new StringBuilder(10000); // Pre-allocate 10KB buffer
            _bufferLock = new object();
            _batchSize = batchSize;
            _flushInterval = TimeSpan.FromSeconds(flushIntervalSeconds);
            _lastFlush = DateTime.Now;
            _messageCount = 0;
            _logFileName = logFileName;
        }

        /// <summary>
        /// Logs a message (batched or immediately if critical)
        /// </summary>
        public void Log(string message)
        {
            // ✅ DEPLOYMENT MODE: Skip all logging if deployment mode is enabled (prevents string allocation)
            if (DeploymentConfiguration.DeploymentMode)
                return;
                
            if (string.IsNullOrEmpty(message))
                return;

            try
            {
                lock (_bufferLock)
                {
                    _buffer.AppendLine(message);
                    _messageCount++;

                    // ✅ AUTO-FLUSH: Flush if batch size reached or time interval elapsed
                    bool shouldFlush = false;
                    if (_messageCount >= _batchSize)
                    {
                        shouldFlush = true;
                    }
                    else
                    {
                        var elapsed = DateTime.Now - _lastFlush;
                        if (elapsed >= _flushInterval)
                        {
                            shouldFlush = true;
                        }
                    }

                    if (shouldFlush)
                    {
                        FlushInternal();
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ SAFETY: Fallback to direct logging if batching fails
                try
                {
                    _originalLogger($"[BatchedLogger] ERROR: Failed to batch log message. Error: {ex.Message}. Original message: {message}");
                }
                catch
                {
                    // Last resort: ignore (don't crash the application)
                }
            }
        }

        /// <summary>
        /// ✅ CRITICAL: Force immediate flush (call before important operations or on completion)
        /// </summary>
        public void Flush()
        {
            lock (_bufferLock)
            {
                FlushInternal();
            }
        }

        private void FlushInternal()
        {
            if (_buffer.Length == 0)
                return;

            try
            {
                string batchContent = _buffer.ToString();
                _originalLogger(batchContent);

                // Clear buffer and reset counters
                _buffer.Clear();
                _messageCount = 0;
                _lastFlush = DateTime.Now;
            }
            catch (Exception ex)
            {
                // ✅ SAFETY: Try to log the error, but don't fail entirely
                try
                {
                    // Log error to a safe location (don't use _originalLogger as it might fail too)
                    System.Diagnostics.Debug.WriteLine($"[BatchedLogger] ERROR flushing logs: {ex.Message}");
                    if (!string.IsNullOrEmpty(_logFileName))
                    {
                        SafeFileLogger.SafeAppendText("batched_logger_errors.log", 
                            $"[{DateTime.Now}] ERROR flushing batch for '{_logFileName}': {ex.Message}\n");
                    }
                    
                    // ✅ CRITICAL: Clear buffer anyway to prevent memory leak
                    _buffer.Clear();
                    _messageCount = 0;
                    _lastFlush = DateTime.Now;
                }
                catch
                {
                    // Last resort: clear buffer to prevent memory leak
                    _buffer.Clear();
                    _messageCount = 0;
                }
            }
        }

        /// <summary>
        /// ✅ CRITICAL: Dispose pattern - ensures all buffered logs are flushed
        /// </summary>
        public void Dispose()
        {
            Flush();
        }
    }
}

