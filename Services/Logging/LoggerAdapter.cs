using System;
using System.Text;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Logging
{
    /// <summary>
    /// Team D: Adapter that wraps DebugLogger + SafeFileLogger for ILogger interface.
    /// Preserves existing log file names and message prefixes.
    /// Provides structured logging capability with correlation IDs.
    /// </summary>
    public class LoggerAdapter : ILogger
    {
        private static ILogger _defaultInstance;
        private static readonly object _lock = new object();
        
        private readonly bool _isDeploymentMode;
        private string _correlationId;
        
        /// <summary>
        /// Creates a logger adapter that wraps existing logging infrastructure
        /// </summary>
        public LoggerAdapter(string? correlationId = null, bool isDeploymentMode = false)
        {
            _correlationId = correlationId ?? Guid.NewGuid().ToString("N").Substring(0, 8);
            _isDeploymentMode = isDeploymentMode || DeploymentConfiguration.DeploymentMode;
        }
        
        public string CorrelationId
        {
            get => _correlationId;
            set => _correlationId = value ?? Guid.NewGuid().ToString("N").Substring(0, 8);
        }
        
        /// <summary>
        /// Gets the default logger instance (thread-safe singleton)
        /// </summary>
        public static ILogger Default
        {
            get
            {
                if (_defaultInstance == null)
                {
                    lock (_lock)
                    {
                        if (_defaultInstance == null)
                        {
                            _defaultInstance = new LoggerAdapter();
                        }
                    }
                }
                return _defaultInstance;
            }
        }
        
        public void Info(string message, string? scope = null)
        {
            if (_isDeploymentMode)
                return; // Skip in deployment mode
            
            // ✅ PRESERVE: Use existing DebugLogger.Info with same format
            string formattedMessage = FormatMessage(message, scope, "INFO");
            DebugLogger.Info(formattedMessage);
            
            // ✅ PRESERVE: Use existing SafeFileLogger with same file names
            SafeFileLogger.SafeAppendText("placement_debug.log", $"{formattedMessage}\n");
        }
        
        public void Warning(string message, string? scope = null)
        {
            // Warnings always logged (even in deployment mode)
            string formattedMessage = FormatMessage(message, scope, "WARN");
            DebugLogger.Warning(formattedMessage);
            SafeFileLogger.SafeAppendText("placement_errors.log", $"{formattedMessage}\n");
        }
        
        public void Error(string message, Exception? exception = null, string? scope = null)
        {
            // Errors always logged (even in deployment mode)
            string formattedMessage = FormatMessage(message, scope, "ERROR");
            
            if (exception != null)
            {
                formattedMessage += $"\nException: {exception.GetType().Name}: {exception.Message}";
                if (exception.StackTrace != null && !_isDeploymentMode)
                {
                    formattedMessage += $"\nStackTrace: {exception.StackTrace}";
                }
            }
            
            DebugLogger.Error(formattedMessage);
            SafeFileLogger.SafeAppendText("placement_errors.log", $"{formattedMessage}\n");
        }
        
        public void Debug(string message, string? scope = null)
        {
            if (_isDeploymentMode)
                return; // Skip debug in deployment mode
            
            // ✅ PRESERVE: Use existing DebugLogger format
            string formattedMessage = FormatMessage(message, scope, "DEBUG");
            DebugLogger.Info(formattedMessage); // DebugLogger doesn't have Debug method, use Info
        }
        
        public void LogStructured(string level, string message, object? data = null, string? scope = null)
        {
            if (_isDeploymentMode && level != "ERROR" && level != "WARN")
                return; // Skip non-critical in deployment mode
            
            // Format: [UTC ISO][Stage][Level] Message | zoneId= | sleeveId= | data={json}
            var timestamp = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
            var scopePart = scope != null ? $"[{scope}]" : "";
            var levelPart = $"[{level}]";
            
            var sb = new StringBuilder();
            sb.Append($"[{timestamp}]{scopePart}{levelPart} {message}");
            
            if (data != null)
            {
                try
                {
                    var json = System.Text.Json.JsonSerializer.Serialize(data);
                    sb.Append($" | data={json}");
                }
                catch
                {
                    sb.Append($" | data={data}");
                }
            }
            
            string formattedMessage = sb.ToString();
            
            // Route to appropriate logger based on level
            switch (level.ToUpper())
            {
                case "ERROR":
                    Error(formattedMessage, null, scope);
                    break;
                case "WARN":
                case "WARNING":
                    Warning(formattedMessage, scope);
                    break;
                case "INFO":
                    Info(formattedMessage, scope);
                    break;
                case "DEBUG":
                    Debug(formattedMessage, scope);
                    break;
            }
        }
        
        private string FormatMessage(string message, string scope, string level)
        {
            // ✅ PRESERVE: Maintain existing message format with correlation ID
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var scopePart = scope != null ? $"[{scope}]" : "";
            var correlationId = CorrelationId;
            var correlationPart = correlationId != null ? $" | correlationId={correlationId}" : "";
            
            return $"[{timestamp}]{scopePart} [{level}]{correlationPart} {message}";
        }
    }
}

