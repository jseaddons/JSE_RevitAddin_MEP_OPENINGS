using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ErrorHandling
{
    /// <summary>
    /// Team D: Fail-safe error policy implementation.
    /// Continues processing on errors (non-critical) and logs appropriately.
    /// Prevents crashes and data loss by always continuing unless explicitly critical.
    /// </summary>
    public class FailSafeErrorPolicy : IErrorPolicy
    {
        private readonly ILogger _logger;
        private readonly bool _isDeploymentMode;
        
        /// <summary>
        /// Creates a fail-safe error policy with logging
        /// </summary>
        public FailSafeErrorPolicy(ILogger logger = null, bool isDeploymentMode = false)
        {
            _logger = logger ?? LoggerAdapter.Default;
            _isDeploymentMode = isDeploymentMode;
        }
        
        public ErrorHandlingResult Handle(
            Exception exception,
            string scope,
            object context = null,
            bool critical = false)
        {
            var result = new ErrorHandlingResult
            {
                Exception = exception,
                Scope = scope ?? "Unknown",
                Context = context,
                ShouldContinue = !critical, // Only stop on critical errors
                LogLevel = DetermineLogLevel(exception, critical),
                Message = FormatErrorMessage(exception, scope, context)
            };
            
            // Log according to policy
            if (!_isDeploymentMode || critical)
            {
                switch (result.LogLevel)
                {
                    case "Error":
                        _logger.Error(result.Message, exception, scope);
                        break;
                    case "Warning":
                        _logger.Warning(result.Message, scope);
                        break;
                    case "Info":
                        _logger.Info(result.Message, scope);
                        break;
                    case "Debug":
                        _logger.Debug(result.Message, scope);
                        break;
                }
            }
            
            // ✅ FAIL-SAFE: Always continue unless critical
            // This prevents one error from aborting entire placement operation
            return result;
        }
        
        public ErrorSummary AggregateErrors(IEnumerable<ErrorHandlingResult> errors)
        {
            if (errors == null)
                return new ErrorSummary();
            
            var errorList = errors.ToList();
            var summary = new ErrorSummary
            {
                TotalErrors = errorList.Count,
                CriticalErrors = errorList.Count(e => !e.ShouldContinue),
                Warnings = errorList.Count(e => e.LogLevel == "Warning"),
                ErrorMessages = errorList.Select(e => e.Message).ToList(),
                ShouldAbort = errorList.Any(e => !e.ShouldContinue) // Abort only if any critical
            };
            
            return summary;
        }
        
        private string DetermineLogLevel(Exception exception, bool critical)
        {
            if (critical)
                return "Error";
            
            // Determine log level based on exception type
            return exception switch
            {
                ArgumentNullException => "Warning",
                ArgumentException => "Warning",
                InvalidOperationException => "Error",
                System.Data.SQLite.SQLiteException => "Error",
                _ => "Error"
            };
        }
        
        private string FormatErrorMessage(Exception exception, string scope, object context)
        {
            string contextInfo = context?.ToString() ?? "No context";
            return $"[{scope}] {exception.GetType().Name}: {exception.Message} | Context: {contextInfo}";
        }
    }
}

