using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor.Safety
{
    /// <summary>
    /// Failure preprocessor for multi-floor operations.
    /// Handles Revit warnings and errors gracefully to prevent crashes.
    /// </summary>
    public class MultiFloorFailurePreprocessor : IFailuresPreprocessor
    {
        private readonly string _floorName;
        private readonly Action<string> _logger;
        private readonly HashSet<string> _suppressedWarnings;
        
        /// <summary>
        /// Patterns for errors that are known to be harmless and can be treated as warnings
        /// </summary>
        private static readonly string[] HarmlessErrorPatterns = new[]
        {
            "duplicate",
            "coincident",
            "slightly off axis",
            "very small arc",
            "instability detected",
            "very short line",
            "is very small",
            "is too small",
            "joins overlap",
            "elements are too close"
        };
        
        /// <summary>
        /// Warnings that should always be suppressed for floor processing
        /// </summary>
        private static readonly string[] AlwaysSuppressPatterns = new[]
        {
            "sleeve",
            "opening",
            "mep",
            "clash",
            "intersection"
        };
        
        public MultiFloorFailurePreprocessor(string floorName, Action<string> logger = null)
        {
            _floorName = floorName ?? "Unknown";
            _logger = logger ?? (msg => DebugLogger.Info(msg));
            _suppressedWarnings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
        
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            int suppressedCount = 0;
            int errorCount = 0;
            
            foreach (var failure in failures)
            {
                var severity = failure.GetSeverity();
                var description = failure.GetDescriptionText() ?? "Unknown";
                var failureId = failure.GetFailureDefinitionId()?.Guid.ToString() ?? "NoID";
                
                try
                {
                    if (severity == FailureSeverity.Warning)
                    {
                        // Always suppress warnings to prevent UI blocking
                        fa.DeleteWarning(failure);
                        suppressedCount++;
                        
                        // Log first occurrence of each unique warning
                        if (_suppressedWarnings.Add(description))
                        {
                            _logger?.Invoke($"[Floor {_floorName}] Suppressed warning: {Truncate(description, 100)}");
                        }
                    }
                    else if (severity == FailureSeverity.Error)
                    {
                        if (IsHarmlessError(description))
                        {
                            // Treat harmless errors as warnings
                            fa.DeleteWarning(failure);
                            suppressedCount++;
                            _logger?.Invoke($"[Floor {_floorName}] Handled harmless error as warning: {Truncate(description, 100)}");
                        }
                        else
                        {
                            errorCount++;
                            _logger?.Invoke($"[Floor {_floorName}] CRITICAL ERROR (will fail): {Truncate(description, 100)}");
                            // Don't delete - let it propagate to outer exception handler
                        }
                    }
                    else if (severity == FailureSeverity.DocumentCorruption)
                    {
                        // Document corruption is serious - log and continue but don't suppress
                        _logger?.Invoke($"[Floor {_floorName}] ⚠️ DOCUMENT CORRUPTION WARNING: {Truncate(description, 100)}");
                        errorCount++;
                    }
                }
                catch (Exception ex)
                {
                    // If we fail to process a failure, log it but don't crash
                    _logger?.Invoke($"[Floor {_floorName}] Error processing failure: {ex.Message}");
                }
            }
            
            // Log summary if anything was suppressed
            if (suppressedCount > 0 || errorCount > 0)
            {
                _logger?.Invoke($"[Floor {_floorName}] Failure summary: {suppressedCount} suppressed, {errorCount} critical");
            }
            
            // Continue if no critical errors, otherwise proceed and let outer handler deal with it
            return FailureProcessingResult.Continue;
        }
        
        /// <summary>
        /// Determines if an error is known to be harmless and can be safely ignored
        /// </summary>
        private bool IsHarmlessError(string description)
        {
            if (string.IsNullOrWhiteSpace(description))
                return true; // Empty errors are harmless
                
            var lowerDesc = description.ToLowerInvariant();
            
            // Check against known harmless patterns
            foreach (var pattern in HarmlessErrorPatterns)
            {
                if (lowerDesc.Contains(pattern))
                    return true;
            }
            
            // Check if it's related to our domain (sleeves/openings)
            // These are typically modeling issues, not crashes
            foreach (var pattern in AlwaysSuppressPatterns)
            {
                if (lowerDesc.Contains(pattern))
                    return true;
            }
            
            return false;
        }
        
        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength) + "...";
        }
    }
}
