using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Bulletproof parameter transfer using industry-standard pattern
    /// </summary>
    public class ParameterTransferCommand : ICommand
    {
        private readonly Document _doc;
        private readonly List<ElementId> _openingIds;
        private readonly ParameterTransferConfiguration _config;
        private readonly string _logPrefix;

        public ParameterTransferCommand(Document doc, List<ElementId> openingIds, ParameterTransferConfiguration config)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _openingIds = openingIds ?? throw new ArgumentNullException(nameof(openingIds));
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _logPrefix = "[ParameterTransferCommand]";
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"{_logPrefix} Starting parameter transfer for {_openingIds.Count} openings");

                // ---- 1. VALIDATION: Check document state ----
                if (!_doc.IsModifiable)
                {
                    var msg = "Document is read-only or workshared and not checked out";
                    DebugLogger.Error($"{_logPrefix} {msg}");
                    TaskDialog.Show("Error", msg);
                    return;
                }

                // ---- 2. READ-ONLY: Data gathering (NO transaction) ----
                var transferService = new ParameterTransferService();
                var validationResult = ValidateTransferConfiguration(_config);

                if (!validationResult.IsValid)
                {
                    TaskDialog.Show("Configuration Error", validationResult.ErrorMessage);
                    return;
                }

                DebugLogger.Info($"{_logPrefix} Configuration validated: {_config.Mappings.Count} mappings enabled");

                // ---- 3. SINGLE TRANSACTION: All parameter transfers ----
                using (var t = new Transaction(_doc, "Transfer Parameters to Openings"))
                {
                    if (t.Start() == TransactionStatus.Started)
                    {
                        // Set failure handler to auto-resolve warnings
                        var options = t.GetFailureHandlingOptions();
                        options.SetFailuresPreprocessor(new ParameterTransferWarningSwallower());
                        t.SetFailureHandlingOptions(options);

                        var result = transferService.ExecuteTransferConfigurationInTransaction(_doc, _openingIds, _config);

                        var status = t.Commit();
                        if (status == TransactionStatus.Committed)
                        {
                            ShowTransferResults(result);
                            DebugLogger.Info($"{_logPrefix} Successfully transferred {result.TransferredCount} parameters");
                        }
                        else
                        {
                            DebugLogger.Error($"{_logPrefix} Transaction failed to commit: {status}");
                            TaskDialog.Show("Error", "Failed to transfer parameters. Check log for details.");
                        }
                    }
                    else
                    {
                        DebugLogger.Error($"{_logPrefix} Failed to start transaction");
                        TaskDialog.Show("Error", "Failed to start transaction for parameter transfer.");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Exception during parameter transfer: {ex.Message}");
                DebugLogger.Error($"{_logPrefix} Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Error", $"Exception during parameter transfer: {ex.Message}");
            }
        }

        private ValidationResult ValidateTransferConfiguration(ParameterTransferConfiguration config)
        {
            var result = new ValidationResult { IsValid = true };

            if (config.Mappings.Count == 0)
            {
                result.IsValid = false;
                result.ErrorMessage = "No parameter mappings configured.";
                return result;
            }

            var enabledMappings = config.Mappings.Where(m => m.IsEnabled).ToList();
            if (enabledMappings.Count == 0)
            {
                result.IsValid = false;
                result.ErrorMessage = "No enabled parameter mappings found.";
                return result;
            }

            // Validate each mapping has required parameters
            foreach (var mapping in enabledMappings)
            {
                if (string.IsNullOrEmpty(mapping.SourceParameter) || string.IsNullOrEmpty(mapping.TargetParameter))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"Invalid mapping: Source or Target parameter is empty.";
                    break;
                }
            }

            return result;
        }

        private void ShowTransferResults(ParameterTransferResult result)
        {
            var message = $"Parameter transfer completed successfully!\n\n" +
                         $"Transferred: {result.TransferredCount}\n" +
                         $"Failed: {result.FailedCount}";

            if (result.Warnings.Count > 0)
            {
                message += $"\nWarnings: {result.Warnings.Count}";
            }

            TaskDialog.Show("Success", message);
        }
    }

    /// <summary>
    /// Validation result for configuration checking
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }

    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings during parameter transfer
    /// </summary>
    public class ParameterTransferWarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                // Only dismiss warnings, not errors
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[ParameterTransferWarningSwallower] Dismissed warning: {f.GetDescriptionText()}");
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
