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
            DebugLogger.Info($"{_logPrefix} Parameter Transfer Logic is DEPRECATED and has been moved to a separate project.");
            TaskDialog.Show("Deprecated", "This functionality has been moved to a different service and is disabled in this build.");
            // Logic removed to ensure separation of concerns and resolve build errors if any dependencies were missing.
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

    // ParameterTransferWarningSwallower moved to Models/ParameterTransferWarningSwallower.cs
}
