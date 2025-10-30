using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Suppresses the "identical instances" warning when duplicates are detected in the same place.
    /// </summary>
    public class DuplicateInstanceSuppressor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();
            DebugLogger.Log($"Checking for failures to suppress: {failures.Count} messages found");

            foreach (var failure in failures)
            {
                // Log all warnings for debugging
                if (failure.GetSeverity() == FailureSeverity.Warning || failure.GetSeverity() == FailureSeverity.Error)
                {
                    string desc = failure.GetDescriptionText();
                    DebugLogger.Log($"Failure detected: {desc} (Severity: {failure.GetSeverity()})");
                }

                // Suppress warnings based on descriptions related to duplicate elements
                if (failure.GetSeverity() == FailureSeverity.Warning)
                {
                    string desc = failure.GetDescriptionText();
                    if (desc != null && (
                        desc.IndexOf("identical instances", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        desc.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        desc.IndexOf("same place", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        DebugLogger.Log($"Suppressing duplicate-related warning: {desc}");
                        failuresAccessor.DeleteWarning(failure);
                    }
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
