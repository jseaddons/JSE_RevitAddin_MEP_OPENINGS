using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings during parameter transfer and batch operations.
    /// Shared across commands and services.
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
