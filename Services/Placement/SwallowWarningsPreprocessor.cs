using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Preprocessor that swallows all warnings to speed up transaction commits.
    /// Used during bulk placement to prevent Revit from pausing or processing minor warnings.
    /// </summary>
    public class SwallowWarningsPreprocessor : IFailuresPreprocessor
    {
        public int WarningsDeleted { get; private set; }
        public int ErrorsFound { get; private set; }

        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            IList<FailureMessageAccessor> failures = failuresAccessor.GetFailureMessages();
            if (failures.Count == 0)
            {
                return FailureProcessingResult.Continue;
            }

            bool hasError = false;
            foreach (FailureMessageAccessor failure in failures)
            {
                FailureSeverity severity = failure.GetSeverity();
                if (severity == FailureSeverity.Warning)
                {
                    // Specifically target intersection warnings to suppress them without user popups
                    var failId = failure.GetFailureDefinitionId();
                    if (failId == BuiltInFailures.OverlapFailures.WallRoomSeparationOverlap ||
                        failId == BuiltInFailures.OverlapFailures.SpaceSeparationLinesOverlap ||
                        failId == BuiltInFailures.OverlapFailures.DuplicateInstances)
                    {
                        failuresAccessor.DeleteWarning(failure);
                        WarningsDeleted++;
                    }
                    else
                    {
                        // Swallow all other warnings too
                        failuresAccessor.DeleteWarning(failure);
                        WarningsDeleted++;
                    }
                }
                else if (severity == FailureSeverity.Error)
                {
                    hasError = true;
                    ErrorsFound++;
                }
            }

            if (hasError)
            {
                return FailureProcessingResult.ProceedWithRollBack;
            }

            return FailureProcessingResult.Continue;
        }
    }
}
