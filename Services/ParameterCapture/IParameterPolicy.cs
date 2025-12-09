using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Defines WHAT parameters to capture - separates policy from mechanics.
    /// Different policies (full snapshot vs minimal refresh) can be implemented.
    /// </summary>
    public interface IParameterPolicy
    {
        /// <summary>
        /// Get the complete whitelist of parameters to capture.
        /// </summary>
        HashSet<string> GetWhitelist();

        /// <summary>
        /// Get the must-capture keys that bypass parameter limits.
        /// These critical parameters (System Type, System Name, etc.) are always captured.
        /// </summary>
        HashSet<string> GetMustCaptureKeys();

        /// <summary>
        /// Determine if a parameter should be captured based on policy rules.
        /// </summary>
        bool ShouldCapture(string parameterName, Element element);

        /// <summary>
        /// Map parameter names for category-specific handling.
        /// Example: "System Type" → "Service Type" for Cable Trays
        /// </summary>
        string MapParameterName(string requestedName, Element element);

        /// <summary>
        /// Get the maximum number of parameters to capture (safety limit).
        /// </summary>
        int GetMaxParameterLimit();
    }
}
