using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Interface Segregation Principle (ISP)
    /// Defines HOW to read parameters from Revit elements.
    /// Separates mechanics from policy and storage.
    /// </summary>
    public interface IParameterCapture
    {
        /// <summary>
        /// Look up a parameter by name with fallbacks for built-in parameters and name variations.
        /// </summary>
        Parameter LookupParameter(Element element, string parameterName);

        /// <summary>
        /// Convert a parameter value to a robust invariant string representation.
        /// Handles ElementId, Double, Integer, String storage types.
        /// </summary>
        string ConvertParameterToString(Element owner, Parameter parameter);

        /// <summary>
        /// Capture parameters from an element using the provided policy.
        /// Returns list of key-value pairs for whitelisted parameters only.
        /// </summary>
        List<SerializableKeyValue> CaptureParameters(Element element, IParameterPolicy policy);
    }
}
