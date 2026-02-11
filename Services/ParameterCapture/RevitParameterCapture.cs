using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Responsibility: Read parameters from Revit elements with fallbacks and diagnostics.
    /// 
    /// Features Preserved (28-point compliance):
    /// - Category-specific fallbacks (Cable Trays, Ducts, Pipes)
    /// - Built-in parameter fallbacks
    /// - Type parameter fallbacks
    /// - Diagnostic logging
    /// - Deployment mode support
    /// - String interning for memory optimization
    /// - Value truncation for long parameters
    /// </summary>
    public class RevitParameterCapture : IParameterCapture
    {
        /// <summary>
        /// Look up a parameter with comprehensive fallbacks.
        /// Preserves FEATURE 28: Built-in parameter fallbacks for all categories.
        /// </summary>
        public Parameter LookupParameter(Element element, string parameterName)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
                return null;

            // Try direct lookup first
            var param = element.LookupParameter(parameterName);
            if (param != null) return param;

            // Try type parameter fallback (FEATURE 7: Type parameter support)
            if (element is FamilyInstance fi)
            {
                param = fi.Symbol?.LookupParameter(parameterName);
                if (param != null) return param;
            }

            // Category-specific fallbacks
            return LookupParameterWithCategoryFallbacks(element, parameterName);
        }

        /// <summary>
        /// Category-specific fallbacks for critical parameters.
        /// Preserves FEATURE 12: Cable Tray Service Type mapping.
        /// Preserves FEATURE 13: System Type/Name/Abbreviation fallbacks.
        /// Preserves FEATURE 14: Schedule of Level fallbacks.
        /// </summary>
        private Parameter LookupParameterWithCategoryFallbacks(Element element, string parameterName)
        {
            bool isCableTray = element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_CableTray ||
                              element.Category?.Name?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                              element is CableTray;

            // FEATURE 12: Cable Tray - System Type → Service Type mapping
            if (isCableTray && parameterName.Equals("System Type", StringComparison.OrdinalIgnoreCase))
            {
                return element.LookupParameter("Service Type") ??
                       element.LookupParameter("MEP Service Type");
            }

            // FEATURE 13: System Type fallbacks (Ducts, Pipes)
            if (parameterName.Equals("System Type", StringComparison.OrdinalIgnoreCase))
            {
                return element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) ??
                       element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                       element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM) ??
                       element.LookupParameter("MEP System Type") ??
                       element.LookupParameter("System Classification");
            }

            // FEATURE 13: System Name fallbacks
            if (parameterName.Equals("System Name", StringComparison.OrdinalIgnoreCase))
            {
                return element.LookupParameter("MEP System Name") ??
                       element.LookupParameter("System Name");
            }

            // FEATURE 13: System Abbreviation fallbacks
            if (parameterName.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase))
            {
                return element.LookupParameter("System Abbreviation") ??
                       element.LookupParameter("System Abbr") ??
                       element.LookupParameter("Abbreviation") ??
                       element.LookupParameter("Abbr");
            }

            // FEATURE 14: Schedule of Level fallbacks
            if (parameterName.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                parameterName.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase))
            {
                return element.LookupParameter("Schedule of Level") ??
                       element.LookupParameter("Schedule Level") ??
                       element.LookupParameter("Elevation from Level");
            }

            // FEATURE 15: Reference Level fallbacks
            if (parameterName.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) ||
                parameterName.Equals("Level", StringComparison.OrdinalIgnoreCase))
            {
                var param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                           element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM) ??
                           element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);

                if (param == null)
                {
                    param = element.LookupParameter("Reference Level") ??
                           element.LookupParameter("Level") ??
                           element.LookupParameter("Schedule Level") ??
                           element.LookupParameter("Schedule of Level");
                }

                return param;
            }

            return null;
        }

        /// <summary>
        /// Convert parameter to string with robust handling.
        /// Preserves FEATURE 16: Double rounding to 3 decimals for memory optimization.
        /// Preserves FEATURE 17: ElementId to element name conversion.
        /// </summary>
        public string ConvertParameterToString(Element owner, Parameter p)
        {
            if (p == null) return string.Empty;

            string value = p.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            value = p.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            switch (p.StorageType)
            {
                case StorageType.Integer:
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);

                case StorageType.Double:
                    // FEATURE 16: Round doubles to 3 decimals to reduce string size
                    return Math.Round(p.AsDouble(), 3).ToString(CultureInfo.InvariantCulture);

                case StorageType.ElementId:
                    try
                    {
                        var id = p.AsElementId();
                        if (id == null) return string.Empty;

                        // FEATURE 17: Prefer element name for readability
                        var e = owner?.Document?.GetElement(id);
                        var name = e?.Name;
                        if (!string.IsNullOrWhiteSpace(name)) return name;
                        return id.GetIntegerValue().ToString(CultureInfo.InvariantCulture);
                    }
                    catch { return string.Empty; }

                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// Capture parameters using policy-driven approach.
        /// Preserves FEATURE 18: Must-capture keys bypass limits.
        /// Preserves FEATURE 19: Diagnostic logging.
        /// Preserves FEATURE 20: String interning for memory optimization.
        /// Preserves FEATURE 21: Value truncation for long parameters.
        /// </summary>
        public List<SerializableKeyValue> CaptureParameters(Element element, IParameterPolicy policy)
        {
            var result = new List<SerializableKeyValue>();
            if (element == null || policy == null) return result;

            var whitelist = policy.GetWhitelist();
            var mustCaptureKeys = policy.GetMustCaptureKeys();
            var maxLimit = policy.GetMaxParameterLimit();

            // FEATURE 18: Priority order - must-capture first, then remaining
            var orderedKeys = new List<string>();
            foreach (var k in mustCaptureKeys)
            {
                if (whitelist.Contains(k)) orderedKeys.Add(k);
            }
            foreach (var k in whitelist)
            {
                if (!mustCaptureKeys.Contains(k)) orderedKeys.Add(k);
            }

            // FEATURE 19: Diagnostic logging for duct accessories (deployment mode aware)
            bool isDuctAccessory = element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_DuctAccessory;
            if (isDuctAccessory && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Starting parameter capture\n");
            }

            foreach (var key in orderedKeys)
            {
                // FEATURE 18: Must-capture keys bypass limit
                bool isMustCapture = mustCaptureKeys.Contains(key);
                if (!isMustCapture && result.Count >= maxLimit)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[PARAM_SNAPSHOT] Parameter limit ({maxLimit}) reached for element {element.Id}");
                    break;
                }

                // Map parameter name (e.g., System Type → Service Type for Cable Trays)
                string actualKey = policy.MapParameterName(key, element);

                // Lookup parameter with fallbacks
                var param = LookupParameter(element, actualKey);

                if (param == null)
                {
                    // FEATURE 19: Log missing critical parameters
                    if (isMustCapture && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_CAPTURE] ⚠️⚠️⚠️ CRITICAL: Essential parameter '{actualKey}' (from '{key}') not found on element {element.Id} ({element.Category?.Name})\n");
                    }
                    continue;
                }

                var value = ConvertParameterToString(element, param);
                if (string.IsNullOrWhiteSpace(value))
                {
                    // FEATURE 19: Log empty critical parameters
                    if (isMustCapture && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_CAPTURE] ⚠️⚠️⚠️ CRITICAL: Essential parameter '{key}' found but empty on element {element.Id} ({element.Category?.Name})\n");
                    }
                    continue;
                }

                // FEATURE 21: Truncate long values to prevent memory bloat
                const int MAX_PARAM_VALUE_LENGTH = 200;
                if (value.Length > MAX_PARAM_VALUE_LENGTH)
                {
                    value = value.Substring(0, MAX_PARAM_VALUE_LENGTH) + "...[truncated]";
                }

                // FEATURE 20: Intern strings for memory sharing
                result.Add(new SerializableKeyValue
                {
                    Key = string.Intern(actualKey),
                    Value = string.Intern(value)
                });

                // FEATURE 19: Log captured critical parameters
                if (isMustCapture && !DeploymentConfiguration.DeploymentMode)
                {
                    string logKey = key.Equals(actualKey) ? actualKey : $"{actualKey} (from '{key}')";
                    DebugLogger.Info($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PARAM_CAPTURE] ✅ CAPTURED: Element {element.Id} ({element.Category?.Name}): '{logKey}' = '{value}'\n");
                }
            }

            return result;
        }
    }
}