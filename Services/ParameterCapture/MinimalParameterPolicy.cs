using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Responsibility: Define minimal parameters for refresh (no learned keys, fast capture).
    /// 
    /// Different from SnapshotParameterPolicy - only captures essential parameters needed
    /// during refresh/clash detection phase.
    /// </summary>
    public class MinimalParameterPolicy : IParameterPolicy
    {
        // Minimal set for refresh - only what's needed for clash detection and sizing
        private static readonly HashSet<string> MINIMAL_PARAMETERS = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            // Critical for sizing and identification
            "System Name", "System Abbreviation", "System Type",
            "Size", "Width", "Height", "Diameter",
            "Reference Level", "Schedule of Level", "Schedule Level",
            
            // Critical for categorization
            "Service Type", // Cable Trays
            "System Classification"
        };

        // Same must-capture keys as snapshot policy
        private static readonly HashSet<string> MUST_CAPTURE_KEYS = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "System Type",
            "System Name",
            "System Abbreviation",
            "Reference Level",
            "Schedule of Level",
            "Schedule Level"
        };

        public HashSet<string> GetWhitelist()
        {
            return new HashSet<string>(MINIMAL_PARAMETERS, StringComparer.OrdinalIgnoreCase);
        }

        public HashSet<string> GetMustCaptureKeys()
        {
            return new HashSet<string>(MUST_CAPTURE_KEYS, StringComparer.OrdinalIgnoreCase);
        }

        public bool ShouldCapture(string parameterName, Element element)
        {
            return MUST_CAPTURE_KEYS.Contains(parameterName) ||
                   MINIMAL_PARAMETERS.Contains(parameterName);
        }

        public string MapParameterName(string requestedName, Element element)
        {
            // Same Cable Tray mapping as snapshot policy
            bool isCableTray = element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_CableTray ||
                              element.Category?.Name?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                              element is CableTray;

            if (isCableTray && requestedName.Equals("System Type", StringComparison.OrdinalIgnoreCase))
            {
                return "Service Type";
            }

            return requestedName;
        }

        public int GetMaxParameterLimit()
        {
            return 15; // Smaller limit for refresh (faster)
        }
    }
}
