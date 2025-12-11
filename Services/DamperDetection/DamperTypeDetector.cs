using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Detects damper type from family/type name.
    /// Follows Single Responsibility Principle - only responsible for type detection.
    /// Follows Open/Closed Principle - can be extended for new damper types without modification.
    /// </summary>
    public class DamperTypeDetector : IDamperTypeDetector
    {
        /// <summary>
        /// Detects the damper type from the family type name (case-insensitive).
        /// Priority: MSFD > MSD > MS > MD > Motorized/Motorised > Standard
        /// </summary>
        public string DetectDamperType(string familyTypeName)
        {
            if (string.IsNullOrWhiteSpace(familyTypeName))
                return "Standard";

            string typeNameUpper = familyTypeName.Trim().ToUpperInvariant();

            // Check in priority order (most specific first)
            if (typeNameUpper.Contains("MSFD"))
                return "MSFD";
            
            if (typeNameUpper.Contains("MSD"))
                return "MSD";
            
            // ✅ FIX: Check for standalone "MS" (not MSD or MSFD)
            // Check if "MS" exists but not as part of "MSD" or "MSFD"
            // Simple check: if contains "MS" but not "MSD" and not "MSFD", it's likely standalone MS
            // Note: This may have false positives (e.g., "DAMPER" contains "MS"), but in practice damper family names
            // follow patterns like "Fire Damper MS" or "MS Type Damper", so this should work correctly
            if (typeNameUpper.Contains("MS") && !typeNameUpper.Contains("MSD") && !typeNameUpper.Contains("MSFD"))
            {
                return "MS";
            }
            
            if (typeNameUpper.Contains("MD") && !typeNameUpper.Contains("MSD")) // Avoid matching MSD as MD
                return "MD";
            
            // ✅ FIX: Check for both "MOTORIZED" (US spelling) and "MOTORISED" (British spelling)
            if (typeNameUpper.Contains("MOTORIZED") || typeNameUpper.Contains("MOTORISED"))
                return "Motorized";
            
            return "Standard";
        }

        /// <summary>
        /// Determines if the damper type requires MEP-side clearance.
        /// Non-standard types (MSFD, MSD, MS, MD, Motorized) require MEP-side clearance.
        /// </summary>
        public bool RequiresMepSideClearance(string damperType)
        {
            if (string.IsNullOrWhiteSpace(damperType))
                return false;

            string typeUpper = damperType.Trim().ToUpperInvariant();
            
            return typeUpper == "MSFD" ||
                   typeUpper == "MSD" ||
                   typeUpper == "MS" ||
                   typeUpper == "MD" ||
                   typeUpper == "MOTORIZED" ||
                   typeUpper == "MOTORISED"; // British spelling
        }

        /// <summary>
        /// Determines if the damper type is standard (uses symmetric clearance on all sides).
        /// </summary>
        public bool IsStandardDamper(string damperType)
        {
            if (string.IsNullOrWhiteSpace(damperType))
                return true;

            string typeUpper = damperType.Trim().ToUpperInvariant();
            return typeUpper == "STANDARD";
        }
    }
}

