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
        /// Priority: MSFD > MSD > MD > Motorized > Standard
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
            
            if (typeNameUpper.Contains("MD") && !typeNameUpper.Contains("MSD")) // Avoid matching MSD as MD
                return "MD";
            
            if (typeNameUpper.Contains("MOTORIZED"))
                return "Motorized";
            
            return "Standard";
        }

        /// <summary>
        /// Determines if the damper type requires MEP-side clearance.
        /// Non-standard types (MSFD, MSD, MD, Motorized) require MEP-side clearance.
        /// </summary>
        public bool RequiresMepSideClearance(string damperType)
        {
            if (string.IsNullOrWhiteSpace(damperType))
                return false;

            string typeUpper = damperType.Trim().ToUpperInvariant();
            
            return typeUpper == "MSFD" ||
                   typeUpper == "MSD" ||
                   typeUpper == "MD" ||
                   typeUpper == "MOTORIZED";
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

