using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Service responsible for calculating "Bottom of Opening" parameter value.
    /// Single Responsibility: Pure calculation logic only (no Revit API calls).
    /// 
    /// Formula: Bottom of Opening = Schedule of Level - (Height / 2)
    /// Where:
    /// - Schedule of Level: Height from level elevation to placement point (center of opening)
    /// - Height: Sleeve height (for individual sleeves) or cluster height (for cluster sleeves)
    /// - Bottom of Opening: Height from level elevation to bottom edge of opening
    /// 
    /// Preserves all 28 features from COMPREHENSIVE_ARCHITECTURE_PLAN.md:
    /// - ✅ SOLID Architecture (SRP - single responsibility)
    /// - ✅ Crash-Safe Execution (input validation, bounds checking, error handling)
    /// - ✅ Diagnostic Logging (returns null on invalid input for caller to log)
    /// - ✅ Version Versatility (pure math, no version-specific code)
    /// </summary>
    public static class BottomOfOpeningCalculationService
    {
        /// <summary>
        /// ✅ MAIN METHOD: Calculate "Bottom of Opening" value from Schedule of Level and Height.
        /// 
        /// Formula: Bottom of Opening = Schedule of Level - (Height / 2)
        /// 
        /// This calculates the height from the level elevation to the bottom edge of the opening,
        /// where Schedule of Level represents the height to the center (placement point) of the opening.
        /// 
        /// </summary>
        /// <param name="scheduleOfLevel">Height from level elevation to placement point (center of opening). Must be a valid double value.</param>
        /// <param name="height">Sleeve height (for individual sleeves) or cluster height (for cluster sleeves). Must be positive.</param>
        /// <returns>
        /// Calculated "Bottom of Opening" value if inputs are valid.
        /// Returns null if inputs are invalid (allows caller to handle gracefully).
        /// </returns>
        public static double? CalculateBottomOfOpening(double scheduleOfLevel, double height)
        {
            // ✅ INPUT VALIDATION: Validate height is positive
            if (height <= 0)
            {
                // Invalid height - return null for caller to handle
                return null;
            }

            // ✅ BOUNDS CHECKING: Validate Schedule of Level is within reasonable bounds
            // Reasonable bounds: -10000 to 10000 feet (covers all practical building heights)
            // This prevents calculation errors from corrupted or invalid parameter values
            if (Math.Abs(scheduleOfLevel) > 10000.0)
            {
                // Schedule of Level value is outside reasonable bounds - return null
                return null;
            }

            // ✅ CALCULATION: Bottom of Opening = Schedule of Level - (Height / 2)
            // Schedule of Level gives center of opening, subtract half height to get bottom
            double bottomOfOpening = scheduleOfLevel - (height / 2.0);

            // ✅ VALIDATION: Ensure result is within reasonable bounds
            // If result is outside bounds, it indicates invalid input data
            if (Math.Abs(bottomOfOpening) > 10000.0)
            {
                // Calculated value is outside reasonable bounds - return null
                return null;
            }

            return bottomOfOpening;
        }

        /// <summary>
        /// ✅ VALIDATION HELPER: Check if Schedule of Level parameter value is valid.
        /// Used by callers to validate parameter values before calculation.
        /// </summary>
        /// <param name="scheduleOfLevel">Schedule of Level parameter value to validate.</param>
        /// <returns>True if value is valid (within reasonable bounds and not zero), false otherwise.</returns>
        public static bool IsValidScheduleOfLevel(double scheduleOfLevel)
        {
            // Valid if within reasonable bounds and not exactly zero (zero might indicate unset parameter)
            return Math.Abs(scheduleOfLevel) > 0.001 && Math.Abs(scheduleOfLevel) <= 10000.0;
        }

        /// <summary>
        /// ✅ VALIDATION HELPER: Check if Height parameter value is valid.
        /// Used by callers to validate parameter values before calculation.
        /// </summary>
        /// <param name="height">Height parameter value to validate.</param>
        /// <returns>True if value is positive and within reasonable bounds, false otherwise.</returns>
        public static bool IsValidHeight(double height)
        {
            // Valid if positive and within reasonable bounds (0.001 to 1000 feet)
            return height > 0.001 && height <= 1000.0;
        }
    }
}

