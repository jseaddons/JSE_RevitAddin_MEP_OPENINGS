using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Sizing
{
    /// <summary>
    /// Service for calculating final sleeve dimensions with insulation awareness.
    /// Follows Single Responsibility Principle - only responsible for dimension calculations with insulation.
    /// Follows Open/Closed Principle - can be extended for category-specific logic without modification.
    /// </summary>
    public class InsulationAwareSizingService : IInsulationAwareSizingService
    {
        /// <summary>
        /// Calculates final width/height/diameter including insulation thickness (if applicable) and clearance.
        /// Formula: RawSize + (2 * insulationThickness) + (2 * clearance)
        /// </summary>
        public (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensions(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            bool isInsulated,
            double insulationThickness,
            double clearance)
        {
            // Calculate insulation contribution (0 if not insulated)
            double insulationContribution = isInsulated && insulationThickness > 0 
                ? (2 * insulationThickness) 
                : 0.0;

            // Calculate clearance contribution (on both sides)
            double clearanceContribution = 2 * clearance;

            // Final dimensions: Raw + Insulation + Clearance
            double finalWidth = rawWidth + insulationContribution + clearanceContribution;
            double finalHeight = rawHeight + insulationContribution + clearanceContribution;
            
            // For round elements, diameter should match width/height
            double finalDiameter = rawDiameter > 0 
                ? rawDiameter + insulationContribution + clearanceContribution
                : Math.Max(finalWidth, finalHeight);

            return (finalWidth, finalHeight, finalDiameter);
        }

        /// <summary>
        /// Calculates final width/height/diameter from ClashZone with insulation awareness.
        /// Uses ClashZone.IsInsulated and ClashZone.InsulationThickness automatically.
        /// </summary>
        public (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensionsFromClashZone(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            ClashZone clashZone,
            double clearance)
        {
            if (clashZone == null)
            {
                return CalculateFinalDimensions(rawWidth, rawHeight, rawDiameter, false, 0.0, clearance);
            }

            return CalculateFinalDimensions(
                rawWidth,
                rawHeight,
                rawDiameter,
                clashZone.IsInsulated,
                clashZone.InsulationThickness,
                clearance);
        }

        /// <summary>
        /// Calculates final rounded width/height/diameter including insulation thickness and clearance.
        /// Applies rounding based on provided settings.
        /// </summary>
        public (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensionsRounded(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            bool isInsulated,
            double insulationThickness,
            double clearance,
            double roundingValue,
            bool roundAlwaysUp)
        {
            // 1. Calculate raw dimensions with insulation and clearance
            var (width, height, diameter) = CalculateFinalDimensions(
                rawWidth, rawHeight, rawDiameter, isInsulated, insulationThickness, clearance);

            // 2. Apply rounding logic (logic copied from OpeningSettingsHelper to ensure centralization)
            double roundedWidth = RoundDimension(width, roundingValue, roundAlwaysUp);
            double roundedHeight = RoundDimension(height, roundingValue, roundAlwaysUp);
            double roundedDiameter = RoundDimension(diameter, roundingValue, roundAlwaysUp);

            return (roundedWidth, roundedHeight, roundedDiameter);
        }

        /// <summary>
        /// Calculates final rounded width/height/diameter from ClashZone including insulation thickness and clearance.
        /// Applies rounding based on provided settings.
        /// </summary>
        public (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensionsFromClashZoneRounded(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            ClashZone clashZone,
            double clearance,
            double roundingValue,
            bool roundAlwaysUp)
        {
            if (clashZone == null)
            {
                return CalculateFinalDimensionsRounded(
                    rawWidth, rawHeight, rawDiameter, false, 0.0, clearance, roundingValue, roundAlwaysUp);
            }

            return CalculateFinalDimensionsRounded(
                rawWidth,
                rawHeight,
                rawDiameter,
                clashZone.IsInsulated,
                clashZone.InsulationThickness,
                clearance,
                roundingValue,
                roundAlwaysUp);
        }

        /// <summary>
        /// Internal helper to round a single dimension (same logic as OpeningSettingsHelper but pure)
        /// </summary>
        private double RoundDimension(double dimension, double roundingValue, bool roundAlwaysUp)
        {
            if (roundingValue <= 0) return dimension;

            // Convert to millimeters (Revit internal units are feet)
            // 1 ft = 304.8 mm
            double mmDimension = dimension * 304.8;
            double roundedMm;

            if (roundAlwaysUp)
            {
                // Always round up: 453 -> 500 (with rounding value 50)
                roundedMm = Math.Ceiling(mmDimension / roundingValue) * roundingValue;
            }
            else
            {
                // Round to nearest: 453 -> 450 (with rounding value 50)
                roundedMm = Math.Round(mmDimension / roundingValue) * roundingValue;
            }

            // Convert back to internal units (feet)
            // 1 mm = 1/304.8 ft
            return roundedMm / 304.8;
        }
    }
}

