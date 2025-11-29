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
    }
}

