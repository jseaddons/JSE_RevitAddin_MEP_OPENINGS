namespace JSE_RevitAddin_MEP_OPENINGS.Services.Sizing
{
    /// <summary>
    /// Interface for calculating final sleeve dimensions with insulation awareness.
    /// Follows Interface Segregation Principle - focused interface for sizing calculations only.
    /// </summary>
    public interface IInsulationAwareSizingService
    {
        /// <summary>
        /// Calculates final width/height/diameter including insulation thickness (if applicable) and clearance.
        /// Formula: RawSize + (2 * insulationThickness) + (2 * clearance)
        /// </summary>
        /// <param name="rawWidth">Raw width of MEP element (in Revit internal units)</param>
        /// <param name="rawHeight">Raw height of MEP element (in Revit internal units)</param>
        /// <param name="rawDiameter">Raw diameter of MEP element (in Revit internal units)</param>
        /// <param name="isInsulated">Whether the element is insulated</param>
        /// <param name="insulationThickness">Insulation thickness (in Revit internal units)</param>
        /// <param name="clearance">Clearance to add on each side (in Revit internal units)</param>
        /// <returns>Tuple of (finalWidth, finalHeight, finalDiameter)</returns>
        (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensions(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            bool isInsulated,
            double insulationThickness,
            double clearance);

        /// <summary>
        /// Calculates final width/height/diameter from ClashZone with insulation awareness.
        /// Uses ClashZone.IsInsulated and ClashZone.InsulationThickness automatically.
        /// </summary>
        /// <param name="rawWidth">Raw width of MEP element</param>
        /// <param name="rawHeight">Raw height of MEP element</param>
        /// <param name="rawDiameter">Raw diameter of MEP element</param>
        /// <param name="clashZone">ClashZone containing insulation information</param>
        /// <param name="clearance">Clearance to add on each side (in Revit internal units)</param>
        /// <returns>Tuple of (finalWidth, finalHeight, finalDiameter)</returns>
        (double finalWidth, double finalHeight, double finalDiameter) CalculateFinalDimensionsFromClashZone(
            double rawWidth,
            double rawHeight,
            double rawDiameter,
            Models.ClashZone clashZone,
            double clearance);
    }
}

