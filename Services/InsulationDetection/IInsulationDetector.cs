namespace JSE_RevitAddin_MEP_OPENINGS.Services.InsulationDetection
{
    /// <summary>
    /// Interface for detecting insulation status and thickness from MEP elements.
    /// Follows Interface Segregation Principle - focused interface for insulation detection only.
    /// </summary>
    public interface IInsulationDetector
    {
        /// <summary>
        /// Detects if an element is insulated by checking insulation parameters.
        /// Returns true if insulation thickness > threshold (0.001ft ≈ 0.3mm).
        /// </summary>
        bool IsInsulated(Autodesk.Revit.DB.Element element);

        /// <summary>
        /// Gets the insulation thickness for an element (in Revit internal units - feet).
        /// Returns 0.0 if element is not insulated or thickness cannot be determined.
        /// </summary>
        double GetInsulationThickness(Autodesk.Revit.DB.Element element);

        /// <summary>
        /// Gets insulation information from MepElementSize if available, otherwise detects from element.
        /// Returns (isInsulated, thickness) tuple.
        /// </summary>
        (bool isInsulated, double thickness) GetInsulationInfo(Autodesk.Revit.DB.Element element, Models.MepElementSize mepElementSize);
    }
}

