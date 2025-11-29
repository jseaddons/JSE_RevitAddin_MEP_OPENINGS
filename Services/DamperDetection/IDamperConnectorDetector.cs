using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// Interface for detecting MEP connector presence and direction for dampers.
    /// Follows Interface Segregation Principle - focused interface for connector detection only.
    /// </summary>
    public interface IDamperConnectorDetector
    {
        /// <summary>
        /// Detects if a damper has an MEP connector and returns its direction.
        /// </summary>
        /// <param name="damper">The damper FamilyInstance to check</param>
        /// <param name="useWorldCoordinates">If true, uses world coordinates (for non-standard dampers). If false, uses local coordinates (for standard dampers)</param>
        /// <param name="connector">Output parameter containing the detected connector, or null if not found</param>
        /// <returns>Connector side direction ("Left", "Right", "Top", "Bottom") or empty string if not detected</returns>
        string DetectConnectorSide(FamilyInstance damper, bool useWorldCoordinates, out Connector connector);

        /// <summary>
        /// Checks if a damper has at least one MEP connector.
        /// </summary>
        bool HasMepConnector(FamilyInstance damper);
    }
}

