using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for determining sleeve placement path (PATH 1, PATH 2, PATH 3).
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface IPathDeterminer
    {
        /// <summary>
        /// Determine the placement path based on filter state and clearance settings.
        /// </summary>
        SleevePlacementPath DeterminePath(
            Document doc,
            string filterName,
            string category,
            string logPrefix,
            Dictionary<string, double>? clearanceSettings);
    }
}

