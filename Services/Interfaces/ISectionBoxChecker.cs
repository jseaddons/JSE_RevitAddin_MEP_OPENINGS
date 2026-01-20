using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for checking if clash zones are visible in the current section box.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface ISectionBoxChecker
    {
        /// <summary>
        /// Check if a clash zone is visible in the current 3D section box.
        /// </summary>
        bool IsClashZoneVisibleInCurrentSectionBox(Document doc, ClashZone clashZone, string logPrefix);
    }
}

