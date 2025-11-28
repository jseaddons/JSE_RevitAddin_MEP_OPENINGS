using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Clash zone filtering operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - filtering only.
    /// 
    /// This interface abstracts filtering operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (SRP - single responsibility)
    /// - Preserved filtering logic with SectionBoxHelper reuse
    /// </summary>
    public interface IClashZoneFilterService
    {
        /// <summary>
        /// Filter clash zones by current selection parameters.
        /// Uses SectionBoxHelper for section box filtering.
        /// 
        /// ✅ PRESERVES:
        /// - Section box filtering (SectionBoxHelper reuse)
        /// - Reference file filtering
        /// - Clearance settings filtering
        /// - Prefix filtering
        /// - Fail-safe error handling
        /// </summary>
        /// <param name="clashZones">List of clash zones to filter</param>
        /// <param name="selectedReferenceFiles">List of selected reference files</param>
        /// <param name="currentClearanceSettings">Current clearance settings by category</param>
        /// <param name="currentPrefix">Current prefix for filter names</param>
        /// <param name="document">Revit document</param>
        /// <returns>Filtered list of clash zones</returns>
        List<ClashZone> FilterClashZonesByCurrentSelection(
            List<ClashZone> clashZones,
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document);
    }
}

