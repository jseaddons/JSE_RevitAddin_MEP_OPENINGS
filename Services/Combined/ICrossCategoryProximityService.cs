using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined
{
    /// <summary>
    /// Interface for cross-category proximity detection service.
    /// Detects groups of sleeves from different categories that are within proximity threshold.
    /// </summary>
    public interface ICrossCategoryProximityService
    {
        /// <summary>
        /// Detects proximity groups from a list of unified sleeves.
        /// Only creates groups for sleeves from different categories (cross-category).
        /// </summary>
        /// <param name="sleeves">List of all sleeves to analyze</param>
        /// <param name="proximityThreshold">Maximum distance (in feet) for sleeves to be considered in proximity</param>
        /// <returns>List of proximity groups (each containing 2+ sleeves from different categories)</returns>
        List<ProximityGroup> DetectProximityGroups(
            List<UnifiedSleeve> sleeves,
            double proximityThreshold);
    }
}
