using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Interface for individual sleeve pre-calculation service.
    /// Enables dependency injection and testability.
    /// </summary>
    public interface IIndividualSleevePreCalculationService
    {
        /// <summary>
        /// Pre-calculate placement data for all zones using 3-tier parallel processing.
        /// </summary>
        /// <param name="clashZones">List of clash zones to pre-calculate</param>
        /// <param name="doc">Revit document (for validation only, not mutated)</param>
        /// <returns>Dictionary mapping zone ID to pre-calculation result</returns>
        Dictionary<Guid, SleevePreCalculationResult> PreCalculateAllZones(
            List<ClashZone> clashZones,
            Document doc);
        
        /// <summary>
        /// Clear all thread-local caches. Call at end of placement batch.
        /// </summary>
        void ClearCaches();
    }
}
