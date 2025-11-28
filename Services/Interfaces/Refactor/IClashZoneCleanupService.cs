using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Clash zone cleanup operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - cleanup only.
    /// 
    /// This interface abstracts cleanup operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (SRP - single responsibility)
    /// - Preserved cleanup logic (invalid + duplicates)
    /// </summary>
    public interface IClashZoneCleanupService
    {
        /// <summary>
        /// Remove invalid clash zones (null elements) and duplicates.
        /// 
        /// ✅ PRESERVES:
        /// - Step 1: Remove invalid clash zones (null elements)
        /// - Step 2: Remove duplicates (keep first occurrence)
        /// - Fail-safe error handling
        /// </summary>
        /// <param name="clashZones">List of clash zones to clean up</param>
        /// <param name="document">Revit document for element verification</param>
        /// <returns>Number of clash zones removed</returns>
        int CleanupInvalidClashZones(
            List<ClashZone> clashZones,
            Document document);
    }
}

