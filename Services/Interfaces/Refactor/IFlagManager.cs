using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team F: Core flag management operations for clash zones.
    /// SOLID: Single Responsibility - flag operations only.
    /// 
    /// This interface abstracts flag management operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (DIP - depends on abstractions)
    /// - Preserved optimizations (batch updates, HashSet lookups)
    /// </summary>
    public interface IFlagManager
    {
        /// <summary>
        /// Resets flags for deleted sleeves (database-first, then XML fallback).
        /// Preserves batch update optimization via BatchUpdateFlags().
        /// 
        /// ✅ PRESERVES:
        /// - Batch collection optimization (collect all sleeves once)
        /// - HashSet lookup optimization (O(1) instead of O(n))
        /// - Pre-loaded entries optimization (calculate once, use many times)
        /// - Flag hierarchy (cluster flags take precedence over individual flags)
        /// - Revit API verification (authoritative source, not Global XML)
        /// - SectionBoxHelper reuse for section box filtering
        /// 
        /// </summary>
        /// <param name="clashZones">List of clash zones to check (can be null - will load from database/XML)</param>
        /// <param name="categories">List of MEP categories to process</param>
        /// <param name="refreshLogName">Optional refresh log file name for detailed logging</param>
        /// <returns>Number of flags reset</returns>
        int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null);
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves (database-first).
        /// Preserves batch collection optimization (collect all sleeves once for all categories).
        /// 
        /// ✅ PRESERVES:
        /// - Batch collection (1 Revit API call instead of N calls)
        /// - Category filtering in memory (O(1) lookup via HashSet)
        /// - SectionBoxHelper reuse for section box filtering
        /// - Database-first approach with XML fallback
        /// </summary>
        /// <param name="categories">List of MEP categories to process</param>
        /// <param name="clashZonesByCategory">Optional dictionary of clash zones by category (for optimization)</param>
        /// <param name="refreshLogName">Optional refresh log file name for detailed logging</param>
        /// <returns>Number of instance IDs reset</returns>
        int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null);
        
        /// <summary>
        /// Updates flags after sleeve placement (batch update).
        /// Uses BatchUpdateFlags() for 4-6× performance improvement.
        /// 
        /// ✅ PRESERVES:
        /// - Batch flag updates (BatchUpdateFlags() - 4-6× faster)
        /// - Database-first approach
        /// </summary>
        /// <param name="placedSleeves">List of placed sleeves with clash zone ID, sleeve instance ID, and cluster flag</param>
        void UpdateFlagsAfterPlacement(
            List<(Guid clashZoneId, int sleeveInstanceId, bool isCluster)> placedSleeves);
        
        /// <summary>
        /// Batch update flags for placement (optimized version).
        /// Uses BatchUpdateFlags() for 4-6× performance improvement.
        /// 
        /// ✅ PRESERVES:
        /// - Batch flag updates (BatchUpdateFlags() - 4-6× faster)
        /// - Database-first approach
        /// </summary>
        /// <param name="clashZones">List of clash zones with sleeve IDs</param>
        /// <param name="isCluster">Whether these are cluster sleeves</param>
        /// <param name="category">MEP category name</param>
        /// <param name="filterName">Optional filter name</param>
        void BatchUpdateFlagsForPlacement(
            List<(ClashZone clashZone, int sleeveId)> clashZones, 
            bool isCluster, 
            string category, 
            string filterName = null);
    }
}

