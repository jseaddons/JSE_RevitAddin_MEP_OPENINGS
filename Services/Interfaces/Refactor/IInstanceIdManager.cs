using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team F: Instance ID management for sleeves (separated from flag management).
    /// SOLID: Single Responsibility - instance ID operations only.
    /// 
    /// This interface abstracts instance ID management, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (SRP - single responsibility)
    /// - Preserved optimizations (batch collection, HashSet lookups)
    /// </summary>
    public interface IInstanceIdManager
    {
        /// <summary>
        /// Resets instance IDs for deleted sleeves.
        /// Uses batch collection optimization (collect all sleeves once for all categories).
        /// 
        /// ✅ PRESERVES:
        /// - Batch collection (1 Revit API call instead of N calls per category)
        /// - Category filtering in memory (O(1) lookup via HashSet)
        /// - SectionBoxHelper reuse for section box filtering
        /// - Database-first approach with XML fallback
        /// 
        /// Performance: Reduces N Revit API calls (one per category) to just 1 call,
        /// then filters by category in memory (much faster).
        /// </summary>
        /// <param name="categories">List of MEP categories to process</param>
        /// <param name="clashZonesByCategory">Optional dictionary of clash zones by category (for optimization)</param>
        /// <param name="refreshLogName">Optional refresh log file name for detailed logging</param>
        /// <returns>Number of instance IDs reset</returns>
        int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null);
    }
}

