using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Interface for caching clash zone data during refresh operations.
    /// Database-first architecture - loads from SQLite, caches in memory.
    /// Eliminates redundant database queries (was 4+ loads, now 1 load per filter).
    /// </summary>
    public interface IRefreshDataCacheManager
    {
        /// <summary>
        /// Load all clash zone data from database and cache in memory.
        /// Parallelized for performance (non-Revit operations).
        /// </summary>
        /// <param name="filterNames">Filter names to load</param>
        /// <param name="categories">MEP categories to load</param>
        /// <returns>Loaded XmlCache</returns>
        XmlCache LoadAll(List<string> filterNames, List<string> categories);
        
        /// <summary>
        /// Check if file combo is already processed (O(1) lookup from cache).
        /// Used to determine Replace/Replay/FullDetection mode.
        /// </summary>
        /// <param name="cache">XmlCache to check</param>
        /// <param name="category">MEP category</param>
        /// <param name="linkedFile">Linked file key</param>
        /// <param name="hostFile">Host file key</param>
        /// <returns>True if combo is processed, false otherwise</returns>
        bool IsComboProcessed(XmlCache cache, string category, string linkedFile, string hostFile);
        
        /// <summary>
        /// Check if GUID is already resolved (O(1) lookup from cache).
        /// Used to skip already-resolved clash zones.
        /// </summary>
        /// <param name="cache">XmlCache to check</param>
        /// <param name="category">MEP category</param>
        /// <param name="guid">Clash zone GUID</param>
        /// <returns>True if GUID is resolved, false otherwise</returns>
        bool IsGuidResolved(XmlCache cache, string category, Guid guid);
        
        /// <summary>
        /// Check if all file combos are already processed (database check).
        /// Checks IsFilterComboNew flag in FileCombos table.
        /// </summary>
        /// <param name="selectedFilterNames">Filter names</param>
        /// <param name="selectedMepCategories">MEP categories</param>
        /// <param name="selectedReferenceFiles">Linked file keys</param>
        /// <param name="selectedHostFiles">Host file keys</param>
        /// <returns>True if all combos are processed, false otherwise</returns>
        bool AreAllFileCombosProcessed(
            List<string> selectedFilterNames,
            List<string> selectedMepCategories, 
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles);
        
        /// <summary>
        /// Load existing clash zones from database cache.
        /// Used by IntersectionProcessor.PrepareExistingZones().
        /// </summary>
        /// <param name="selectedFilterNames">Filter names</param>
        /// <param name="selectedMepCategories">MEP categories</param>
        /// <returns>List of existing clash zones</returns>
        List<ClashZone> LoadExistingClashZones(
            List<string> selectedFilterNames,
            List<string> selectedMepCategories);
        
        /// <summary>
        /// Update cache with new clash zones.
        /// Used by IntersectionProcessor.PostProcess().
        /// </summary>
        /// <param name="cache">XmlCache to update</param>
        /// <param name="clashZones">New clash zones to cache</param>
        void UpdateCache(XmlCache cache, List<ClashZone> clashZones);
    }
}
