using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// ✅ INCREMENTAL PROCESSING SERVICE
    /// Manages filtering of sleeves based on CategoryProcessingMarkers
    /// 
    /// Usage:
    /// 1. GetNewSleevesOnly(allSleeves, category) → Returns only sleeves to process
    /// 2. After processing → MarkCategoryProcessed(category, newCount, processedIds)
    /// 
    /// Result: Each run only processes NEW sleeves added since last run
    /// </summary>
    public class IncrementalProcessingService
    {
        private readonly CategoryProcessingMarkerRepository _markerRepository;
        private readonly Action<string> _logger;

        public IncrementalProcessingService(CategoryProcessingMarkerRepository markerRepository, Action<string>? logger = null)
        {
            _markerRepository = markerRepository ?? throw new ArgumentNullException(nameof(markerRepository));
            _logger = logger ?? (_ => { });
        }

        /// <summary>
        /// ✅ Filter sleeves to only return NEW ones (not processed in previous runs)
        /// 
        /// Example:
        /// - Last run processed 10 sleeves (LastProcessedCount = 10)
        /// - Current run has 15 sleeves total
        /// - This returns only sleeves 11-15 (indices 10-14)
        /// </summary>
        public List<int> GetNewSleevesOnly(List<int> allSleeveIds, string category)
        {
            if (allSleeveIds == null || allSleeveIds.Count == 0)
                return new List<int>();

            if (string.IsNullOrWhiteSpace(category))
                return allSleeveIds;

            // Get the last processing marker
            var (lastProcessedCount, lastSleeveIds) = _markerRepository.GetMarker(category);

            // If no previous processing, return all sleeves
            if (lastProcessedCount == 0)
            {
                _logger?.Invoke($"[INCREMENTAL] 🆕 Category '{category}': First time processing, will process all {allSleeveIds.Count} sleeves");
                return allSleeveIds;
            }

            // If current count <= last count, nothing new to process
            if (allSleeveIds.Count <= lastProcessedCount)
            {
                _logger?.Invoke($"[INCREMENTAL] ✅ Category '{category}': No new sleeves (current={allSleeveIds.Count}, last processed={lastProcessedCount})");
                return new List<int>();
            }

            // Return only NEW sleeves (from lastProcessedCount onwards)
            // Example: If lastProcessedCount=10, return sleeves from index 10 onwards
            var newSleeves = allSleeveIds.Skip(lastProcessedCount).ToList();
            
            _logger?.Invoke($"[INCREMENTAL] ⏭️ Category '{category}': Processing NEW sleeves only: {newSleeves.Count} new sleeves (total={allSleeveIds.Count}, last processed={lastProcessedCount})");
            _logger?.Invoke($"[INCREMENTAL] 📍 New sleeve IDs: {string.Join(", ", newSleeves.Take(5))}{ (newSleeves.Count > 5 ? "..." : "")}");

            return newSleeves;
        }

        /// <summary>
        /// ✅ Mark a category as processed up to the given count
        /// Call this AFTER successfully processing sleeves for the category
        /// </summary>
        public bool MarkCategoryProcessed(string category, int newCount, List<int> processedSleeveIds = null)
        {
            if (string.IsNullOrWhiteSpace(category))
                return false;

            // Serialize sleeve IDs as comma-separated string for storage
            string sleeveIdsJson = processedSleeveIds != null && processedSleeveIds.Count > 0
                ? string.Join(",", processedSleeveIds.Take(100)) // Store first 100 IDs to avoid bloat
                : null;

            bool success = _markerRepository.UpdateMarker(category, newCount, sleeveIdsJson);
            
            if (success)
            {
                _logger?.Invoke($"[INCREMENTAL] ✅ Marked category '{category}' as processed: count={newCount}, sleeveCount={processedSleeveIds?.Count ?? 0}");
            }
            else
            {
                _logger?.Invoke($"[INCREMENTAL] ⚠️ Failed to mark category '{category}' as processed");
            }

            return success;
        }

        /// <summary>
        /// ✅ Reset category marker (used when category needs full reprocessing)
        /// </summary>
        public bool ResetCategoryMarker(string category)
        {
            bool success = _markerRepository.ResetMarker(category);
            
            if (success)
            {
                _logger?.Invoke($"[INCREMENTAL] 🔄 Reset category '{category}' - will reprocess all sleeves on next run");
            }
            else
            {
                _logger?.Invoke($"[INCREMENTAL] ⚠️ Failed to reset marker for category '{category}'");
            }

            return success;
        }

        /// <summary>
        /// ✅ Get summary of all categories and their processing state
        /// </summary>
        public string GetProcessingSummary()
        {
            var allMarkers = _markerRepository.GetAllMarkers();
            
            if (allMarkers.Count == 0)
            {
                return "[INCREMENTAL] 📊 No categories have been processed yet";
            }

            var lines = new List<string> { "[INCREMENTAL] 📊 CATEGORY PROCESSING STATE:" };
            
            foreach (var kvp in allMarkers.OrderBy(x => x.Key))
            {
                var (count, ids) = kvp.Value;
                lines.Add($"  - {kvp.Key}: LastProcessedCount={count}");
            }

            return string.Join("\n", lines);
        }
    }
}
