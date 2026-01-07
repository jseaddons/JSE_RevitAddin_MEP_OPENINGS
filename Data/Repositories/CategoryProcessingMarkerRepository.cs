using System;
using System.Collections.Generic;
using System.Data.SQLite;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// ✅ INCREMENTAL PROCESSING: Repository for CategoryProcessingMarkers
    /// Tracks the last processed count per category to enable skipping already-processed sleeves
    /// 
    /// Pattern: Read marker → Process only NEW sleeves (count > LastProcessedCount) → Update marker
    /// Benefit: Next run only processes new sleeves, not from count 1
    /// </summary>
    public class CategoryProcessingMarkerRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public CategoryProcessingMarkerRepository(SleeveDbContext context, Action<string>? logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (_ => { });
        }

        /// <summary>
        /// ✅ Get the last processing marker for a category
        /// Returns (LastProcessedCount, LastProcessedSleeveIds) or (0, null) if no marker exists
        /// </summary>
        public (int lastCount, string lastSleeveIds) GetMarker(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return (0, null);

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT LastProcessedCount, LastProcessedSleeveIds
                        FROM CategoryProcessingMarkers
                        WHERE Category = @category";
                    
                    cmd.Parameters.AddWithValue("@category", category);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int lastCount = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                            string lastIds = reader.IsDBNull(1) ? null : reader.GetString(1);
                            
                            _logger?.Invoke($"[MARKER] 📍 Category '{category}': LastProcessedCount={lastCount}");
                            return (lastCount, lastIds);
                        }
                    }
                }

                _logger?.Invoke($"[MARKER] 🆕 Category '{category}': No marker exists (first run)");
                return (0, null);
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to read marker for '{category}': {ex.Message}");
                return (0, null);
            }
        }

        /// <summary>
        /// ✅ Update the processing marker for a category
        /// Sets the new LastProcessedCount and LastProcessedSleeveIds
        /// </summary>
        public bool UpdateMarker(string category, int newCount, string sleeveIds = null)
        {
            if (string.IsNullOrWhiteSpace(category))
                return false;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        INSERT OR REPLACE INTO CategoryProcessingMarkers 
                        (Category, LastProcessedCount, LastProcessedSleeveIds, MarkedAt, UpdatedAt)
                        VALUES (@category, @count, @sleeveIds, datetime('now', '+5 hours', '+30 minutes'), datetime('now', '+5 hours', '+30 minutes'))";
                    
                    cmd.Parameters.AddWithValue("@category", category);
                    cmd.Parameters.AddWithValue("@count", newCount);
                    cmd.Parameters.AddWithValue("@sleeveIds", sleeveIds ?? (object)DBNull.Value);

                    int rowsAffected = cmd.ExecuteNonQuery();
                    
                    if (rowsAffected > 0)
                    {
                        _logger?.Invoke($"[MARKER] ✅ Updated category '{category}': LastProcessedCount={newCount}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to update marker for '{category}': {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// ✅ Reset the marker for a category (used when category is cleared/refreshed)
        /// </summary>
        public bool ResetMarker(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return false;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        DELETE FROM CategoryProcessingMarkers
                        WHERE Category = @category";
                    
                    cmd.Parameters.AddWithValue("@category", category);

                    int rowsAffected = cmd.ExecuteNonQuery();
                    
                    if (rowsAffected > 0)
                    {
                        _logger?.Invoke($"[MARKER] 🔄 Reset marker for category '{category}'");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to reset marker for '{category}': {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// ✅ Get all markers for debugging/monitoring
        /// </summary>
        public Dictionary<string, (int count, string ids)> GetAllMarkers()
        {
            var markers = new Dictionary<string, (int, string)>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT Category, LastProcessedCount, LastProcessedSleeveIds
                        FROM CategoryProcessingMarkers
                        ORDER BY UpdatedAt DESC";

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string category = reader.GetString(0);
                            int count = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            string ids = reader.IsDBNull(2) ? null : reader.GetString(2);
                            
                            markers[category] = (count, ids);
                        }
                    }
                }

                _logger?.Invoke($"[MARKER] 📊 Retrieved {markers.Count} processing markers");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to read all markers: {ex.Message}");
            }

            return markers;
        }

        /// <summary>
        /// ✅ Reset all processing markers (Delete all entries)
        /// </summary>
        public void ResetAllMarkers()
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM CategoryProcessingMarkers";
                    cmd.ExecuteNonQuery();
                    _logger?.Invoke("[MARKER] ♻️ Reset all category markers successfully");
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to reset markers: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ NEW: Reset markers for a specific level only (e.g., "M|Level 1", "P|Level 1")
        /// </summary>
        public void ResetMarkersForLevel(string levelName)
        {
            if (string.IsNullOrEmpty(levelName)) return;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Delete all rows where Category ends with "|LevelName"
                    cmd.CommandText = "DELETE FROM CategoryProcessingMarkers WHERE Category LIKE @pattern";
                    cmd.Parameters.AddWithValue("@pattern", $"%|{levelName}");
                    int deleted = cmd.ExecuteNonQuery();
                    _logger?.Invoke($"[MARKER] ♻️ Reset {deleted} markers for level '{levelName}'");
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[MARKER] ⚠️ Failed to reset markers for level '{levelName}': {ex.Message}");
            }
        }
    }
}
