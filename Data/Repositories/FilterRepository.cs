using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository for managing filter metadata in SQLite.
    /// </summary>
    public class FilterRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public FilterRepository(SleeveDbContext context, Action<string>? logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (_ => { });
        }

        public int EnsureFilter(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return -1;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT FilterId FROM Filters
                    WHERE FilterName = @FilterName AND Category = @Category";
                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                var existingId = cmd.ExecuteScalar();
                if (existingId != null && int.TryParse(existingId.ToString(), out int filterId))
                {
                    return filterId;
                }
            }

            using (var insertCmd = _context.Connection.CreateCommand())
            {
                // ✅ FIX: IsFilterComboNew in Filters table is deprecated - flag is now in FileCombos table
                // Set to 0 (default) since we don't use it anymore
                insertCmd.CommandText = @"
                    INSERT INTO Filters (FilterName, Category, IsFilterComboNew, CreatedAt, UpdatedAt)
                    VALUES (@FilterName, @Category, 0, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
                    SELECT last_insert_rowid();";
                insertCmd.Parameters.AddWithValue("@FilterName", filterName);
                insertCmd.Parameters.AddWithValue("@Category", category);

                var newId = insertCmd.ExecuteScalar();
                if (newId != null && int.TryParse(newId.ToString(), out int insertedId))
                {
                    _logger($"[SQLite] ✅ Registered filter '{filterName}' (Category='{category}') in database.");
                    return insertedId;
                }
            }

            _logger($"[SQLite] ⚠️ Failed to register filter '{filterName}' (Category='{category}').");
            return -1;
        }

        public int GetFilterId(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return -1;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT FilterId FROM Filters
                    WHERE FilterName = @FilterName AND Category = @Category";
                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                var result = cmd.ExecuteScalar();
                if (result != null && int.TryParse(result.ToString(), out int filterId))
                {
                    return filterId;
                }
            }

            return -1;
        }

        public void UpdateFilterName(string oldName, string category, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName) || string.IsNullOrWhiteSpace(category))
                return;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE Filters
                    SET FilterName = @NewName,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE FilterName = @OldName AND Category = @Category";
                cmd.Parameters.AddWithValue("@NewName", newName);
                cmd.Parameters.AddWithValue("@OldName", oldName);
                cmd.Parameters.AddWithValue("@Category", category);

                var affected = cmd.ExecuteNonQuery();
                if (affected > 0)
                {
                    _logger($"[SQLite] ✅ Renamed filter '{oldName}' → '{newName}' (Category='{category}').");
                }
            }
        }

        public void DeleteFilter(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return;

            using (var cmd = _context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    DELETE FROM Filters
                    WHERE FilterName = @FilterName AND Category = @Category";
                cmd.Parameters.AddWithValue("@FilterName", filterName);
                cmd.Parameters.AddWithValue("@Category", category);

                var affected = cmd.ExecuteNonQuery();
                if (affected > 0)
                {
                    _logger($"[SQLite] ✅ Deleted filter '{filterName}' (Category='{category}') from database.");
                }
            }
        }
        
        /// <summary>
        /// Gets all filters from database
        /// </summary>
        public List<(string FilterName, string Category)> GetAllFilters()
        {
            var filters = new List<(string, string)>();
            
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT FilterName, Category
                        FROM Filters
                        ORDER BY FilterName, Category";
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var filterName = reader.GetString(0);
                            var category = reader.GetString(1);
                            filters.Add((filterName, category));
                        }
                    }
                }
                
                _logger($"[SQLite] ✅ Loaded {filters.Count} filters from database.");
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error loading filters from database: {ex.Message}");
            }
            
            return filters;
        }
        
        /// <summary>
        /// ⚠️ DEPRECATED: This method is no longer used - flag is now in FileCombos table
        /// Use FileCombos.IsFilterComboNew instead (reset via OpeningCommandOrchestrator.ResetFilterComboFlagAfterPlacement)
        /// Kept for backward compatibility only
        /// </summary>
        [Obsolete("IsFilterComboNew flag is now in FileCombos table. Use FileCombos.IsFilterComboNew instead.")]
        public void ResetFilterComboFlag(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return;

            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE Filters
                        SET IsFilterComboNew = 0,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE FilterName = @FilterName AND Category = @Category";
                    cmd.Parameters.AddWithValue("@FilterName", filterName);
                    cmd.Parameters.AddWithValue("@Category", category);

                    var affected = cmd.ExecuteNonQuery();
                    if (affected > 0)
                    {
                        _logger($"[SQLite] ✅ Reset IsFilterComboNew=false for filter '{filterName}' (Category='{category}') - combo marked as used");
                    }
                    else
                    {
                        _logger($"[SQLite] ⚠️ Filter '{filterName}' (Category='{category}') not found - cannot reset IsFilterComboNew flag");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error resetting IsFilterComboNew flag: {ex.Message}");
            }
        }

        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE METHOD - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
        /// ✅ PHASE 2: Saves filter UI state (SelectedHostElementTypes and OpeningSettings) to database
        /// This method is PROTECTED - removing or modifying it will cause UI state to be lost
        /// Called from: FilterManagementService.SaveFilter, SaveFilterAuto, CreateFilter, CopyFilter
        /// </summary>
        public void SaveFilterUIState(string filterName, string category, List<string> selectedHostElementTypes, OpeningSettings openingSettings)
        {
            // ⚠️⚠️⚠️ PROTECTED METHOD: DO NOT REMOVE OR MODIFY THIS VALIDATION ⚠️⚠️⚠️
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
            {
                _logger($"[SQLite] ⚠️ SaveFilterUIState: Invalid parameters - filterName or category is empty");
                return;
            }

            try
            {
                var filterId = GetFilterId(filterName, category);
                if (filterId <= 0)
                {
                    _logger($"[SQLite] ⚠️ Filter '{filterName}' (Category='{category}') not found - cannot save UI state");
                    return;
                }

                // ⚠️⚠️⚠️ PROTECTED CODE: This UPDATE statement saves UI state - DO NOT MODIFY ⚠️⚠️⚠️
                // Removing or modifying this SQL will cause UI state (SelectedHostElementTypes, OpeningSettings) to be lost
                using (var cmd = _context.Connection.CreateCommand())
                {
                    var hostTypesJson = selectedHostElementTypes != null && selectedHostElementTypes.Count > 0
                        ? JsonSerializer.Serialize(selectedHostElementTypes)
                        : "[]";
                    
                    var settingsJson = openingSettings != null
                        ? JsonSerializer.Serialize(openingSettings)
                        : null;

                    cmd.CommandText = @"
                        UPDATE Filters
                        SET SelectedHostElementTypes = @SelectedHostElementTypes,
                            OpeningSettings = @OpeningSettings,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE FilterId = @FilterId";
                    cmd.Parameters.AddWithValue("@FilterId", filterId);
                    cmd.Parameters.AddWithValue("@SelectedHostElementTypes", hostTypesJson);
                    cmd.Parameters.AddWithValue("@OpeningSettings", settingsJson ?? (object)DBNull.Value);

                    var affected = cmd.ExecuteNonQuery();
                    if (affected > 0)
                    {
                        _logger($"[SQLite] ✅ Saved UI state for filter '{filterName}' (Category='{category}')");
                    }
                    else
                    {
                        _logger($"[SQLite] ⚠️ SaveFilterUIState: No rows updated for filter '{filterName}' (Category='{category}') - FilterId={filterId}");
                    }
                }
            }
            catch (Exception ex)
            {
                // ⚠️ CRITICAL: Log error but don't throw - UI state save failure is non-blocking
                _logger($"[SQLite] ❌ Error saving filter UI state: {ex.Message}");
            }
        }

        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE METHOD - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
        /// ✅ PHASE 2: Loads filter UI state (SelectedHostElementTypes and OpeningSettings) from database
        /// This method is PROTECTED - removing or modifying it will cause UI state to not be restored
        /// Called from: FilterManagementService.CreateFilterFromCurrentUIState, LoadFilterFromXmlFile
        /// </summary>
        public (List<string> SelectedHostElementTypes, OpeningSettings OpeningSettings) LoadFilterUIState(string filterName, string category)
        {
            // ⚠️⚠️⚠️ PROTECTED METHOD: DO NOT REMOVE OR MODIFY THIS VALIDATION ⚠️⚠️⚠️
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return (new List<string>(), null);

            try
            {
                // ⚠️⚠️⚠️ PROTECTED CODE: This SELECT statement loads UI state - DO NOT MODIFY ⚠️⚠️⚠️
                // Removing or modifying this SQL will cause UI state (SelectedHostElementTypes, OpeningSettings) to not be restored
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT SelectedHostElementTypes, OpeningSettings
                        FROM Filters
                        WHERE FilterName = @FilterName AND Category = @Category";
                    cmd.Parameters.AddWithValue("@FilterName", filterName);
                    cmd.Parameters.AddWithValue("@Category", category);

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            var hostTypesJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                            var settingsJson = reader.IsDBNull(1) ? null : reader.GetString(1);

                            List<string> hostTypes = new List<string>();
                            if (!string.IsNullOrWhiteSpace(hostTypesJson) && hostTypesJson != "[]")
                            {
                                try
                                {
                                    hostTypes = JsonSerializer.Deserialize<List<string>>(hostTypesJson) ?? new List<string>();
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to deserialize SelectedHostElementTypes for filter '{filterName}'");
                                }
                            }

                            OpeningSettings settings = null;
                            if (!string.IsNullOrWhiteSpace(settingsJson))
                            {
                                try
                                {
                                    settings = JsonSerializer.Deserialize<OpeningSettings>(settingsJson);
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to deserialize OpeningSettings for filter '{filterName}'");
                                }
                            }

                            return (hostTypes, settings);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error loading filter UI state: {ex.Message}");
            }

            return (new List<string>(), null);
        }
    }
}

