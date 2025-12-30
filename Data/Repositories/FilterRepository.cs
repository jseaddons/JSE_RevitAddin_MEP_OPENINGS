using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.ErrorHandling;

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
            // ✅ NORMALIZE: Remove category suffixes from filter name to prevent duplicates
            // This ensures "Plumbing_pipes" and "Plumbing" both resolve to "Plumbing"
            filterName = FilterNameHelper.NormalizeBaseName(filterName, filterName, category);
            
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return -1;

            try
            {
                // ✅ LOG: SELECT operation
                DatabaseOperationLogger.LogSelect(
                    "Filters",
                    $"FilterName='{filterName}' AND Category='{category}'");

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
                        DatabaseOperationLogger.LogSelect(
                            "Filters",
                            $"FilterName='{filterName}' AND Category='{category}'",
                            resultCount: 1,
                            sampleRow: new Dictionary<string, object> { { "FilterId", filterId } });
                        return filterId;
                    }
                }

                // ✅ LOG: INSERT operation
                var insertParams = new Dictionary<string, object>
                {
                    { "FilterName", filterName },
                    { "Category", category },
                    { "IsFilterComboNew", 0 }
                };

                DatabaseOperationLogger.LogOperation(
                    "INSERT",
                    "Filters",
                    insertParams,
                    additionalInfo: "Creating new filter");

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
                        DatabaseOperationLogger.LogOperation(
                            "INSERT",
                            "Filters",
                            insertParams,
                            rowsAffected: 1,
                            additionalInfo: $"✅ Created FilterId={insertedId}");
                        _logger($"[SQLite] ✅ Registered filter '{filterName}' (Category='{category}') in database (FilterId={insertedId}).");
                        return insertedId;
                    }
                }

                _logger($"[SQLite] ⚠️ Failed to register filter '{filterName}' (Category='{category}').");
                return -1;
            }
            catch (Exception ex)
            {
                throw new FilterOperationException(
                    filterName,
                    category,
                    "EnsureFilter",
                    $"Failed to ensure filter exists: {ex.Message}",
                    ex);
            }
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
        /// ✅ PHASE 2: Saves filter UI state (SelectedHostCategories, OpeningSettings, and file/category lists) to database
        /// This method is PROTECTED - removing or modifying it will cause UI state to be lost
        /// Called from: FilterManagementService.SaveFilter, SaveFilterAuto, CreateFilter, CopyFilter
        /// </summary>
        public void SaveFilterUIState(string filterName, string category, List<string> selectedHostCategories, OpeningSettings openingSettings, 
            List<string> selectedMepCategoryNames = null, List<string> selectedReferenceFiles = null, List<string> selectedHostFiles = null)
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
                    throw new FilterOperationException(
                        filterName,
                        category,
                        "SaveFilterUIState",
                        $"Filter not found (FilterId={filterId})");
                }

                // ⚠️⚠️⚠️ PROTECTED CODE: This UPDATE statement saves UI state - DO NOT MODIFY ⚠️⚠️⚠️
                // Removing or modifying this SQL will cause UI state (SelectedHostCategories, AdoptToDocumentFlag) to be lost
                // ✅ STANDARDIZED: Using SelectedHostCategories only (removed SelectedHostCategories duplicate)
                // ✅ NORMALIZE: Store as comma-separated string (not JSON) for readability - user wants "Floors" not ["Floors"]
                var hostCategoriesJson = selectedHostCategories != null && selectedHostCategories.Count > 0
                    ? string.Join(", ", selectedHostCategories.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim()))
                    : null;
                
                // ✅ SCHEMA REDESIGN: Store only AdoptToDocument flag (0 or 1), not entire OpeningSettings object
                var adoptToDocumentFlag = openingSettings?.AdoptToDocument == true ? 1 : 0;
                
                _logger($"[SQLite-DEBUG] SaveFilterUIState prep: FilterId={filterId}, HostCatsJson='{hostCategoriesJson}', AdoptToDocumentFlag={adoptToDocumentFlag}");

                // ✅ LOG: UPDATE operation with all parameters
                var updateParams = new Dictionary<string, object>
                {
                    { "FilterId", filterId },
                    { "SelectedHostCategories", hostCategoriesJson },
                    { "AdoptToDocumentFlag", adoptToDocumentFlag }
                };

                DatabaseOperationLogger.LogOperation(
                    "UPDATE",
                    "Filters",
                    updateParams,
                    additionalInfo: $"Saving UI state for FilterId={filterId}");

                using (var cmd = _context.Connection.CreateCommand())
                {
                    // ✅ MIGRATION: Populate JSON array columns from existing DocKey values
                    // If SelectedReferenceFiles is not provided but ReferenceDocKey exists, use it
                    var referenceFilesJson = selectedReferenceFiles != null && selectedReferenceFiles.Count > 0
                        ? JsonSerializer.Serialize(selectedReferenceFiles)
                        : null;
                    var hostFilesJson = selectedHostFiles != null && selectedHostFiles.Count > 0
                        ? JsonSerializer.Serialize(selectedHostFiles)
                        : null;
                    var mepCategoriesJson = selectedMepCategoryNames != null && selectedMepCategoryNames.Count > 0
                        ? JsonSerializer.Serialize(selectedMepCategoryNames)
                        : null;
                    
                    cmd.CommandText = @"
                        UPDATE Filters
                        SET SelectedHostCategories = @SelectedHostCategories,
                            AdoptToDocumentFlag = @AdoptToDocumentFlag,
                            SelectedMepCategoryNames = @SelectedMepCategoryNames,
                            SelectedReferenceFiles = @SelectedReferenceFiles,
                            SelectedHostFiles = @SelectedHostFiles,
                            UpdatedAt = CURRENT_TIMESTAMP
                        WHERE FilterId = @FilterId";
                    cmd.Parameters.AddWithValue("@FilterId", filterId);
                    cmd.Parameters.AddWithValue("@SelectedHostCategories", hostCategoriesJson);
                    cmd.Parameters.AddWithValue("@AdoptToDocumentFlag", adoptToDocumentFlag);
                    cmd.Parameters.AddWithValue("@SelectedMepCategoryNames", mepCategoriesJson ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@SelectedReferenceFiles", referenceFilesJson ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@SelectedHostFiles", hostFilesJson ?? (object)DBNull.Value);

                    _logger($"[SQLite-DEBUG] About to execute UPDATE with AdoptToDocumentFlag={adoptToDocumentFlag}");
                    var affected = cmd.ExecuteNonQuery();
                    _logger($"[SQLite-DEBUG] UPDATE result: {affected} rows affected");
                    if (affected > 0)
                    {
                        DatabaseOperationLogger.LogOperation(
                            "UPDATE",
                            "Filters",
                            updateParams,
                            rowsAffected: affected,
                            additionalInfo: $"✅ Saved UI state - HostCategories: {selectedHostCategories?.Count ?? 0}, AdoptToDocumentFlag: {adoptToDocumentFlag}");
                        _logger($"[SQLite] ✅ Saved UI state for filter '{filterName}' (Category='{category}') - HostCategories: {selectedHostCategories?.Count ?? 0}, AdoptToDocumentFlag: {adoptToDocumentFlag}");
                    }
                    else
                    {
                        _logger($"[SQLite] ⚠️ SaveFilterUIState: No rows updated for filter '{filterName}' (Category='{category}') - FilterId={filterId}");
                        throw new FilterOperationException(
                            filterName,
                            category,
                            "SaveFilterUIState",
                            $"No rows updated (FilterId={filterId})");
                    }
                }
            }
            catch (FilterOperationException)
            {
                throw; // Re-throw custom exceptions
            }
            catch (Exception ex)
            {
                // ⚠️ CRITICAL: Log error but don't throw - UI state save failure is non-blocking
                _logger($"[SQLite] ❌ Error saving filter UI state: {ex.Message}");
                throw new FilterOperationException(
                    filterName,
                    category,
                    "SaveFilterUIState",
                    $"Unexpected error: {ex.Message}",
                    ex);
            }
        }

        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE METHOD - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
        /// ✅ PHASE 2: Loads filter UI state (SelectedHostCategories and OpeningSettings) from database
        /// This method is PROTECTED - removing or modifying it will cause UI state to not be restored
        /// Called from: FilterManagementService.CreateFilterFromCurrentUIState, LoadFilterFromXmlFile
        /// ✅ STANDARDIZED: Using SelectedHostCategories only (removed SelectedHostCategories duplicate)
        /// </summary>
        public (List<string> SelectedHostCategories, OpeningSettings OpeningSettings, 
            List<string> SelectedMepCategoryNames, List<string> SelectedReferenceFiles, List<string> SelectedHostFiles) 
            LoadFilterUIState(string filterName, string category)
        {
            // ⚠️⚠️⚠️ PROTECTED METHOD: DO NOT REMOVE OR MODIFY THIS VALIDATION ⚠️⚠️⚠️
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return (new List<string>(), null, new List<string>(), new List<string>(), new List<string>());

            _logger($"[SQLite-DEBUG] LoadFilterUIState called for filter='{filterName}', category='{category}'");

            try
            {
                    // ⚠️⚠️⚠️ PROTECTED CODE: This SELECT statement loads UI state - DO NOT MODIFY ⚠️⚠️⚠️
                    // Removing or modifying this SQL will cause UI state (SelectedHostCategories, AdoptToDocumentFlag, file/category lists) to not be restored
                    // NOTE: Handles both old database (with OpeningSettings) and new database (with AdoptToDocumentFlag) during migration
                    using (var cmd = _context.Connection.CreateCommand())
                    {
                        // ✅ TRY NEW SCHEMA FIRST (AdoptToDocumentFlag)
                        cmd.CommandText = @"
                        SELECT SelectedHostCategories, AdoptToDocumentFlag, SelectedMepCategoryNames, SelectedReferenceFiles, SelectedHostFiles
                        FROM Filters
                        WHERE FilterName = @FilterName AND Category = @Category";
                        cmd.Parameters.AddWithValue("@FilterName", filterName);
                        cmd.Parameters.AddWithValue("@Category", category);

                    try
                    {
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var hostCategoriesJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                                var adoptToDocumentFlag = reader.IsDBNull(1) ? 1 : reader.GetInt32(1);
                                
                                _logger($"[SQLite-DEBUG] Row found - hostCategoriesJson='{hostCategoriesJson}', adoptToDocumentFlag={adoptToDocumentFlag}");
                                _logger($"[SQLite-DEBUG] Column 2 (SelectedMepCategoryNames): {(reader.IsDBNull(2) ? "NULL" : reader.GetString(2))}");
                                _logger($"[SQLite-DEBUG] Column 3 (SelectedReferenceFiles): {(reader.IsDBNull(3) ? "NULL" : reader.GetString(3))}");
                                _logger($"[SQLite-DEBUG] Column 4 (SelectedHostFiles): {(reader.IsDBNull(4) ? "NULL" : reader.GetString(4))}");

                            List<string> hostCategories = new List<string>();
                            if (!string.IsNullOrWhiteSpace(hostCategoriesJson))
                            {
                                try
                                {
                                    // ✅ NORMALIZE: Handle both comma-separated string (new format) and JSON array (old format)
                                    if (hostCategoriesJson.StartsWith("[") && hostCategoriesJson.EndsWith("]"))
                                    {
                                        // Old JSON format - deserialize
                                        hostCategories = JsonSerializer.Deserialize<List<string>>(hostCategoriesJson) ?? new List<string>();
                                    }
                                    else
                                    {
                                        // New comma-separated format - split by comma
                                        hostCategories = hostCategoriesJson.Split(',')
                                            .Select(c => c.Trim())
                                            .Where(c => !string.IsNullOrWhiteSpace(c))
                                            .ToList();
                                    }
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to parse SelectedHostCategories for filter '{filterName}'");
                                }
                            }

                            // ✅ SCHEMA REDESIGN: Build minimal OpeningSettings with only AdoptToDocument flag
                            OpeningSettings settings = new OpeningSettings
                            {
                                AdoptToDocument = adoptToDocumentFlag == 1
                            };

                            // ✅ CRITICAL FIX: Load full file and category lists from database
                            List<string> mepCategories = new List<string>();
                            List<string> referenceFiles = new List<string>();
                            List<string> hostFiles = new List<string>();
                            
                            try
                            {
                                var mepCategoriesJson = reader.IsDBNull(2) ? null : reader.GetString(2);
                                if (!string.IsNullOrWhiteSpace(mepCategoriesJson))
                                {
                                    mepCategories = JsonSerializer.Deserialize<List<string>>(mepCategoriesJson) ?? new List<string>();
                                }
                            }
                            catch
                            {
                                _logger($"[SQLite] ⚠️ Failed to deserialize SelectedMepCategoryNames for filter '{filterName}'");
                            }
                            
                            try
                            {
                                var referenceFilesJson = reader.IsDBNull(3) ? null : reader.GetString(3);
                                if (!string.IsNullOrWhiteSpace(referenceFilesJson))
                                {
                                    referenceFiles = JsonSerializer.Deserialize<List<string>>(referenceFilesJson) ?? new List<string>();
                                }
                            }
                            catch
                            {
                                _logger($"[SQLite] ⚠️ Failed to deserialize SelectedReferenceFiles for filter '{filterName}'");
                            }
                            
                            try
                            {
                                var hostFilesJson = reader.IsDBNull(4) ? null : reader.GetString(4);
                                if (!string.IsNullOrWhiteSpace(hostFilesJson))
                                {
                                    hostFiles = JsonSerializer.Deserialize<List<string>>(hostFilesJson) ?? new List<string>();
                                }
                            }
                            catch
                            {
                                _logger($"[SQLite] ⚠️ Failed to deserialize SelectedHostFiles for filter '{filterName}'");
                            }

                            // ✅ CRITICAL: If JSON arrays are empty, try fallback to DocKey columns (pre-migration data)
                            if (mepCategories.Count == 0 || referenceFiles.Count == 0 || hostFiles.Count == 0)
                            {
                                try
                                {
                                    cmd.Parameters.Clear();
                                    cmd.CommandText = @"
                                    SELECT ReferenceCategory, ReferenceDocKey, HostDocKey
                                    FROM Filters
                                    WHERE FilterName = @FilterName AND Category = @Category";
                                    cmd.Parameters.AddWithValue("@FilterName", filterName);
                                    cmd.Parameters.AddWithValue("@Category", category);

                                    using (var fallbackReader = cmd.ExecuteReader())
                                    {
                                        if (fallbackReader.Read())
                                        {
                                            // Fallback to ReferenceCategory for MEP categories
                                            if (mepCategories.Count == 0)
                                            {
                                                var refCat = fallbackReader.IsDBNull(0) ? null : fallbackReader.GetString(0);
                                                if (!string.IsNullOrWhiteSpace(refCat))
                                                {
                                                    mepCategories = new List<string> { refCat };
                                                    _logger($"[SQLite] ℹ️ Populated MEP categories from legacy ReferenceCategory: {refCat}");
                                                }
                                            }

                                            // Fallback to ReferenceDocKey for reference files
                                            if (referenceFiles.Count == 0)
                                            {
                                                var refDoc = fallbackReader.IsDBNull(1) ? null : fallbackReader.GetString(1);
                                                if (!string.IsNullOrWhiteSpace(refDoc))
                                                {
                                                    referenceFiles = new List<string> { refDoc };
                                                    _logger($"[SQLite] ℹ️ Populated reference files from legacy ReferenceDocKey: {refDoc}");
                                                }
                                            }

                                            // Fallback to HostDocKey for host files
                                            if (hostFiles.Count == 0)
                                            {
                                                var hostDoc = fallbackReader.IsDBNull(2) ? null : fallbackReader.GetString(2);
                                                if (!string.IsNullOrWhiteSpace(hostDoc))
                                                {
                                                    hostFiles = new List<string> { hostDoc };
                                                    _logger($"[SQLite] ℹ️ Populated host files from legacy HostDocKey: {hostDoc}");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception fallbackEx)
                                {
                                    _logger($"[SQLite] ⚠️ Fallback to DocKey columns failed: {fallbackEx.Message}");
                                }
                            }

                            return (hostCategories, settings, mepCategories, referenceFiles, hostFiles);
                        }
                        else
                        {
                            _logger($"[SQLite-DEBUG] ⚠️ No row found for filter='{filterName}', category='{category}'");
                            // No rows found - return empty/default values
                            return (new List<string>(), null, new List<string>(), new List<string>(), new List<string>());
                        }
                    }
                    }
                    catch (System.Data.SQLite.SQLiteException colEx) when (colEx.Message.Contains("AdoptToDocumentFlag"))
                    {
                        // ✅ MIGRATION FALLBACK: AdoptToDocumentFlag column doesn't exist yet
                        // Try loading from old OpeningSettings column for backward compatibility
                        _logger($"[SQLite] ⚠️ AdoptToDocumentFlag column not found - trying legacy OpeningSettings column");
                        
                        cmd.Parameters.Clear();
                        cmd.CommandText = @"
                        SELECT SelectedHostCategories, OpeningSettings, SelectedMepCategoryNames, SelectedReferenceFiles, SelectedHostFiles
                        FROM Filters
                        WHERE FilterName = @FilterName AND Category = @Category";
                        cmd.Parameters.AddWithValue("@FilterName", filterName);
                        cmd.Parameters.AddWithValue("@Category", category);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                var hostCategoriesJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                                var settingsJson = reader.IsDBNull(1) ? null : reader.GetString(1);

                                List<string> hostCategories = new List<string>();
                                if (!string.IsNullOrWhiteSpace(hostCategoriesJson))
                                {
                                    try
                                    {
                                        // ✅ NORMALIZE: Handle both comma-separated string (new format) and JSON array (old format)
                                        if (hostCategoriesJson.StartsWith("[") && hostCategoriesJson.EndsWith("]"))
                                        {
                                            // Old JSON format - deserialize
                                            hostCategories = JsonSerializer.Deserialize<List<string>>(hostCategoriesJson) ?? new List<string>();
                                        }
                                        else
                                        {
                                            // New comma-separated format - split by comma
                                            hostCategories = hostCategoriesJson.Split(',')
                                                .Select(c => c.Trim())
                                                .Where(c => !string.IsNullOrWhiteSpace(c))
                                                .ToList();
                                        }
                                    }
                                    catch
                                    {
                                        _logger($"[SQLite] ⚠️ Failed to parse SelectedHostCategories for filter '{filterName}'");
                                    }
                                }

                                // ✅ LEGACY: Load OpeningSettings from old column
                                OpeningSettings settings = null;
                                if (!string.IsNullOrWhiteSpace(settingsJson))
                                {
                                    try
                                    {
                                        settings = JsonSerializer.Deserialize<OpeningSettings>(settingsJson);
                                    }
                                    catch
                                    {
                                        // Create minimal settings if deserialization fails
                                        settings = new OpeningSettings { AdoptToDocument = true };
                                        _logger($"[SQLite] ⚠️ Failed to deserialize OpeningSettings for filter '{filterName}' - using default");
                                    }
                                }
                                else
                                {
                                    // Create default settings if none found
                                    settings = new OpeningSettings { AdoptToDocument = true };
                                }

                                // ✅ CRITICAL FIX: Load full file and category lists from database
                                List<string> mepCategories = new List<string>();
                                List<string> referenceFiles = new List<string>();
                                List<string> hostFiles = new List<string>();
                                
                                try
                                {
                                    var mepCategoriesJson = reader.IsDBNull(2) ? null : reader.GetString(2);
                                    if (!string.IsNullOrWhiteSpace(mepCategoriesJson))
                                    {
                                        mepCategories = JsonSerializer.Deserialize<List<string>>(mepCategoriesJson) ?? new List<string>();
                                    }
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to deserialize SelectedMepCategoryNames for filter '{filterName}'");
                                }
                                
                                try
                                {
                                    var referenceFilesJson = reader.IsDBNull(3) ? null : reader.GetString(3);
                                    if (!string.IsNullOrWhiteSpace(referenceFilesJson))
                                    {
                                        referenceFiles = JsonSerializer.Deserialize<List<string>>(referenceFilesJson) ?? new List<string>();
                                    }
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to deserialize SelectedReferenceFiles for filter '{filterName}'");
                                }
                                
                                try
                                {
                                    var hostFilesJson = reader.IsDBNull(4) ? null : reader.GetString(4);
                                    if (!string.IsNullOrWhiteSpace(hostFilesJson))
                                    {
                                        hostFiles = JsonSerializer.Deserialize<List<string>>(hostFilesJson) ?? new List<string>();
                                    }
                                }
                                catch
                                {
                                    _logger($"[SQLite] ⚠️ Failed to deserialize SelectedHostFiles for filter '{filterName}'");
                                }

                                // ✅ CRITICAL: If JSON arrays are empty, try fallback to DocKey columns (pre-migration data)
                                if (mepCategories.Count == 0 || referenceFiles.Count == 0 || hostFiles.Count == 0)
                                {
                                    try
                                    {
                                        var fallbackCmd = _context.Connection.CreateCommand();
                                        fallbackCmd.CommandText = @"
                                        SELECT ReferenceCategory, ReferenceDocKey, HostDocKey
                                        FROM Filters
                                        WHERE FilterName = @FilterName AND Category = @Category";
                                        fallbackCmd.Parameters.AddWithValue("@FilterName", filterName);
                                        fallbackCmd.Parameters.AddWithValue("@Category", category);

                                        using (var fallbackReader = fallbackCmd.ExecuteReader())
                                        {
                                            if (fallbackReader.Read())
                                            {
                                                // Fallback to ReferenceCategory for MEP categories
                                                if (mepCategories.Count == 0)
                                                {
                                                    var refCat = fallbackReader.IsDBNull(0) ? null : fallbackReader.GetString(0);
                                                    if (!string.IsNullOrWhiteSpace(refCat))
                                                    {
                                                        mepCategories = new List<string> { refCat };
                                                        _logger($"[SQLite] ℹ️ Populated MEP categories from legacy ReferenceCategory: {refCat}");
                                                    }
                                                }

                                                // Fallback to ReferenceDocKey for reference files
                                                if (referenceFiles.Count == 0)
                                                {
                                                    var refDoc = fallbackReader.IsDBNull(1) ? null : fallbackReader.GetString(1);
                                                    if (!string.IsNullOrWhiteSpace(refDoc))
                                                    {
                                                        referenceFiles = new List<string> { refDoc };
                                                        _logger($"[SQLite] ℹ️ Populated reference files from legacy ReferenceDocKey: {refDoc}");
                                                    }
                                                }

                                                // Fallback to HostDocKey for host files
                                                if (hostFiles.Count == 0)
                                                {
                                                    var hostDoc = fallbackReader.IsDBNull(2) ? null : fallbackReader.GetString(2);
                                                    if (!string.IsNullOrWhiteSpace(hostDoc))
                                                    {
                                                        hostFiles = new List<string> { hostDoc };
                                                        _logger($"[SQLite] ℹ️ Populated host files from legacy HostDocKey: {hostDoc}");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception fallbackEx)
                                    {
                                        _logger($"[SQLite] ⚠️ Fallback to DocKey columns failed: {fallbackEx.Message}");
                                    }
                                }

                                _logger($"[SQLite] ✅ Loaded UI state using legacy OpeningSettings column");
                                return (hostCategories, settings, mepCategories, referenceFiles, hostFiles);
                            }
                            else
                            {
                                // No rows found - return empty/default values
                                return (new List<string>(), null, new List<string>(), new List<string>(), new List<string>());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error loading filter UI state: {ex.Message}");
                return (new List<string>(), null, new List<string>(), new List<string>(), new List<string>());
            }
        }
    }
}

