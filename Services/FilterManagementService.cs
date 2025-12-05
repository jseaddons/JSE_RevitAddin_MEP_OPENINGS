using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service class for managing filter operations (New, Copy, Rename, Delete, Save, Load)
    /// </summary>
    public class FilterManagementService
    {
        private readonly Action<string> _log;
        private readonly Action<string> _updateStatus;
        private readonly Document _document;
        private List<OpeningFilter> _filters;

        public FilterManagementService(Action<string> log, Action<string> updateStatus)
        {
            _log = log;
            _updateStatus = updateStatus;
            _filters = new List<OpeningFilter>();
            _document = null; // Document not available in this constructor
        }

        public FilterManagementService(Document document, Action<string> log, Action<string> updateStatus)
        {
            _document = document;
            _log = log;
            _updateStatus = updateStatus;
            _filters = new List<OpeningFilter>();
        }

        private void Log(string message)
        {
            _log?.Invoke(message);
        }

        private void UpdateStatusSafe(string message)
        {
            _updateStatus?.Invoke(message);
        }

        /// <summary>
        /// ✅ PUBLIC: Get display category for a filter (needed for database registration)
        /// ⚠️ NO FALLBACK: Returns empty string if SelectedMepCategoryName is not set
        /// This prevents incorrect category assignment (e.g., "Ducts" when "Cable Trays" is selected)
        /// </summary>
        public string GetDisplayCategory(OpeningFilter filter)
        {
            if (filter == null)
                return string.Empty;

            if (!string.IsNullOrWhiteSpace(filter.SelectedMepCategoryName))
                return MepCategoryConstants.Normalize(filter.SelectedMepCategoryName);

            // ✅ NO FALLBACK: Return empty string instead of using filter.Category enum
            // This forces callers to get category from UI state, not from stale filter object
            return string.Empty;
        }

        /// <summary>
        /// ✅ INTERNAL: Allows access to FilterRepository for advanced operations
        /// Used by refresh service to save UI state after filter creation
        /// </summary>
        internal void UseFilterRepository(Action<FilterRepository> action)
        {
            if (_document == null || action == null)
                return;

            try
            {
                using (var context = new SleeveDbContext(_document, msg => Log($"[FILTER_MGMT][SQLite] {msg}")))
                {
                    var repository = new FilterRepository(context, msg => Log($"[FILTER_MGMT][SQLite] {msg}"));
                    action(repository);
                }
            }
            catch (Exception ex)
            {
                Log($"[FILTER_MGMT] SQLite filter operation failed: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ DATABASE-FIRST: Register filter in database (primary storage).
        /// Returns FilterId if successful, -1 if failed.
        /// </summary>
        public int RegisterFilterInDatabase(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return -1;

            int filterId = -1;
            UseFilterRepository(repo => 
            {
                filterId = repo.EnsureFilter(filterName, category);
            });
            return filterId;
        }
        
        /// <summary>
        /// ✅ CHECK FILTER SAVED: Checks if filter is saved (exists in database or XML)
        /// Returns true if filter is saved, false if it's a new/unsaved filter
        /// </summary>
        public bool IsFilterSaved(string filterName)
        {
            if (string.IsNullOrWhiteSpace(filterName))
                return false;

            try
            {
                // Check database first (primary storage)
                if (_document != null)
                {
                    int filterId = -1;
                    UseFilterRepository(repo =>
                    {
                        // Check if filter exists for any category
                        var allFilters = repo.GetAllFilters();
                        if (allFilters.Any(f => f.FilterName.Equals(filterName, StringComparison.OrdinalIgnoreCase)))
                        {
                            filterId = 1; // Filter exists
                        }
                    });
                    
                    if (filterId > 0)
                    {
                        return true; // Filter exists in database
                    }
                }
                
                // Check XML file (backward compatibility)
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                }
                
                var xmlFilePath = Path.Combine(filterDir, $"{filterName}.xml");
                if (File.Exists(xmlFilePath))
                {
                    return true; // Filter exists in XML
                }
                
                return false; // Filter not saved
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error checking if filter is saved: {ex.Message}");
                return false; // Assume not saved on error
            }
        }

        private void UpdateFilterNameInDatabase(string oldName, string category, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName) || string.IsNullOrWhiteSpace(category))
                return;

            UseFilterRepository(repo => repo.UpdateFilterName(oldName, category, newName));
        }

        private void DeleteFilterFromDatabase(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return;

            UseFilterRepository(repo => repo.DeleteFilter(filterName, category));
        }

        #region Public Methods

        /// <summary>
        /// Creates a new filter from current UI state and auto-saves it
        /// </summary>
        public void CreateNewFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Creating new filter");
                
                var filterName = GetFilterNameFromUser("New Filter", "Enter filter name:", "New Filter");
                if (string.IsNullOrEmpty(filterName)) return;

                var newFilter = CreateFilterFromCurrentUIState(filterName);
                
                // ✅ CRITICAL FIX: Get category from UI state, NOT from filter object (which may have stale/wrong category)
                // ⚠️ NO FALLBACK: Must get from UI state, not from filter object's enum property
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _log($"[FILTER_MGMT] ✅ Using category from UI state: '{categoryDisplay}' (from {selectedCategories.Count} selected categories)");
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Error getting category from UI state: {ex.Message}");
                    }
                }
                
                // ⚠️ CRITICAL: If category is empty, prompt user to select a category (don't use stale filter object data)
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _log($"[FILTER_MGMT] ❌ Category is empty - cannot create filter without category");
                    ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before creating the filter.");
                    return; // Exit early - don't create without category
                }
                
                // ✅ CRITICAL: DATABASE-FIRST - Register filter in database FIRST (before adding to UI)
                // Filter is only usable if it exists in DB, so we must create it in DB before UI
                var filterId = RegisterFilterInDatabase(newFilter.Name, categoryDisplay);
                
                if (filterId > 0)
                {
                    _log($"[FILTER_MGMT] ✅ Created filter '{filterName}' in database (FilterId={filterId})");
                    
                    // ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                    // ✅ STEP 2: Save UI state to database (captured from CreateFilterFromCurrentUIState)
                    // This ensures SelectedHostCategories and OpeningSettings are preserved when filter is loaded
                    // PROTECTED CODE: Removing this will cause UI state to be lost when filters are saved/loaded
                    try
                    {
                        UseFilterRepository(repo =>
                        {
                            // 🔍 DEBUG: Log what's being saved for new filter
                            _log($"[FILTER_MGMT] 🔍 CreateNewFilter about to call SaveFilterUIState with:");
                            _log($"[FILTER_MGMT]   Filter Name: '{newFilter.Name}'");
                            _log($"[FILTER_MGMT]   Category Display: '{categoryDisplay}'");
                            _log($"[FILTER_MGMT]   Host Categories ({newFilter.SelectedHostCategories?.Count ?? 0}): {string.Join(", ", newFilter.SelectedHostCategories ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   MEP Categories ({newFilter.SelectedMepCategoryNames?.Count ?? 0}): {string.Join(", ", newFilter.SelectedMepCategoryNames ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   Reference Files ({newFilter.SelectedReferenceFiles?.Count ?? 0}): {string.Join(", ", newFilter.SelectedReferenceFiles ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   Host Files ({newFilter.SelectedHostFiles?.Count ?? 0}): {string.Join(", ", newFilter.SelectedHostFiles ?? new List<string>())}");
                            
                            // ✅ STANDARDIZED: Use SelectedHostCategories
                            repo.SaveFilterUIState(
                                newFilter.Name,
                                categoryDisplay,
                                newFilter.SelectedHostCategories ?? new List<string>(),
                                newFilter.OpeningSettings,
                                newFilter.SelectedMepCategoryNames,    // ✅ FIXED: Pass MEP category names
                                newFilter.SelectedReferenceFiles,      // ✅ FIXED: Pass reference files
                                newFilter.SelectedHostFiles            // ✅ FIXED: Pass host files
                            );
                        });
                        _log($"[FILTER_MGMT] ✅ Saved UI state for new filter '{filterName}' to database (MepCategories: {newFilter.SelectedMepCategoryNames?.Count ?? 0}, RefFiles: {newFilter.SelectedReferenceFiles?.Count ?? 0}, HostFiles: {newFilter.SelectedHostFiles?.Count ?? 0})");
                    }
                    catch (Exception uiStateEx)
                    {
                        // ⚠️ CRITICAL: Log error but don't fail filter creation - UI state save is important but non-blocking
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not save UI state to database (non-critical): {uiStateEx.Message}");
                    }
                    
                    // ✅ CRITICAL FIX: Update filter object fields with current UI state
                    newFilter.SelectedMepCategoryName = categoryDisplay;  // ✅ CRITICAL: Update singular field for GetDisplayCategory
                    newFilter.SelectedMepCategoryNames = FilterUiStateProvider.GetSelectedMepCategoryNames?.Invoke() ?? new List<string>();
                    newFilter.SelectedReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
                    newFilter.SelectedHostFiles = FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();
                    _log($"[FILTER_MGMT] ✅ Updated filter object with UI state: Category='{categoryDisplay}', MepCats={newFilter.SelectedMepCategoryNames.Count}, RefFiles={newFilter.SelectedReferenceFiles.Count}, HostFiles={newFilter.SelectedHostFiles.Count}");
                    
                    // ✅ STEP 3: Only add to UI AFTER successful DB creation
                AddFilterToList(filterListBox, newFilter);
                    _updateStatus($"Created new filter: {filterName}");
                
                    // ✅ STEP 4: XML SECOND - Save to XML for backward compatibility (optional, non-blocking)
                    try
                    {
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }
                }
                var filePath = Path.Combine(filterDir, $"{filterName}.xml");
                SaveFilterToXmlFile(newFilter, filePath);
                        _log($"[FILTER_MGMT] ✅ Saved filter '{filterName}' to XML (backward compatibility)");
                    }
                    catch (Exception xmlEx)
                    {
                        // ✅ NON-BLOCKING: XML save failure doesn't prevent filter creation (DB is primary)
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not save filter to XML (non-critical): {xmlEx.Message}");
                    }
                }
                else
                {
                    // ✅ CRITICAL: If DB creation fails, filter is NOT added to UI and NOT usable
                    _log($"[FILTER_MGMT] ❌ Failed to create filter '{filterName}' in database - filter NOT added to UI");
                    ShowError($"Failed to create filter '{filterName}' in database. Filter cannot be used until it exists in the database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error creating new filter: {ex.Message}");
                ShowError($"Error creating new filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Copies the selected filter
        /// </summary>
        public void CopyFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Copying filter");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    ShowWarning("Please select a filter to copy.");
                    return;
                }

                var newName = GetFilterNameFromUser("Copy Filter", "Enter name for copied filter:", 
                    $"{selectedFilter.Name} Copy");
                if (string.IsNullOrEmpty(newName)) return;

                var copiedFilter = new OpeningFilter
                {
                    Name = newName,
                    Category = selectedFilter.Category,
                    OpeningType = selectedFilter.OpeningType,
                    IsEnabled = selectedFilter.IsEnabled,
                    ClashZoneStorage = selectedFilter.ClashZoneStorage,
                    LastModified = DateTime.Now,
                    // ✅ COPY UI STATE: Copy all UI state properties from source filter
                    SelectedMepCategoryNames = selectedFilter.SelectedMepCategoryNames != null ? new List<string>(selectedFilter.SelectedMepCategoryNames) : null,
                    SelectedMepCategoryName = selectedFilter.SelectedMepCategoryName,
                    SelectedReferenceFiles = selectedFilter.SelectedReferenceFiles != null ? new List<string>(selectedFilter.SelectedReferenceFiles) : null,
                    SelectedHostFiles = selectedFilter.SelectedHostFiles != null ? new List<string>(selectedFilter.SelectedHostFiles) : null,
                    SelectedHostCategories = selectedFilter.SelectedHostCategories != null ? new List<string>(selectedFilter.SelectedHostCategories) : null,
                    OpeningSettings = selectedFilter.OpeningSettings // ✅ CRITICAL: Copy OpeningSettings as well
                };

                // ✅ CRITICAL FIX: Get category from UI state OR from copied filter's SelectedMepCategoryNames
                // ⚠️ NO FALLBACK: Must get from UI state or copied filter's SelectedMepCategoryNames, not from enum
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _log($"[FILTER_MGMT] ✅ Using category from UI state: '{categoryDisplay}' (from {selectedCategories.Count} selected categories)");
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Error getting category from UI state: {ex.Message}");
                    }
                }
                
                // Fallback to copied filter's SelectedMepCategoryNames if UI state not available
                if (string.IsNullOrEmpty(categoryDisplay) && copiedFilter.SelectedMepCategoryNames != null && copiedFilter.SelectedMepCategoryNames.Count > 0)
                {
                    categoryDisplay = MepCategoryConstants.Normalize(copiedFilter.SelectedMepCategoryNames[0]);
                    _log($"[FILTER_MGMT] ✅ Using category from copied filter: '{categoryDisplay}'");
                }
                
                // ⚠️ CRITICAL: If category is empty, prompt user to select a category (don't use stale filter object data)
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _log($"[FILTER_MGMT] ❌ Category is empty - cannot copy filter without category");
                    ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before copying the filter.");
                    return; // Exit early - don't copy without category
                }

                // ✅ CRITICAL: DATABASE-FIRST - Register filter in database FIRST (before adding to UI)
                // Filter is only usable if it exists in DB, so we must create it in DB before UI
                var filterId = RegisterFilterInDatabase(copiedFilter.Name, categoryDisplay);
                
                if (filterId > 0)
                {
                    _log($"[FILTER_MGMT] ✅ Copied filter '{selectedFilter.Name}' to '{newName}' in database (FilterId={filterId})");
                    
                    // ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                    // ✅ STEP 2: Save UI state to database (copied from source filter)
                    // This ensures SelectedHostCategories and OpeningSettings are preserved when filter is copied
                    // PROTECTED CODE: Removing this will cause UI state to be lost when filters are copied
                    try
                    {
                        UseFilterRepository(repo =>
                        {
                            // ✅ STANDARDIZED: Use SelectedHostCategories
                            repo.SaveFilterUIState(
                                copiedFilter.Name,
                                categoryDisplay,
                                copiedFilter.SelectedHostCategories ?? new List<string>(),
                                copiedFilter.OpeningSettings,
                                copiedFilter.SelectedMepCategoryNames,    // ✅ FIXED: Pass MEP category names
                                copiedFilter.SelectedReferenceFiles,      // ✅ FIXED: Pass reference files
                                copiedFilter.SelectedHostFiles            // ✅ FIXED: Pass host files
                            );
                        });
                        _log($"[FILTER_MGMT] ✅ Saved UI state for copied filter '{newName}' to database (MepCategories: {copiedFilter.SelectedMepCategoryNames?.Count ?? 0}, RefFiles: {copiedFilter.SelectedReferenceFiles?.Count ?? 0}, HostFiles: {copiedFilter.SelectedHostFiles?.Count ?? 0})");
                    }
                    catch (Exception uiStateEx)
                    {
                        // ⚠️ CRITICAL: Log error but don't fail filter copy - UI state save is important but non-blocking
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not save UI state to database (non-critical): {uiStateEx.Message}");
                    }
                    
                    // ✅ STEP 3: Only add to UI AFTER successful DB creation
                AddFilterToList(filterListBox, copiedFilter);
                    _updateStatus($"Copied filter to: {newName}");
                
                    // ✅ STEP 4: XML SECOND - Save to XML for backward compatibility (optional, non-blocking)
                    try
                    {
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }
                }
                var filePath = Path.Combine(filterDir, $"{newName}.xml");
                SaveFilterToXmlFile(copiedFilter, filePath);
                        _log($"[FILTER_MGMT] ✅ Saved copied filter '{newName}' to XML (backward compatibility)");
                    }
                    catch (Exception xmlEx)
                    {
                        // ✅ NON-BLOCKING: XML save failure doesn't prevent filter creation (DB is primary)
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not save copied filter to XML (non-critical): {xmlEx.Message}");
                    }
                }
                else
                {
                    // ✅ CRITICAL: If DB creation fails, filter is NOT added to UI and NOT usable
                    _log($"[FILTER_MGMT] ❌ Failed to copy filter '{selectedFilter.Name}' to '{newName}' in database - filter NOT added to UI");
                    ShowError($"Failed to copy filter to '{newName}' in database. Filter cannot be used until it exists in the database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error copying filter: {ex.Message}");
                ShowError($"Error copying filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Renames the selected filter
        /// </summary>
        public void RenameFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Renaming filter");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    ShowWarning("Please select a filter to rename.");
                    return;
                }

                var categoryDisplay = GetDisplayCategory(selectedFilter);
                var newName = GetFilterNameFromUser("Rename Filter", "Enter new name:", selectedFilter.Name);
                if (string.IsNullOrEmpty(newName)) return;

                var oldName = selectedFilter.Name;
                selectedFilter.Name = newName;
                selectedFilter.LastModified = DateTime.Now;

                // Ensure internal list is updated
                var existing = _filters.FirstOrDefault(f => f.Name == oldName);
                if (existing != null)
                {
                    existing.Name = newName;
                    existing.LastModified = selectedFilter.LastModified;
                }
                else
                {
                    // If item wasn't tracked, start tracking it now
                    _filters.Add(selectedFilter);
                }

                // Update the ListBox item text to reflect the new name
                if (filterListBox != null)
                {
                    var idx = filterListBox.Items.IndexOf(oldName);
                    if (idx >= 0)
                    {
                        filterListBox.Items[idx] = newName;
                        filterListBox.SelectedIndex = idx; // keep selection on renamed item
                    }
                    else
                    {
                        // Fallback: refresh the whole list if we didn't find the old item
                        RefreshFilterList(filterListBox);
                    }
                }
                
                // ✅ AUTO-SAVE: Save renamed filter automatically (with new name)
                // ✅ OOP: Use ProjectPathService when document is available
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }
                }
                var newFilePath = Path.Combine(filterDir, $"{newName}.xml");
                SaveFilterToXmlFile(selectedFilter, newFilePath);
                
                // ✅ DELETE OLD FILE: Remove old file if it exists (rename operation)
                var oldFilePath = Path.Combine(filterDir, $"{oldName}.xml");
                if (File.Exists(oldFilePath) && oldFilePath != newFilePath)
                {
                    try
                    {
                        File.Delete(oldFilePath);
                        _log($"[FILTER_MGMT] Deleted old filter file: {oldFilePath}");
                    }
                    catch (Exception deleteEx)
                    {
                        _log($"[FILTER_MGMT] Warning: Could not delete old filter file: {deleteEx.Message}");
                    }
                }
                
                _log($"[FILTER_MGMT] Renamed and auto-saved filter '{oldName}' to '{newName}'");
                _updateStatus($"Renamed filter to: {newName}");

                UpdateFilterNameInDatabase(oldName, categoryDisplay, newName);
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error renaming filter: {ex.Message}");
                ShowError($"Error renaming filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Seeds default filters into the UI and internal storage.
        /// Only seeds if filter doesn't already exist (either in list or as saved file).
        /// </summary>
        public void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames)
        {
            if (filterListBox == null) return;
            
            var filterDir = GetDefaultFilterDirectory();
            bool dirExists = Directory.Exists(filterDir);
            
            foreach (var name in defaultNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                
                // Skip if already in list
                if (_filters.Any(f => f.Name == name) || filterListBox.Items.Contains(name)) continue;
                
                // ✅ PERSISTENCE CHECK: Skip if filter file already exists (don't overwrite saved filters)
                if (dirExists)
                {
                    var filePath = Path.Combine(filterDir, $"{name}.xml");
                    if (File.Exists(filePath))
                    {
                        _log($"[FILTER_MGMT] Skipping default filter '{name}' - file already exists: {filePath}");
                        continue; // Don't seed default if file exists - LoadAllSavedFilters will load it
                    }
                }
                
                var filter = CreateFilterFromCurrentUIState(name);
                AddFilterToList(filterListBox, filter);
            }
        }

        /// <summary>
        /// Deletes the selected filter and its file from disk
        /// </summary>
        public void DeleteFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Deleting filter");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    ShowWarning("Please select a filter to delete.");
                    return;
                }

                var categoryDisplay = GetDisplayCategory(selectedFilter);
                var result = MessageBox.Show(
                    $"Are you sure you want to delete the filter '{selectedFilter.Name}'?", 
                    "Confirm Delete", 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    RemoveFilterFromList(filterListBox, selectedFilter);
                    
                    // ✅ DELETE FILE: Remove filter file from disk
                    var filterDir = GetDefaultFilterDirectory();
                    var filePath = Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            File.Delete(filePath);
                            _log($"[FILTER_MGMT] Deleted filter file: {filePath}");
                        }
                        catch (Exception deleteEx)
                        {
                            _log($"[FILTER_MGMT] Warning: Could not delete filter file: {deleteEx.Message}");
                        }
                    }
                    
                    _log($"[FILTER_MGMT] Deleted filter: {selectedFilter.Name}");
                    _updateStatus($"Deleted filter: {selectedFilter.Name}");

                    DeleteFilterFromDatabase(selectedFilter.Name, categoryDisplay);
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error deleting filter: {ex.Message}");
                ShowError($"Error deleting filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Saves the selected filter to XML file automatically to destined path without showing file dialog
        /// </summary>
        public void SaveFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Saving filter (silent mode)");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    ShowWarning("Please select a filter to save.");
                    return;
                }

                // ✅ CRITICAL FIX: Collect CURRENT UI state before saving (filter object may have stale data)
                // ⚠️ CRITICAL: Collect CURRENT UI state, not stale filter data
                // This ensures the latest UI selections are persisted to the database
                var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? selectedFilter.SelectedHostCategories ?? new List<string>();
                var currentMepCategoryNames = FilterUiStateProvider.GetSelectedMepCategoryNames?.Invoke() ?? new List<string>();
                var currentReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
                var currentHostFiles = FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();
                
                // 🔍 DEBUG: Log what was collected from delegates
                _log($"[FILTER_MGMT] 🔍 SaveFilter collecting UI state:");
                _log($"[FILTER_MGMT]   Host Categories ({currentHostCategories?.Count ?? 0}): {string.Join(", ", currentHostCategories ?? new List<string>())}");
                _log($"[FILTER_MGMT]   MEP Categories ({currentMepCategoryNames?.Count ?? 0}): {string.Join(", ", currentMepCategoryNames ?? new List<string>())}");
                _log($"[FILTER_MGMT]   Reference Files ({currentReferenceFiles?.Count ?? 0}): {string.Join(", ", currentReferenceFiles ?? new List<string>())}");
                _log($"[FILTER_MGMT]   Host Files ({currentHostFiles?.Count ?? 0}): {string.Join(", ", currentHostFiles ?? new List<string>())}");
                
                // ✅ CRITICAL FIX: Get category from UI state (currently selected MEP category), not from filter object
                // The filter object may have stale/wrong category (e.g., "Ducts" when user selected "Cable Trays")
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            // ✅ CRITICAL: Use first selected category (most common case: single category selected)
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _log($"[FILTER_MGMT] ✅ Using category from UI state: '{categoryDisplay}' (from {selectedCategories.Count} selected categories: {string.Join(", ", selectedCategories)})");
                        }
                        else
                        {
                            _log($"[FILTER_MGMT] ⚠️ GetSelectedMepCategoryNames returned null or empty list");
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Error getting category from UI state: {ex.Message}");
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] ⚠️ FilterUiStateProvider.GetSelectedMepCategoryNames is null - delegate not registered");
                }
                
                // ⚠️ CRITICAL FIX: If category is empty from UI, try to use the filter's SelectedMepCategoryName as fallback
                // This handles the case where delegate returns empty but filter has a saved category
                if (string.IsNullOrEmpty(categoryDisplay) && !string.IsNullOrEmpty(selectedFilter.SelectedMepCategoryName))
                {
                    categoryDisplay = MepCategoryConstants.Normalize(selectedFilter.SelectedMepCategoryName);
                    _log($"[FILTER_MGMT] ⚠️ UI state was empty, falling back to filter's SelectedMepCategoryName: '{categoryDisplay}'");
                }
                
                // ⚠️ CRITICAL: If category is still empty, prompt user to select a category (don't use stale filter object data)
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _log($"[FILTER_MGMT] ❌ Category is empty - cannot save filter without category");
                    ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before saving the filter.");
                    return; // Exit early - don't save without category
                }
                
                // ✅ FIX: Get OpeningSettings from UI, not from stale filter object
                // OpeningSettings must be collected from UI using GetClearanceSettings delegate
                OpeningSettings currentOpeningSettings = null;
                if (FilterUiStateProvider.GetClearanceSettings != null)
                {
                    try
                    {
                        // GetClearanceSettings takes category as parameter, not filter name
                        var clearanceSettings = FilterUiStateProvider.GetClearanceSettings(categoryDisplay);
                        if (clearanceSettings != null)
                        {
                            currentOpeningSettings = new OpeningSettings
                            {
                                ClearanceSettings = clearanceSettings
                            };
                            // Also preserve SelectedMepType if available
                            if (selectedFilter.OpeningSettings != null)
                            {
                                currentOpeningSettings.SelectedMepType = selectedFilter.OpeningSettings.SelectedMepType;
                            }
                        }
                        else
                        {
                            // Fallback to filter's OpeningSettings if UI doesn't provide
                            currentOpeningSettings = selectedFilter.OpeningSettings;
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Could not get OpeningSettings from UI: {ex.Message}, using filter's OpeningSettings");
                        currentOpeningSettings = selectedFilter.OpeningSettings;
                    }
                }
                else
                {
                    // Fallback: Use filter's OpeningSettings if delegate not registered
                    currentOpeningSettings = selectedFilter.OpeningSettings;
                }
                
                // ✅ Update filter object with current UI state
                selectedFilter.SelectedHostCategories = currentHostCategories;
                
                // ✅ DATABASE-FIRST: Register filter in database FIRST (primary storage)
                // categoryDisplay is already declared above (line 596)
                var filterId = RegisterFilterInDatabase(selectedFilter.Name, categoryDisplay);
                
                if (filterId > 0)
                {
                    _log($"[FILTER_MGMT] ✅ Saved filter '{selectedFilter.Name}' in database (FilterId={filterId})");
                    
                    // ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                    // ✅ PHASE 2: Save UI state to database (using CURRENT UI state, not stale filter data)
                    // This ensures SelectedHostCategories and OpeningSettings are preserved when filter is saved
                    // PROTECTED CODE: Removing this will cause UI state to be lost when filters are saved
                    try
                    {
                        UseFilterRepository(repo =>
                        {
                            // 🔍 DEBUG: Log the parameters being passed
                            _log($"[FILTER_MGMT] 🔍 About to call SaveFilterUIState with:");
                            _log($"[FILTER_MGMT]   Filter Name: '{selectedFilter.Name}'");
                            _log($"[FILTER_MGMT]   Category Display: '{categoryDisplay}'");
                            _log($"[FILTER_MGMT]   Current Host Categories ({currentHostCategories?.Count ?? 0}): {string.Join(", ", currentHostCategories ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   Current MEP Category Names ({currentMepCategoryNames?.Count ?? 0}): {string.Join(", ", currentMepCategoryNames ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   Current Reference Files ({currentReferenceFiles?.Count ?? 0}): {string.Join(", ", currentReferenceFiles ?? new List<string>())}");
                            _log($"[FILTER_MGMT]   Current Host Files ({currentHostFiles?.Count ?? 0}): {string.Join(", ", currentHostFiles ?? new List<string>())}");
                            
                            repo.SaveFilterUIState(
                                selectedFilter.Name,
                                categoryDisplay,
                                currentHostCategories ?? new List<string>(), // Maps to SelectedHostCategories in DB
                                currentOpeningSettings,
                                currentMepCategoryNames,    // ✅ FIXED: Pass MEP category names
                                currentReferenceFiles,      // ✅ FIXED: Pass reference files
                                currentHostFiles            // ✅ FIXED: Pass host files
                            );
                        });
                        _log($"[FILTER_MGMT] ✅ Saved UI state for filter '{selectedFilter.Name}' to database (HostCategories: {currentHostCategories?.Count ?? 0}, MepCategories: {currentMepCategoryNames?.Count ?? 0}, RefFiles: {currentReferenceFiles?.Count ?? 0}, HostFiles: {currentHostFiles?.Count ?? 0})");
                    }
                    catch (Exception uiStateEx)
                    {
                        // ⚠️ CRITICAL: Log error but don't fail filter save - UI state save is important but non-blocking
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not save UI state to database (non-critical): {uiStateEx.Message}");
                    }
                    
                    // ✅ CRITICAL FIX: Update in-memory filter object with current UI state
                    // This ensures the filter object has the latest data when selected again
                    selectedFilter.OpeningSettings = currentOpeningSettings;
                    selectedFilter.SelectedHostCategories = currentHostCategories;
                    selectedFilter.SelectedMepCategoryName = categoryDisplay;  // ✅ CRITICAL: Update singular field for GetDisplayCategory
                    selectedFilter.SelectedMepCategoryNames = currentMepCategoryNames ?? new List<string>();  // ✅ CRITICAL: Update plural field
                    selectedFilter.SelectedReferenceFiles = currentReferenceFiles ?? new List<string>();  // ✅ CRITICAL: Update reference files
                    selectedFilter.SelectedHostFiles = currentHostFiles ?? new List<string>();  // ✅ CRITICAL: Update host files
                    UpdateFilterInMemory(selectedFilter.Name, selectedFilter);
                    _log($"[FILTER_MGMT] ✅ Updated in-memory filter '{selectedFilter.Name}' with current UI state (Category: '{categoryDisplay}', MepCats: {currentMepCategoryNames?.Count ?? 0}, RefFiles: {currentReferenceFiles?.Count ?? 0}, HostFiles: {currentHostFiles?.Count ?? 0})");
                    
                _updateStatus($"Saved filter: {selectedFilter.Name}");

                    // ✅ XML SECOND: Save to XML for backward compatibility (optional, non-blocking)
                    if (!DeploymentConfiguration.DisableXmlCreation)
                    {
                        try
                        {
                            var filterDir = GetDefaultFilterDirectory();
                            var filePath = System.IO.Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                            SaveFilterToXmlFile(selectedFilter, filePath);
                            _log($"[FILTER_MGMT] ✅ Saved filter '{selectedFilter.Name}' to XML (backward compatibility): {filePath}");
                        }
                        catch (Exception xmlEx)
                        {
                            // ✅ NON-BLOCKING: XML save failure doesn't prevent filter save (DB is primary)
                            _log($"[FILTER_MGMT] ⚠️ Warning: Could not save filter to XML (non-critical): {xmlEx.Message}");
                        }
                    }
                    else
                    {
                        _log($"[FILTER_MGMT] ⚠️ XML creation disabled - skipping XML save for filter '{selectedFilter.Name}'");
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] ❌ Failed to save filter '{selectedFilter.Name}' in database");
                    ShowError($"Failed to save filter '{selectedFilter.Name}' in database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error saving filter: {ex.Message}");
                ShowError($"Error saving filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Saves filter automatically to default location without user dialog
        /// </summary>
        public void SaveFilterAuto(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Auto-saving filter");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _log("[FILTER_MGMT] No filter selected for auto-save");
                    return;
                }

                // ✅ OOP: Use ProjectPathService when document is available
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }
                }

                // ✅ DATABASE-FIRST: Register filter in database FIRST (primary storage)
                // ✅ CRITICAL FIX: Get category from UI state (currently selected MEP category), not from filter object
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _log($"[FILTER_MGMT] ✅ Auto-save: Using category from UI state: '{categoryDisplay}' (from {selectedCategories.Count} categories: {string.Join(", ", selectedCategories)})");
                        }
                        else
                        {
                            _log($"[FILTER_MGMT] ⚠️ Auto-save: GetSelectedMepCategoryNames returned null or empty list");
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Auto-save: Error getting category from UI state: {ex.Message}");
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] ⚠️ Auto-save: FilterUiStateProvider.GetSelectedMepCategoryNames is null - delegate not registered");
                }
                
                // ⚠️ CRITICAL: If category is empty, skip auto-save (don't use stale filter object data)
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _log($"[FILTER_MGMT] ❌ Auto-save skipped: Category is empty - cannot save filter without category");
                    // Don't show error for auto-save (non-blocking), just skip
                    return; // Exit early - don't auto-save without category
                }
                
                var filterId = RegisterFilterInDatabase(selectedFilter.Name, categoryDisplay);
                
                if (filterId > 0)
                {
                    _log($"[FILTER_MGMT] ✅ Auto-saved filter '{selectedFilter.Name}' in database (FilterId={filterId})");
                    
                    // ✅ CRITICAL FIX: Collect CURRENT UI state before saving (filter object may have stale data)
                    // ⚠️ CRITICAL: Collect CURRENT UI state, not stale filter data
                    // This ensures the latest UI selections are persisted to the database
                                        var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();

                    
                    // ✅ Get category display name once (used for both clearance settings and database registration)
                    // categoryDisplay is already declared above (line 740)
                    
                    // ✅ FIX: Get OpeningSettings from UI, not from stale filter object
                    // OpeningSettings must be collected from UI using GetClearanceSettings delegate
                    OpeningSettings currentOpeningSettings = null;
                    if (FilterUiStateProvider.GetClearanceSettings != null)
                    {
                        try
                        {
                            // GetClearanceSettings takes category as parameter, not filter name
                            var clearanceSettings = FilterUiStateProvider.GetClearanceSettings(categoryDisplay);
                            if (clearanceSettings != null)
                            {
                                currentOpeningSettings = new OpeningSettings
                                {
                                    ClearanceSettings = clearanceSettings
                                };
                                // Also preserve SelectedMepType if available
                                if (selectedFilter.OpeningSettings != null)
                                {
                                    currentOpeningSettings.SelectedMepType = selectedFilter.OpeningSettings.SelectedMepType;
                                }
                            }
                            else
                            {
                                // Fallback to filter's OpeningSettings if UI doesn't provide
                                currentOpeningSettings = selectedFilter.OpeningSettings;
                            }
                        }
                        catch (Exception ex)
                        {
                            _log($"[FILTER_MGMT] ⚠️ Could not get OpeningSettings from UI: {ex.Message}, using filter's OpeningSettings");
                            currentOpeningSettings = selectedFilter.OpeningSettings;
                        }
                    }
                    else
                    {
                        // Fallback: Use filter's OpeningSettings if delegate not registered
                        currentOpeningSettings = selectedFilter.OpeningSettings;
                    }
                    
                    // ✅ Update filter object with current UI state
                    selectedFilter.SelectedHostCategories = currentHostCategories;
                    
                    // ⚠️⚠️⚠️ CRITICAL: UI STATE PERSISTENCE - DO NOT REMOVE OR MODIFY ⚠️⚠️⚠️
                    // ✅ PHASE 2: Save UI state to database (using CURRENT UI state, not stale filter data)
                    // This ensures SelectedHostCategories and OpeningSettings are preserved when filter is auto-saved
                    // PROTECTED CODE: Removing this will cause UI state to be lost when filters are auto-saved
                    try
                    {
                        UseFilterRepository(repo =>
                        {
                            repo.SaveFilterUIState(
                                selectedFilter.Name,
                                categoryDisplay,
                                currentHostCategories ?? new List<string>(), // Maps to SelectedHostCategories in DB
                                currentOpeningSettings
                            );
                        });
                        _log($"[FILTER_MGMT] ✅ Auto-saved UI state for filter '{selectedFilter.Name}' to database (HostCategories: {currentHostCategories?.Count ?? 0})");
                    }
                    catch (Exception uiStateEx)
                    {
                        // ⚠️ CRITICAL: Log error but don't fail filter auto-save - UI state save is important but non-blocking
                        _log($"[FILTER_MGMT] ⚠️ Warning: Could not auto-save UI state to database (non-critical): {uiStateEx.Message}");
                    }
                    
                    // ✅ CRITICAL FIX: Update in-memory filter object with current UI state
                    // This ensures the filter object has the latest data when selected again
                    selectedFilter.OpeningSettings = currentOpeningSettings;
                    selectedFilter.SelectedHostCategories = currentHostCategories;
                    UpdateFilterInMemory(selectedFilter.Name, selectedFilter);
                    _log($"[FILTER_MGMT] ✅ Updated in-memory filter '{selectedFilter.Name}' with current UI state (auto-save)");
                    
                    // ✅ XML SECOND: Save to XML for backward compatibility (optional, non-blocking)
                    if (!DeploymentConfiguration.DisableXmlCreation)
                    {
                        try
                        {
                            var filePath = Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                SaveFilterToXmlFile(selectedFilter, filePath);
                            _log($"[FILTER_MGMT] ✅ Auto-saved filter '{selectedFilter.Name}' to XML (backward compatibility): {filePath}");
                        }
                        catch (Exception xmlEx)
                        {
                            // ✅ NON-BLOCKING: XML save failure doesn't prevent filter save (DB is primary)
                            _log($"[FILTER_MGMT] ⚠️ Warning: Could not auto-save filter to XML (non-critical): {xmlEx.Message}");
                        }
                    }
                    else
                    {
                        _log($"[FILTER_MGMT] ⚠️ XML creation disabled - skipping XML auto-save for filter '{selectedFilter.Name}'");
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] ⚠️ Warning: Could not auto-save filter '{selectedFilter.Name}' in database");
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error auto-saving filter: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ DATABASE-FIRST: Loads all saved filters from database FIRST, then XML as fallback
        /// </summary>
        public void LoadAllSavedFilters(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Loading all saved filters (DATABASE-FIRST)");
                
                var loadedFilters = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int dbLoadedCount = 0;
                int xmlLoadedCount = 0;
                
                // ✅ STEP 1: Load from DATABASE FIRST (primary source)
                try
                {
                    UseFilterRepository(repo =>
                    {
                        var dbFilters = repo.GetAllFilters();
                        foreach (var (filterName, category) in dbFilters)
                        {
                            if (string.IsNullOrWhiteSpace(filterName))
                                continue;
                            
                            // Only add if not already in list (avoid duplicates)
                            if (!filterListBox.Items.Contains(filterName) && !_filters.Any(f => f.Name == filterName))
                            {
                                // ✅ FIX: Load saved UI state from database instead of using current UI state
                                var (hostCategories, openingSettings, mepCategories, refFiles, hostFiles) = repo.LoadFilterUIState(filterName, category);
                                
                                // Create filter with saved UI state
                                var filter = new OpeningFilter
                                {
                                    Name = filterName,
                                    SelectedMepCategoryName = category,
                                    SelectedHostCategories = hostCategories ?? new List<string>(),
                                    OpeningSettings = openingSettings,
                                    SelectedMepCategoryNames = mepCategories ?? new List<string>(),
                                    SelectedReferenceFiles = refFiles ?? new List<string>(),
                                    SelectedHostFiles = hostFiles ?? new List<string>(),
                                    ClashZoneStorage = new ClashZoneStorage
                                    {
                                        Filters = new List<FilterGroupForStorage>(),
                                        ClashZones = new List<ClashZone>()
                                    }
                                };
                                
                                // ✅ CRITICAL FIX: If mepCategories loaded but category empty, use first MEP category
                                if (string.IsNullOrEmpty(filter.SelectedMepCategoryName) && mepCategories != null && mepCategories.Count > 0)
                                {
                                    filter.SelectedMepCategoryName = mepCategories[0];
                                    _log($"[FILTER_MGMT] ℹ️ Category was empty, using first MEP category: '{mepCategories[0]}'");
                                }
                                
                                AddFilterToList(filterListBox, filter);
                                loadedFilters.Add(filterName);
                                dbLoadedCount++;
                                _log($"[FILTER_MGMT] ✅ Loaded filter '{filterName}' from database with saved UI state (HostCats={hostCategories?.Count ?? 0}, OpeningSettings={openingSettings != null})");
                            }
                        }
                    });
                }
                catch (Exception dbEx)
                {
                    _log($"[FILTER_MGMT] ⚠️ Warning: Could not load filters from database: {dbEx.Message}");
                }

                // ✅ STEP 2: Load from XML as fallback (for filters not in DB, or to get full filter details)
                var filterDir = GetDefaultFilterDirectory();
                if (Directory.Exists(filterDir))
                {
                var xmlFiles = Directory.GetFiles(filterDir, "*.xml");
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var filter = LoadFilterFromXmlFile(xmlFile);
                        if (filter != null)
                        {
                                // Only add if not already loaded from DB (avoid duplicates)
                                if (!loadedFilters.Contains(filter.Name) && 
                                    !filterListBox.Items.Contains(filter.Name) && 
                                    !_filters.Any(f => f.Name == filter.Name))
                            {
                                AddFilterToList(filterListBox, filter);
                                    loadedFilters.Add(filter.Name);
                                    xmlLoadedCount++;
                                    
                                    // Register in DB (so it's available for future DB-first loading)
                                RegisterFilterInDatabase(filter.Name, GetDisplayCategory(filter));
                                    _log($"[FILTER_MGMT] ✅ Loaded filter '{filter.Name}' from XML (fallback)");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] Error loading filter from {xmlFile}: {ex.Message}");
                        // Continue loading other filters even if one fails
                    }
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] Filter directory does not exist: {filterDir} (using DB-only filters)");
                }
                
                _log($"[FILTER_MGMT] ✅ Loaded {dbLoadedCount} filters from database, {xmlLoadedCount} from XML (total: {loadedFilters.Count})");
                _updateStatus($"Loaded {loadedFilters.Count} saved filters ({dbLoadedCount} from DB, {xmlLoadedCount} from XML)");
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error loading saved filters: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the default directory for storing filters
        /// </summary>
        public string GetDefaultFilterDirectory()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var projectName = GetCurrentProjectNameSafe();
            var filterDir = Path.Combine(appDataPath, "JSE_MEP_Openings", "Projects", projectName, "Filters");
            return filterDir;
        }

        private string GetCurrentProjectNameSafe()
        {
            try
            {
                // Use environment variable set by the host or fallback to process title
                var title = Environment.GetEnvironmentVariable("JSE_ACTIVE_DOC_TITLE");
                if (!string.IsNullOrWhiteSpace(title)) return NormalizeProjectName(title);
            }
            catch { }
            return "Default";
        }

        private string NormalizeProjectName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name.Trim();
        }

        /// <summary>
        /// Loads a filter automatically from default location (DATABASE-FIRST)
        /// </summary>
        public OpeningFilter LoadFilterAuto(string filterName)
        {
            try
            {
                _log($"[FILTER_MGMT] Auto-loading filter: {filterName}");
                
                // ✅ DATABASE-FIRST: Try to load UI state from DB first
                OpeningFilter loadedFilter = null;
                List<string> hostCategories = null;
                OpeningSettings openingSettings = null;
                List<string> mepCategories = null;
                List<string> refFiles = null;
                List<string> hostFiles = null;
                
                try
                {
                    UseFilterRepository(repo =>
                    {
                        // Try to load from DB using first available category
                        var allFilters = repo.GetAllFilters();
                        var filterRecord = allFilters.FirstOrDefault(f => f.FilterName.Equals(filterName, StringComparison.OrdinalIgnoreCase));
                        
                        if (filterRecord.FilterName != null)
                        {
                            var (hostCats, settings, mepCats, refF, hostF) = repo.LoadFilterUIState(filterName, filterRecord.Category);
                            hostCategories = hostCats;
                            openingSettings = settings;
                            mepCategories = mepCats;
                            refFiles = refF;
                            hostFiles = hostF;
                            _log($"[FILTER_MGMT] ✅ Loaded UI state from DB for filter '{filterName}': HostCategories={hostCategories?.Count ?? 0}, OpeningSettings={(settings != null ? "Yes" : "No")}, MepCategories={mepCats?.Count ?? 0}, RefFiles={refF?.Count ?? 0}, HostFiles={hostF?.Count ?? 0}");
                        }
                    });
                }
                catch (Exception dbEx)
                {
                    _log($"[FILTER_MGMT] ⚠️ Could not load from DB: {dbEx.Message} - falling back to XML");
                }
                
                // ✅ FALLBACK: If DB has no data or failed, load from XML
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                }
                
                var filePath = Path.Combine(filterDir, $"{filterName}.xml");
                _log($"[FILTER_MGMT] Looking for filter file at: {filePath}");
                
                if (File.Exists(filePath))
                {
                    loadedFilter = LoadFilterFromXmlFile(filePath);
                    _log($"[FILTER_MGMT] Loaded filter from XML: {filePath}");
                    
                    // ✅ DEBUG: Log what was loaded from DB
                    _log($"[FILTER_MGMT-DEBUG] DB Load Results - hostCategories: {(hostCategories == null ? "NULL" : hostCategories.Count + " items: [" + string.Join(", ", hostCategories) + "]")}");
                    _log($"[FILTER_MGMT-DEBUG] DB Load Results - mepCategories: {(mepCategories == null ? "NULL" : mepCategories.Count + " items: [" + string.Join(", ", mepCategories) + "]")}");
                    _log($"[FILTER_MGMT-DEBUG] DB Load Results - refFiles: {(refFiles == null ? "NULL" : refFiles.Count + " items: [" + string.Join(", ", refFiles) + "]")}");
                    _log($"[FILTER_MGMT-DEBUG] DB Load Results - hostFiles: {(hostFiles == null ? "NULL" : hostFiles.Count + " items: [" + string.Join(", ", hostFiles) + "]")}");
                    
                    // ✅ MERGE: Apply DB UI state on top of XML filter
                    // ✅ CRITICAL: Always assign hostCategories if loaded from DB (even if empty, to clear previous values)
                    if (hostCategories != null)
                    {
                        loadedFilter.SelectedHostCategories = hostCategories;
                        _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedHostCategories = {hostCategories.Count} items: [{string.Join(", ", hostCategories)}]");
                    }
                    if (openingSettings != null)
                    {
                        loadedFilter.OpeningSettings = openingSettings;
                        _log($"[FILTER_MGMT] ✅ Applied DB UI state: OpeningSettings updated");
                    }
                    // ✅ CRITICAL FIX: Also assign MEP categories, reference files, and host files
                    // Always assign if loaded from DB (even if empty, to clear previous values)
                    if (mepCategories != null && mepCategories.Count > 0)
                    {
                        loadedFilter.SelectedMepCategoryNames = mepCategories;
                        // ✅ CRITICAL FIX: Also set singular field so GetDisplayCategory() works
                        loadedFilter.SelectedMepCategoryName = mepCategories[0];
                        _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedMepCategoryNames = {mepCategories.Count} items: [{string.Join(", ", mepCategories)}], SelectedMepCategoryName = '{mepCategories[0]}'");
                    }
                    if (refFiles != null)
                    {
                        loadedFilter.SelectedReferenceFiles = refFiles;
                        _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedReferenceFiles = {refFiles.Count} items: [{string.Join(", ", refFiles)}]");
                    }
                    if (hostFiles != null)
                    {
                        loadedFilter.SelectedHostFiles = hostFiles;
                        _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedHostFiles = {hostFiles.Count} items: [{string.Join(", ", hostFiles)}]");
                    }
                }
                else
                {
                    _log($"[FILTER_MGMT] Filter file not found: {filePath}");
                    
                    // ✅ CRITICAL FIX: If XML doesn't exist but we have DB data, create a filter from DB data
                    if (hostCategories != null || openingSettings != null || (mepCategories != null && mepCategories.Count > 0) || 
                        (refFiles != null && refFiles.Count > 0) || (hostFiles != null && hostFiles.Count > 0))
                    {
                        _log($"[FILTER_MGMT] ⚠️ XML file not found but DB has UI state - creating filter from DB data");
                        
                        // Create a basic filter object from DB data
                        loadedFilter = new OpeningFilter
                        {
                            Name = filterName,
                            Category = Models.MepCategory.Ducts, // Default, will be updated if we can determine from category
                            IsEnabled = true,
                            LastModified = DateTime.Now
                        };
                        
                        // Apply all DB UI state
                        if (hostCategories != null)
                        {
                            loadedFilter.SelectedHostCategories = hostCategories;
                            _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedHostCategories = {hostCategories.Count} items: [{string.Join(", ", hostCategories)}]");
                        }
                        if (openingSettings != null)
                        {
                            loadedFilter.OpeningSettings = openingSettings;
                            _log($"[FILTER_MGMT] ✅ Applied DB UI state: OpeningSettings updated");
                        }
                        if (mepCategories != null && mepCategories.Count > 0)
                        {
                            loadedFilter.SelectedMepCategoryNames = mepCategories;
                            // ✅ CRITICAL FIX: Also set singular field so GetDisplayCategory() works
                            loadedFilter.SelectedMepCategoryName = mepCategories[0];
                            _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedMepCategoryNames = {mepCategories.Count} items: [{string.Join(", ", mepCategories)}], SelectedMepCategoryName = '{mepCategories[0]}'");
                        }
                        if (refFiles != null)
                        {
                            loadedFilter.SelectedReferenceFiles = refFiles;
                            _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedReferenceFiles = {refFiles.Count} items: [{string.Join(", ", refFiles)}]");
                        }
                        if (hostFiles != null)
                        {
                            loadedFilter.SelectedHostFiles = hostFiles;
                            _log($"[FILTER_MGMT] ✅ Applied DB UI state: SelectedHostFiles = {hostFiles.Count} items: [{string.Join(", ", hostFiles)}]");
                        }
                    }
                    else
                    {
                        _log($"[FILTER_MGMT] No XML file and no DB data - returning null");
                        return null; // No XML and no DB data - return null
                    }
                }

                _log($"[FILTER_MGMT] Auto-loaded filter '{filterName}' (Host categories: {loadedFilter?.SelectedHostCategories?.Count ?? 0})");
                return loadedFilter;
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error auto-loading filter '{filterName}': {ex.Message}");
                return null; // ✅ CORRECT: Return null on error - caller should handle
            }
        }

        /// <summary>
        /// Loads a filter from XML file
        /// </summary>
        public void LoadFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Loading filter");
                
                // ✅ OOP: Use ProjectPathService when document is available
                string filterDir;
                if (_document != null)
                {
                    ProjectPathService.EnsureFiltersDirectory(_document);
                    filterDir = ProjectPathService.GetFiltersDirectory(_document);
                }
                else
                {
                    filterDir = GetDefaultFilterDirectory();
                    if (!Directory.Exists(filterDir))
                    {
                        Directory.CreateDirectory(filterDir);
                    }
                }
                
                // ✅ OPTION 1: Try to load from database first (if filter exists in DB)
                // Show a dialog to let user choose: Load from DB or Load from XML file
                var choiceDialog = new System.Windows.Forms.Form
                {
                    Text = "Load Filter",
                    Size = new System.Drawing.Size(400, 150),
                    StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false
                };
                
                var label = new System.Windows.Forms.Label
                {
                    Text = "Choose how to load the filter:",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(350, 20)
                };
                
                var dbButton = new System.Windows.Forms.Button
                {
                    Text = "Load from Database",
                    Location = new System.Drawing.Point(20, 50),
                    Size = new System.Drawing.Size(160, 30),
                    DialogResult = System.Windows.Forms.DialogResult.Yes
                };
                
                var xmlButton = new System.Windows.Forms.Button
                {
                    Text = "Load from XML File",
                    Location = new System.Drawing.Point(200, 50),
                    Size = new System.Drawing.Size(160, 30),
                    DialogResult = System.Windows.Forms.DialogResult.No
                };
                
                choiceDialog.Controls.Add(label);
                choiceDialog.Controls.Add(dbButton);
                choiceDialog.Controls.Add(xmlButton);
                choiceDialog.AcceptButton = dbButton;
                choiceDialog.CancelButton = xmlButton;
                
                var choice = choiceDialog.ShowDialog();
                
                if (choice == System.Windows.Forms.DialogResult.Yes)
                {
                    // ✅ LOAD FROM DATABASE
                    try
                    {
                        UseFilterRepository(repo =>
                        {
                            var allFilters = repo.GetAllFilters();
                            if (allFilters.Count == 0)
                            {
                                MessageBox.Show("No filters found in database.", "No Filters", 
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            
                            // Show a dialog to select which filter to load
                            var filterSelectDialog = new System.Windows.Forms.Form
                            {
                                Text = "Select Filter from Database",
                                Size = new System.Drawing.Size(400, 300),
                                StartPosition = System.Windows.Forms.FormStartPosition.CenterParent,
                                FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
                                MaximizeBox = false,
                                MinimizeBox = false
                            };
                            
                            var filterList = new System.Windows.Forms.ListBox
                            {
                                Location = new System.Drawing.Point(20, 20),
                                Size = new System.Drawing.Size(340, 200)
                            };
                            
                            foreach (var filterRecord in allFilters)
                            {
                                filterList.Items.Add($"{filterRecord.FilterName} ({filterRecord.Category})");
                            }
                            
                            var loadButton = new System.Windows.Forms.Button
                            {
                                Text = "Load",
                                Location = new System.Drawing.Point(200, 230),
                                Size = new System.Drawing.Size(80, 30),
                                DialogResult = System.Windows.Forms.DialogResult.OK
                            };
                            
                            var cancelButton = new System.Windows.Forms.Button
                            {
                                Text = "Cancel",
                                Location = new System.Drawing.Point(290, 230),
                                Size = new System.Drawing.Size(80, 30),
                                DialogResult = System.Windows.Forms.DialogResult.Cancel
                            };
                            
                            filterSelectDialog.Controls.Add(filterList);
                            filterSelectDialog.Controls.Add(loadButton);
                            filterSelectDialog.Controls.Add(cancelButton);
                            filterSelectDialog.AcceptButton = loadButton;
                            filterSelectDialog.CancelButton = cancelButton;
                            
                            if (filterSelectDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && filterList.SelectedItem != null)
                            {
                                var selectedText = filterList.SelectedItem.ToString();
                                var filterName = selectedText.Split('(')[0].Trim();
                                var category = selectedText.Split('(')[1].TrimEnd(')').Trim();
                                
                                // Load filter from database
                                var loadedFilter = LoadFilterAuto(filterName);
                                if (loadedFilter != null)
                                {
                                    AddFilterToList(filterListBox, loadedFilter);
                                    try { FilterUiStateProvider.ApplyFilterToUi?.Invoke(loadedFilter); } catch { }
                                    
                                    _log($"[FILTER_MGMT] Loaded filter '{loadedFilter.Name}' from database");
                                    _updateStatus($"Loaded filter: {loadedFilter.Name}");
                                    
                                    MessageBox.Show($"Filter '{loadedFilter.Name}' loaded successfully from database!", "Success", 
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show($"Could not load filter '{filterName}' from database.", "Error", 
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        });
                    }
                    catch (Exception dbEx)
                    {
                        _log($"[FILTER_MGMT] Error loading filter from database: {dbEx.Message}");
                        MessageBox.Show($"Error loading filter from database: {dbEx.Message}", "Error", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (choice == System.Windows.Forms.DialogResult.No)
                {
                    // ✅ LOAD FROM XML FILE (original behavior)
                    var openDialog = new OpenFileDialog
                    {
                        Title = "Load Filter from XML File",
                        Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
                        DefaultExt = "xml",
                        InitialDirectory = filterDir  // Set default directory
                    };

                    if (openDialog.ShowDialog() == DialogResult.OK)
                    {
                        var loadedFilter = LoadFilterFromXmlFile(openDialog.FileName);
                        
                        if (loadedFilter != null)
                        {
                            AddFilterToList(filterListBox, loadedFilter);
                            try { FilterUiStateProvider.ApplyFilterToUi?.Invoke(loadedFilter); } catch { }
                            
                            _log($"[FILTER_MGMT] Loaded filter '{loadedFilter.Name}' from: {openDialog.FileName}");
                            _updateStatus($"Loaded filter: {loadedFilter.Name}");
                            
                            MessageBox.Show($"Filter '{loadedFilter.Name}' loaded successfully!", "Success", 
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error loading filter: {ex.Message}");
                ShowError($"Error loading filter: {ex.Message}");
            }
        }

        #endregion

        #region Private Helper Methods

        private string GetFilterNameFromUser(string title, string prompt, string defaultValue)
        {
            var inputDialog = new InputDialog(title, prompt, defaultValue);
            try
            {
                // Make sure dialog appears on top and is modal
                inputDialog.TopMost = true;
                var result = inputDialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    var filterName = inputDialog.InputText?.Trim();
                    if (string.IsNullOrEmpty(filterName))
                    {
                        ShowWarning("Filter name cannot be empty.");
                        return null;
                    }
                    return filterName;
                }
            }
            finally
            {
                inputDialog.Dispose();
            }
            return null;
        }

        /// <summary>
        /// ✅ PUBLIC: Creates a new filter from current UI state (needed for saving new filters)
        /// </summary>
        public OpeningFilter CreateFilterFromCurrentUIState(string filterName)
        {
            var filter = new OpeningFilter
            {
                Name = filterName,
                Category = Models.MepCategory.Ducts,
                OpeningType = Models.OpeningType.RectangularSleeves,
                IsEnabled = true,
                LastModified = DateTime.Now,
                ClashZoneStorage = null
            };

            // ✅ PHASE 2: Try to load UI state from database first
            try
            {
                var categoryDisplay = GetDisplayCategory(filter);
                
                // ✅ CRITICAL FIX: If GetDisplayCategory returns empty (SelectedMepCategoryName not set),
                // try to get category from SelectedMepCategoryNames list or from mepCategories
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    if (filter.SelectedMepCategoryNames != null && filter.SelectedMepCategoryNames.Count > 0)
                    {
                        categoryDisplay = MepCategoryConstants.Normalize(filter.SelectedMepCategoryNames[0]);
                        _log($"[FILTER_MGMT] ℹ️ GetDisplayCategory was empty, using first MEP category from list: '{categoryDisplay}'");
                    }
                    else
                    {
                        _log($"[FILTER_MGMT] ⚠️ Cannot load UI state - no category available (SelectedMepCategoryName and SelectedMepCategoryNames are both empty)");
                        categoryDisplay = ""; // Will cause query to return no results, but won't crash
                    }
                }
                
                UseFilterRepository(repo =>
                {
                    var (hostCategories, settings, mepCategories, refFiles, hostFiles) = repo.LoadFilterUIState(filterName, categoryDisplay);
                    if (hostCategories != null && hostCategories.Count > 0)
                    {
                        // ✅ CRITICAL FIX: Assign hostCategories to filter object (was missing!)
                        filter.SelectedHostCategories = hostCategories;
                        _log($"[FILTER_MGMT] ✅ Loaded SelectedHostCategories from database: {string.Join(", ", hostCategories)}");
                    }
                    if (settings != null)
                    {
                        filter.OpeningSettings = settings;
                        _log($"[FILTER_MGMT] ✅ Loaded OpeningSettings from database for filter '{filterName}'");
                    }
                    if (mepCategories != null && mepCategories.Count > 0)
                    {
                        filter.SelectedMepCategoryNames = mepCategories;
                        // ✅ CRITICAL FIX: Also set singular field so GetDisplayCategory() returns correct value
                        filter.SelectedMepCategoryName = mepCategories[0];
                        _log($"[FILTER_MGMT] ✅ Loaded SelectedMepCategoryNames from database: {string.Join(", ", mepCategories)}, set SelectedMepCategoryName to '{mepCategories[0]}'");
                    }
                    if (refFiles != null && refFiles.Count > 0)
                    {
                        filter.SelectedReferenceFiles = refFiles;
                        _log($"[FILTER_MGMT] ✅ Loaded SelectedReferenceFiles from database: {string.Join(", ", refFiles)}");
                    }
                    if (hostFiles != null && hostFiles.Count > 0)
                    {
                        filter.SelectedHostFiles = hostFiles;
                        _log($"[FILTER_MGMT] ✅ Loaded SelectedHostFiles from database: {string.Join(", ", hostFiles)}");
                    }
                });
            }
            catch (Exception dbEx)
            {
                _log($"[FILTER_MGMT] ⚠️ Could not load UI state from database for filter '{filterName}': {dbEx.Message}");
            }

            // Pull current UI selections via provider if available (fallback or override)
            try
            {
                var cats = FilterUiStateProvider.GetSelectedMepCategoryNames?.Invoke();
                if (cats != null && cats.Count > 0)
                {
                    filter.SelectedMepCategoryNames = new List<string>(cats);
                    filter.SelectedMepCategoryName = cats[0];
                }

                var refs = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke();
                if (refs != null) filter.SelectedReferenceFiles = new List<string>(refs);

                var hosts = FilterUiStateProvider.GetSelectedHostFiles?.Invoke();
                if (hosts != null) filter.SelectedHostFiles = new List<string>(hosts);

                // ✅ Use UI state from provider if database had no data, or if user changed it
                var hostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke();
                if (hostCategories != null && hostCategories.Count > 0)
                {
                    // Override with current UI state (user may have changed it)
                    filter.SelectedHostCategories = new List<string>(hostCategories);
                }
                else if (filter.SelectedHostCategories == null || filter.SelectedHostCategories.Count == 0)
                {
                    // No UI state from provider and no database data - initialize empty
                    filter.SelectedHostCategories = new List<string>();
                }
            }
            catch { }

            return filter;
        }

        private OpeningFilter GetSelectedFilter(ListBox filterListBox)
        {
            if (filterListBox?.SelectedItem != null)
            {
                var selectedName = filterListBox.SelectedItem.ToString();
                return _filters.FirstOrDefault(f => f.Name == selectedName) ?? 
                       new OpeningFilter
                       {
                           Name = selectedName,
                           Category = Models.MepCategory.Ducts,
                           OpeningType = Models.OpeningType.RectangularSleeves,
                           IsEnabled = true,
                           LastModified = DateTime.Now
                       };
            }
            return null;
        }

        private void AddFilterToList(ListBox filterListBox, OpeningFilter filter)
        {
            if (filterListBox == null || filter == null)
                return;

            if (IsDerivedCategoryFilterName(filterListBox, filter.Name))
            {
                _log($"[FILTER_MGMT] Skipping derived category filter '{filter.Name}' from UI list");
                return;
            }

            filterListBox.Items.Add(filter.Name);
            _filters.Add(filter);
            _log($"[FILTER_MGMT] Added filter '{filter.Name}' to list");
        }

        private void RemoveFilterFromList(ListBox filterListBox, OpeningFilter filter)
        {
            if (filterListBox != null)
            {
                filterListBox.Items.Remove(filter.Name);
                _filters.Remove(filter);
                _log($"[FILTER_MGMT] Removed filter '{filter.Name}' from list");
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL FIX: Updates the in-memory filter object to match the saved state
        /// This ensures that when a filter is selected again, it uses the updated UI state
        /// </summary>
        public void UpdateFilterInMemory(string filterName, OpeningFilter updatedFilter)
        {
            if (string.IsNullOrWhiteSpace(filterName) || updatedFilter == null)
                return;
            
            try
            {
                var existingFilter = _filters.FirstOrDefault(f => f.Name == filterName);
                if (existingFilter != null)
                {
                    // Update all properties from the updated filter
                    existingFilter.SelectedMepCategoryNames = updatedFilter.SelectedMepCategoryNames != null 
                        ? new List<string>(updatedFilter.SelectedMepCategoryNames) 
                        : null;
                    existingFilter.SelectedMepCategoryName = updatedFilter.SelectedMepCategoryName;
                    existingFilter.SelectedReferenceFiles = updatedFilter.SelectedReferenceFiles != null 
                        ? new List<string>(updatedFilter.SelectedReferenceFiles) 
                        : null;
                    existingFilter.SelectedHostFiles = updatedFilter.SelectedHostFiles != null 
                        ? new List<string>(updatedFilter.SelectedHostFiles) 
                        : null;
                    existingFilter.SelectedHostCategories = updatedFilter.SelectedHostCategories != null 
                        ? new List<string>(updatedFilter.SelectedHostCategories) 
                        : null;
                    existingFilter.OpeningSettings = updatedFilter.OpeningSettings;
                    existingFilter.LastModified = updatedFilter.LastModified;
                    _log($"[FILTER_MGMT] ✅ Updated in-memory filter '{filterName}' with saved UI state");
                }
                else
                {
                    // Filter not in memory yet - add it
                    _filters.Add(updatedFilter);
                    _log($"[FILTER_MGMT] Added filter '{filterName}' to in-memory list (was not tracked)");
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error updating filter in memory: {ex.Message}");
            }
        }

        private void RefreshFilterList(ListBox filterListBox)
        {
            if (filterListBox != null)
            {
                filterListBox.Refresh();
                _log("[FILTER_MGMT] Refreshed filter list");
            }
        }

        private bool IsDerivedCategoryFilterName(ListBox listBox, string filterName)
        {
            if (string.IsNullOrWhiteSpace(filterName))
                return false;

            var normalized = FilterNameHelper.NormalizeBaseName(filterName);
            if (filterName.Equals(normalized, StringComparison.OrdinalIgnoreCase))
                return false;

            bool BaseExists(string candidate)
            {
                if (string.IsNullOrWhiteSpace(candidate))
                    return false;

                if (_filters.Any(f => string.Equals(f.Name, candidate, StringComparison.OrdinalIgnoreCase)))
                    return true;

                if (listBox != null)
                {
                    foreach (var item in listBox.Items)
                    {
                        if (item is string itemName &&
                            string.Equals(itemName, candidate, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }

                try
                {
                    var filterDir = GetDefaultFilterDirectory();
                    var baseFile = Path.Combine(filterDir, candidate + ".xml");
                    if (File.Exists(baseFile))
                        return true;
                }
                catch
                {
                    // Ignore path issues; absence of file just means base not found
                }

                return false;
            }

            return BaseExists(normalized);
        }

        /// <summary>
        /// Saves a filter to an XML file
        /// </summary>
        public void SaveFilterToXmlFile(OpeningFilter filter, string filePath)
        {
            // ✅ PHASE 2: If XML creation is disabled, persist UI state to DB instead
            if (DeploymentConfiguration.DisableXmlCreation)
            {
                Log($"[FILTER_MGMT] ⚠️ XML creation disabled - persisting UI state to DB instead. Filter='{filter?.Name}', Path='{filePath}'");

                try
                {
                    // Determine category for DB registration: prefer display category from filter, else fallback to enum name
                    var categoryDisplay = GetDisplayCategory(filter);
                    if (string.IsNullOrWhiteSpace(categoryDisplay) && filter != null)
                    {
                        try { categoryDisplay = MepCategoryConstants.Normalize(filter.SelectedMepCategoryName ?? filter.Category.ToString()); } catch { categoryDisplay = string.Empty; }
                    }

                    UseFilterRepository(repo =>
                    {
                        repo.SaveFilterUIState(
                            filter?.Name ?? string.Empty,
                            categoryDisplay ?? string.Empty,
                            filter?.SelectedHostCategories ?? new List<string>(),
                            filter?.OpeningSettings,
                            filter?.SelectedMepCategoryNames,    // ✅ FIXED: Pass MEP category names
                            filter?.SelectedReferenceFiles,      // ✅ FIXED: Pass reference files
                            filter?.SelectedHostFiles            // ✅ FIXED: Pass host files
                        );
                    });

                    Log($"[FILTER_MGMT] ✅ Persisted UI state to DB for filter '{filter?.Name}' (via SaveFilterToXmlFile fallback)");
                }
                catch (Exception ex)
                {
                    Log($"[FILTER_MGMT] ⚠️ Failed to persist UI state to DB in SaveFilterToXmlFile: {ex.Message}");
                }

                return;
            }
            
            try
            {
                Log($"[FILTER_MGMT] SaveFilterToXmlFile called: filePath='{filePath}', filter.Name='{filter?.Name}'");
                
                if (filter == null)
                {
                    Log("[FILTER_MGMT] ❌ ERROR: Filter is null!");
                    throw new ArgumentNullException(nameof(filter));
                }

                // ✅ DIAGNOSTIC: Check if directory exists
                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrWhiteSpace(directory))
                {
                    Log($"[FILTER_MGMT] ❌ ERROR: Directory path is null or empty for filePath: '{filePath}'");
                    throw new InvalidOperationException($"Directory path is null or empty for filePath: '{filePath}'");
                }
                
                if (!Directory.Exists(directory))
                {
                    Log($"[FILTER_MGMT] ❌ ERROR: Directory does not exist: '{directory}'");
                    throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'");
                }
                
                Log($"[FILTER_MGMT] ✅ Directory exists: '{directory}'");

                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                Log($"[FILTER_MGMT] ✅ Serializer created successfully");

                if (filter.ClashZoneStorage == null)
                {
                    filter.ClashZoneStorage = new ClashZoneStorage();
                }

                if (filter.ClashZoneStorage.Filters == null)
                {
                    filter.ClashZoneStorage.Filters = new List<FilterGroupForStorage>();
                }

                if (filter.ClashZoneStorage.ClashZones == null)
                {
                    filter.ClashZoneStorage.ClashZones = new List<ClashZone>();
                }

                var hasTreeZones = filter.ClashZoneStorage.Filters
                    .SelectMany(fg => fg?.FileCombos ?? Enumerable.Empty<FilterFileComboGroup>())
                    .Any(fc => fc?.ClashZones != null && fc.ClashZones.Count > 0);

                if (!hasTreeZones)
                {
                    if (filter.ClashZoneStorage.Filters.Count == 0)
                    {
                        filter.ClashZoneStorage.Filters.Add(new FilterGroupForStorage
                        {
                            Name = filter.Name ?? "Unknown",
                            FileCombos = new List<FilterFileComboGroup>()
                        });
                    }

                    var filterGroup = filter.ClashZoneStorage.Filters[0];
                    if (filterGroup.FileCombos == null)
                    {
                        filterGroup.FileCombos = new List<FilterFileComboGroup>();
                    }

                    if (filterGroup.FileCombos.All(fc => fc?.ClashZones == null || fc.ClashZones.Count == 0))
                    {
                        var defaultCombo = filterGroup.FileCombos.FirstOrDefault();
                        if (defaultCombo == null)
                        {
                            defaultCombo = new FilterFileComboGroup
                            {
                                LinkedFile = "Unknown",
                                HostFile = "Unknown",
                                ClashZones = new List<ClashZone>()
                            };
                            filterGroup.FileCombos.Add(defaultCombo);
                        }

                        if (defaultCombo.ClashZones == null)
                        {
                            defaultCombo.ClashZones = new List<ClashZone>();
                        }

                        if (filter.ClashZoneStorage.AllZones != null)
                        {
                            defaultCombo.ClashZones.AddRange(filter.ClashZoneStorage.AllZones.Where(z => z != null));
                        }
                    }
                }

                var allZonesList = new List<ClashZone>();
                if (filter.ClashZoneStorage.AllZones != null)
                {
                    allZonesList.AddRange(filter.ClashZoneStorage.AllZones.Where(z => z != null));
                }

                if (allZonesList.Count == 0)
                {
                    var treeZones = filter.ClashZoneStorage.Filters
                        .SelectMany(fg => fg.FileCombos ?? Enumerable.Empty<FilterFileComboGroup>())
                        .SelectMany(fc => fc.ClashZones ?? new List<ClashZone>())
                        .Where(z => z != null);

                    allZonesList.AddRange(treeZones);
                }

                allZonesList = allZonesList
                    .GroupBy(z => z.Id)
                    .Select(g => g.First())
                    .ToList();

                DeduplicateFileCombos(filter.ClashZoneStorage);

                if (filter.ClashZoneStorage.Filters != null)
                {
                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                    {
                        if (filterGroup?.FileCombos == null) continue;

                        foreach (var fileCombo in filterGroup.FileCombos)
                        {
                            if (fileCombo == null) continue;

                            if (fileCombo.ClashZones == null)
                            {
                                fileCombo.ClashZones = new List<ClashZone>();
                            }
                            else
                            {
                                fileCombo.ClashZones.RemoveAll(z => z == null);
                            }

                            foreach (var z in fileCombo.ClashZones)
                            {
                                NormalizeIntersectionPoint(z);
                            }
                        }
                    }
                }

                if (filter.ClashZoneStorage.AllZones != null)
                {
                    foreach (var z in filter.ClashZoneStorage.AllZones)
                    {
                        NormalizeIntersectionPoint(z);
                    }
                }

                Log($"[FILTER_MGMT] About to serialize filter to: {filePath}");
                
                using (var writer = new System.IO.StreamWriter(filePath))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        try
                        {
                            System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [FILTER-SAVE] Writing filter XML: {filePath}\n");
                        }
                        catch { }
                    }
                    
                    Log($"[FILTER_MGMT] StreamWriter created, about to serialize...");
                    serializer.Serialize(writer, filter);
                    Log($"[FILTER_MGMT] ✅ Serialization completed successfully");
                }
                
                // ✅ VERIFY: Check if file was actually created
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    Log($"[FILTER_MGMT] ✅ File verified: {filePath} ({fileInfo.Length} bytes)");
                }
                else
                {
                    Log($"[FILTER_MGMT] ❌ ERROR: File was NOT created at: {filePath}");
                    throw new FileNotFoundException($"File was not created after serialization: {filePath}");
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        var lastWrite = System.IO.File.GetLastWriteTime(filePath);
                        System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [FILTER-SAVE] Completed write: {filePath} (LastWrite={lastWrite:yyyy-MM-dd HH:mm:ss})\n");
                    }
                    catch { }
                }
                Log($"[FILTER_MGMT] ✅ Saved filter to XML: {filePath}");
            }
            catch (Exception ex)
            {
                var errorMsg = $"[FILTER_MGMT] ❌ ERROR serializing filter to XML: {ex.Message}";
                Log(errorMsg);
                Log($"[FILTER_MGMT] Exception type: {ex.GetType().Name}");
                Log($"[FILTER_MGMT] Stack trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Log($"[FILTER_MGMT] Inner exception: {ex.InnerException.Message}");
                }
                // ✅ CRITICAL: Also log to DebugLogger for visibility
                try
                {
                    DebugLogger.Error(errorMsg);
                    DebugLogger.Error($"Stack: {ex.StackTrace}");
                }
                catch { }
                throw;
            }
        }

        /// <summary>
        /// Helper method to normalize intersection point coordinates
        /// </summary>
        private void NormalizeIntersectionPoint(Models.ClashZone z)
        {
            if (z == null)
                return;

            // If the serialized intersection coordinates are already non-zero, respect them.
            bool hasSerializedIntersection =
                Math.Abs(z.IntersectionPointX) > 1e-9 ||
                Math.Abs(z.IntersectionPointY) > 1e-9 ||
                Math.Abs(z.IntersectionPointZ) > 1e-9;

            if (hasSerializedIntersection)
                return;

            bool hasIP = z.IntersectionPoint != null;
            bool isIPZero = hasIP &&
                            Math.Abs(z.IntersectionPoint.X) < 1e-9 &&
                            Math.Abs(z.IntersectionPoint.Y) < 1e-9 &&
                            Math.Abs(z.IntersectionPoint.Z) < 1e-9;

            if (hasIP && !isIPZero)
            {
                z.IntersectionPointX = z.IntersectionPoint.X;
                z.IntersectionPointY = z.IntersectionPoint.Y;
                z.IntersectionPointZ = z.IntersectionPoint.Z;
                return;
            }

            // If detection ever failed to populate IntersectionPoint, leave the serialized values at zero.
            // Downstream replay logic will skip placement when IntersectionPoint is (0,0,0), avoiding host-centre fallbacks.
        }

        public OpeningFilter LoadFilterFromXmlFile(string filePath)
        {
            try
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new System.IO.StreamReader(filePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);

                    _log($"[FILTER_MGMT] Loaded filter from XML: {filePath}");
                    
                    // ✅ PHASE 2: Load UI state from database (database is primary, XML is fallback)
                    try
                    {
                        var categoryDisplay = GetDisplayCategory(filter);
                        
                        // ✅ CRITICAL FIX: If GetDisplayCategory returns empty, try to get from SelectedMepCategoryNames list
                        if (string.IsNullOrEmpty(categoryDisplay) && filter.SelectedMepCategoryNames != null && filter.SelectedMepCategoryNames.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(filter.SelectedMepCategoryNames[0]);
                            _log($"[FILTER_MGMT] ℹ️ GetDisplayCategory was empty, using first MEP category from XML list: '{categoryDisplay}'");
                        }
                        else if (string.IsNullOrEmpty(categoryDisplay))
                        {
                            _log($"[FILTER_MGMT] ⚠️ GetDisplayCategory returned empty and no MEP categories in filter - cannot load UI state from database");
                        }
                        
                        UseFilterRepository(repo =>
                        {
                            var (hostCategories, settings, mepCategories, refFiles, hostFiles) = repo.LoadFilterUIState(filter.Name, categoryDisplay);
                            if (hostCategories != null && hostCategories.Count > 0)
                            {
                                filter.SelectedHostCategories = hostCategories;
                                _log($"[FILTER_MGMT] ✅ Loaded SelectedHostCategories from database for filter '{filter.Name}': {string.Join(", ", hostCategories)}");
                            }
                            else if (filter.SelectedHostCategories == null || filter.SelectedHostCategories.Count == 0)
                            {
                                // ✅ FALLBACK: Use XML data if database has no data
                                _log($"[FILTER_MGMT] ⚠️ No SelectedHostCategories in database, using XML data (if available)");
                            }
                            if (mepCategories != null && mepCategories.Count > 0)
                            {
                                filter.SelectedMepCategoryNames = mepCategories;
                                // ✅ CRITICAL FIX: Also set singular field so GetDisplayCategory() returns correct value
                                filter.SelectedMepCategoryName = mepCategories[0];
                                _log($"[FILTER_MGMT] ✅ Loaded SelectedMepCategoryNames from database for filter '{filter.Name}': {string.Join(", ", mepCategories)}, set SelectedMepCategoryName to '{mepCategories[0]}'");
                            }
                            if (refFiles != null && refFiles.Count > 0)
                            {
                                filter.SelectedReferenceFiles = refFiles;
                                _log($"[FILTER_MGMT] ✅ Loaded SelectedReferenceFiles from database for filter '{filter.Name}': {string.Join(", ", refFiles)}");
                            }
                            if (hostFiles != null && hostFiles.Count > 0)
                            {
                                filter.SelectedHostFiles = hostFiles;
                                _log($"[FILTER_MGMT] ✅ Loaded SelectedHostFiles from database for filter '{filter.Name}': {string.Join(", ", hostFiles)}");
                            }
                            
                            if (settings != null)
                            {
                                filter.OpeningSettings = settings;
                                _log($"[FILTER_MGMT] ✅ Loaded OpeningSettings from database for filter '{filter.Name}'");
                            }
                            else if (filter.OpeningSettings == null)
                            {
                                // ✅ FALLBACK: Use XML data if database has no data
                                _log($"[FILTER_MGMT] ⚠️ No OpeningSettings in database, using XML data (if available)");
                            }
                        });
                    }
                    catch (Exception dbEx)
                    {
                        _log($"[FILTER_MGMT] ⚠️ Could not load UI state from database for filter '{filter.Name}', using XML data: {dbEx.Message}");
                    }
                    
                    // ✅ CRITICAL: Reconstruct SleevePlacementPoint from XML-serializable properties
                    if (filter?.ClashZoneStorage?.AllZones != null)
                    {
                        foreach (var clashZone in filter.ClashZoneStorage.AllZones)
                        {
                            clashZone.EnsureSleevePlacementPointReconstructed();
                        }
                        _log($"[FILTER_MGMT] Reconstructed SleevePlacementPoint for {filter.ClashZoneStorage.AllZones.Count} clash zones");
                        
                        // ✅ GLOBAL XML FLAG MANAGEMENT: Sync flags from Global XML immediately after loading Filter XML
                        // Flags are NOT persisted to Filter XML - must sync from Global XML (single source of truth)
                        try
                        {
                            // Only sync flags if document is available
                            if (_document != null)
                            {
                                var flagManager = new FlagManager(_document);
                                var clashZonesByCategory = filter.ClashZoneStorage.AllZones
                                    .GroupBy(cz => cz.MepElementCategory)
                                    .ToList();
                                
                                foreach (var categoryGroup in clashZonesByCategory)
                                {
                                    var category = categoryGroup.Key;
                                    var categoryClashZones = categoryGroup.ToList();
                                    
                                    if (!string.IsNullOrWhiteSpace(category) && categoryClashZones.Count > 0)
                                    {
                                        flagManager.SyncFlagsFromGlobal(categoryClashZones, category);
                                        _log($"[FILTER_MGMT] Synced flags from Global XML for {categoryClashZones.Count} clash zones in category '{category}'");
                                    }
                                }
                            }
                            else
                            {
                                _log($"[FILTER_MGMT] Warning: Document not available - flags will be synced from Global XML when document is available");
                            }
                        }
                        catch (Exception flagEx)
                        {
                            _log($"[FILTER_MGMT] Warning: Failed to sync flags from Global XML: {flagEx.Message}");
                            // Continue - flags will default to false, will be synced later if needed
                        }
                    }
                    
                    // Parameter transfer functionality removed
                    _log($"[FILTER_MGMT] Parameter transfer functionality has been removed from filter '{filter.Name}'");
                    
                    return filter;
                }
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error deserializing filter from XML: {ex.Message}");
                throw;
            }
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void MergeClashZoneData(ClashZone target, ClashZone source)
        {
            if (target == null || source == null) return;

            if (HasValue(source.SourceDocKey)) target.SourceDocKey = source.SourceDocKey;
            if (HasValue(source.HostDocKey)) target.HostDocKey = source.HostDocKey;
            if (HasValue(source.StructuralElementDocumentTitle)) target.StructuralElementDocumentTitle = source.StructuralElementDocumentTitle;

            if (source.MepParameterValues != null && source.MepParameterValues.Count > 0)
                target.MepParameterValues = source.MepParameterValues;
            if (source.HostParameterValues != null && source.HostParameterValues.Count > 0)
                target.HostParameterValues = source.HostParameterValues;

            if (source.StructuralElementThickness > 0) target.StructuralElementThickness = source.StructuralElementThickness;
            if (source.StructuralElementNormal != null) target.StructuralElementNormal = source.StructuralElementNormal;

            if (Math.Abs(source.IntersectionPointX) > 1e-9 || Math.Abs(source.IntersectionPointY) > 1e-9 || Math.Abs(source.IntersectionPointZ) > 1e-9)
            {
                target.IntersectionPointX = source.IntersectionPointX;
                target.IntersectionPointY = source.IntersectionPointY;
                target.IntersectionPointZ = source.IntersectionPointZ;
            }

            if (HasPlacementPoint(source))
            {
                target.SleevePlacementPointX = source.SleevePlacementPointX;
                target.SleevePlacementPointY = source.SleevePlacementPointY;
                target.SleevePlacementPointZ = source.SleevePlacementPointZ;
            }

            if (HasActivePlacementPoint(source))
            {
                target.SleevePlacementPointActiveDocumentX = source.SleevePlacementPointActiveDocumentX;
                target.SleevePlacementPointActiveDocumentY = source.SleevePlacementPointActiveDocumentY;
                target.SleevePlacementPointActiveDocumentZ = source.SleevePlacementPointActiveDocumentZ;
            }

            if (source.SleeveWidth > 0) target.SleeveWidth = source.SleeveWidth;
            if (source.SleeveHeight > 0) target.SleeveHeight = source.SleeveHeight;
            if (source.SleeveDiameter > 0) target.SleeveDiameter = source.SleeveDiameter;

            if (HasBoundingBox(source))
            {
                target.SleeveBoundingBoxMinX = source.SleeveBoundingBoxMinX;
                target.SleeveBoundingBoxMinY = source.SleeveBoundingBoxMinY;
                target.SleeveBoundingBoxMinZ = source.SleeveBoundingBoxMinZ;
                target.SleeveBoundingBoxMaxX = source.SleeveBoundingBoxMaxX;
                target.SleeveBoundingBoxMaxY = source.SleeveBoundingBoxMaxY;
                target.SleeveBoundingBoxMaxZ = source.SleeveBoundingBoxMaxZ;
            }

            if (source.SleeveInstanceId > 0) target.SleeveInstanceId = source.SleeveInstanceId;
            if (source.ClusterSleeveInstanceId > 0) target.ClusterSleeveInstanceId = source.ClusterSleeveInstanceId;
            if (source.AfterClusterSleevePlacedSleeveInstanceId > 0) target.AfterClusterSleevePlacedSleeveInstanceId = source.AfterClusterSleevePlacedSleeveInstanceId;

            target.IsResolved = source.IsResolved || target.IsResolved;
            target.IsClusterResolved = source.IsClusterResolved || target.IsClusterResolved;

            if (HasValue(source.SleeveFamilyName)) target.SleeveFamilyName = source.SleeveFamilyName;

            if (HasValue(source.MepElementCategory)) target.MepElementCategory = source.MepElementCategory;
            if (source.MepElementWidth > 0) target.MepElementWidth = source.MepElementWidth;
            if (source.MepElementHeight > 0) target.MepElementHeight = source.MepElementHeight;
            if (HasValue(source.MepElementFormattedSize)) target.MepElementFormattedSize = source.MepElementFormattedSize;
            if (HasValue(source.MepElementSystemAbbreviation)) target.MepElementSystemAbbreviation = source.MepElementSystemAbbreviation;

            if (HasValue(source.DocumentPath)) target.DocumentPath = source.DocumentPath;
            if (source.DetectedAt != default) target.DetectedAt = source.DetectedAt;
            if (source.LastUpdated != default) target.LastUpdated = source.LastUpdated;
        }

        private static bool HasValue(string value) => !string.IsNullOrWhiteSpace(value);

        private static bool HasPlacementPoint(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleevePlacementPointX) > 1e-9 || Math.Abs(zone.SleevePlacementPointY) > 1e-9 || Math.Abs(zone.SleevePlacementPointZ) > 1e-9;
        }

        private static bool HasActivePlacementPoint(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleevePlacementPointActiveDocumentX) > 1e-9 || Math.Abs(zone.SleevePlacementPointActiveDocumentY) > 1e-9 || Math.Abs(zone.SleevePlacementPointActiveDocumentZ) > 1e-9;
        }

        private static bool HasBoundingBox(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleeveBoundingBoxMinX) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMinY) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMinZ) > 1e-9 ||
                   Math.Abs(zone.SleeveBoundingBoxMaxX) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMaxY) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMaxZ) > 1e-9;
        }

        private static void DeduplicateFileCombos(ClashZoneStorage storage)
        {
            if (storage?.Filters == null) return;

            foreach (var filterGroup in storage.Filters)
            {
                if (filterGroup?.FileCombos == null || filterGroup.FileCombos.Count <= 1)
                    continue;

                var deduped = new List<FilterFileComboGroup>();

                foreach (var grouping in filterGroup.FileCombos
                             .Where(fc => fc != null)
                             .GroupBy(fc => fc.GetNormalizedKey()))
                {
                    var ordered = grouping
                        .OrderByDescending(fc => fc.ProcessedAt)
                        .ToList();

                    if (ordered.Count == 0)
                        continue;

                    var keeper = ordered.First();

                    foreach (var duplicate in ordered.Skip(1))
                    {
                        if (duplicate?.ClashZones == null || duplicate.ClashZones.Count == 0)
                            continue;

                        if (keeper.ClashZones == null)
                            keeper.ClashZones = new List<ClashZone>();

                        foreach (var zone in duplicate.ClashZones)
                        {
                            if (zone == null) continue;
                            var existing = keeper.ClashZones.FirstOrDefault(z => z != null && z.Id == zone.Id);
                            if (existing == null)
                            {
                                keeper.ClashZones.Add(zone);
                            }
                            else
                            {
                                MergeClashZoneData(existing, zone);
                            }
                        }
                    }

                    if (keeper.ClashZones != null)
                    {
                        keeper.ClashZones = keeper.ClashZones
                            .Where(z => z != null)
                            .GroupBy(z => z.Id)
                            .Select(g => g.First())
                            .ToList();
                    }

                    deduped.Add(keeper);
                }

                filterGroup.FileCombos = deduped;
            }
        }

        #endregion
    }

    /// <summary>
    /// Simple input dialog for getting user input
    /// </summary>
    public class InputDialog : System.Windows.Forms.Form
    {
        private TextBox _inputTextBox;
        private Button _okButton;
        private Button _cancelButton;

        public string InputText => _inputTextBox?.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent(title, prompt, defaultValue);
        }

        private void InitializeComponent(string title, string prompt, string defaultValue)
        {
            Text = title;
            Size = new System.Drawing.Size(400, 150);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true; // Keep dialog on top
            ShowInTaskbar = false; // Don't show in taskbar

            var promptLabel = new Label
            {
                Text = prompt,
                Location = new System.Drawing.Point(12, 12),
                Size = new System.Drawing.Size(360, 20),
                AutoSize = true
            };
            Controls.Add(promptLabel);

            _inputTextBox = new TextBox
            {
                Text = defaultValue,
                Location = new System.Drawing.Point(12, 40),
                Size = new System.Drawing.Size(360, 20)
            };
            Controls.Add(_inputTextBox);

            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(216, 70),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.OK
            };
            Controls.Add(_okButton);

            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(297, 70),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.Cancel
            };
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
            
            // Ensure the dialog is properly modal and visible
            _inputTextBox.Focus();
            _inputTextBox.SelectAll();
        }
    }
}

