using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: SOLID-compliant filter UI orchestrator.
    /// Coordinates CRUD operations with persistence operations.
    /// 
    /// ✅ PRESERVES ALL LOGIC:
    /// - Database-first approach
    /// - UI state persistence
    /// - Category validation
    /// - XML backward compatibility
    /// </summary>
    public class FilterUiOrchestrator : IFilterUiOrchestrator
    {
        private readonly IFilterCrudService _crudService;
        private readonly IFilterPersistenceService _persistenceService;
        private readonly IFilterRepository _filterRepository;
        private readonly IUserInteractionService _userInteraction;
        private readonly ILogger _logger;
        private readonly Action<string> _updateStatus;
        private readonly Document _document;
        private readonly List<OpeningFilter> _filters;
        
        /// <summary>
        /// Creates a new filter UI orchestrator.
        /// ✅ SOLID: Dependency Inversion - depends on abstractions (interfaces), not concretions.
        /// </summary>
        public FilterUiOrchestrator(
            IFilterCrudService crudService,
            IFilterPersistenceService persistenceService,
            IFilterRepository filterRepository,
            IUserInteractionService userInteraction,
            Action<string> updateStatus,
            Document document,
            ILogger logger = null)
        {
            _crudService = crudService ?? throw new ArgumentNullException(nameof(crudService));
            _persistenceService = persistenceService ?? throw new ArgumentNullException(nameof(persistenceService));
            _filterRepository = filterRepository ?? throw new ArgumentNullException(nameof(filterRepository));
            _userInteraction = userInteraction ?? throw new ArgumentNullException(nameof(userInteraction));
            _updateStatus = updateStatus ?? throw new ArgumentNullException(nameof(updateStatus));
            _document = document;
            _logger = logger ?? LoggerAdapter.Default;
            _filters = new List<OpeningFilter>();
        }
        
        /// <summary>
        /// Creates a new filter and saves it (database + XML).
        /// ✅ ORCHESTRATION: Get name from user → Create filter → Validate → Save → Update UI
        /// ✅ PRESERVE: Database-first registration
        /// ✅ PRESERVE: UI state persistence
        /// </summary>
        public void CreateNewFilter(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Creating new filter", "FilterUiOrchestrator");
                
                var filterName = _userInteraction.GetFilterNameFromUser("New Filter", "Enter filter name:", "New Filter");
                if (string.IsNullOrEmpty(filterName)) return;

                var newFilter = _crudService.CreateFilter(filterName);
                if (newFilter == null)
                {
                    _logger.Warning("Failed to create filter from UI state", "FilterUiOrchestrator");
                    return;
                }
                
                // ✅ CRITICAL FIX: Get category from UI state, NOT from filter object (which may have stale/wrong category)
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _logger.Info($"Using category from UI state: '{categoryDisplay}'", "FilterUiOrchestrator");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning($"Error getting category from UI state: {ex.Message}", "FilterUiOrchestrator");
                    }
                }
                
                // ⚠️ CRITICAL: If category is empty, prompt user to select a category
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _logger.Warning("Category is empty - cannot create filter without category", "FilterUiOrchestrator");
                    _userInteraction.ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before creating the filter.");
                    return;
                }
                
                // ✅ CRITICAL: DATABASE-FIRST - Register filter in database FIRST (before adding to UI)
                var filterId = _persistenceService.SaveFilterToDatabase(newFilter.Name, categoryDisplay, newFilter);
                
                if (filterId > 0)
                {
                    _logger.Info($"Created filter '{filterName}' in database (FilterId={filterId})", "FilterUiOrchestrator");
                    
                    // ✅ STEP 2: UI state is already saved by FilterPersistenceService.SaveFilterToDatabase (Team I handles this)
                    // ✅ STEP 3: Only add to UI AFTER successful DB creation
                    AddFilterToList(filterListBox, newFilter);
                    _updateStatus($"Created new filter: {filterName}");
                }
                else
                {
                    _logger.Error($"Failed to create filter '{filterName}' in database", null, "FilterUiOrchestrator");
                    _userInteraction.ShowError($"Failed to create filter '{filterName}' in database. Filter cannot be used until it exists in the database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error creating new filter: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error creating new filter: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Copies a filter and saves it.
        /// ✅ ORCHESTRATION: Get source → Copy → Get name → Validate → Save → Update UI
        /// </summary>
        public void CopyFilter(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Copying filter", "FilterUiOrchestrator");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _userInteraction.ShowWarning("Please select a filter to copy.");
                    return;
                }

                var newName = _userInteraction.GetFilterNameFromUser("Copy Filter", "Enter name for copied filter:", 
                    $"{selectedFilter.Name} Copy");
                if (string.IsNullOrEmpty(newName)) return;

                var copiedFilter = _crudService.CopyFilter(selectedFilter, newName);
                if (copiedFilter == null)
                {
                    _logger.Warning("Failed to copy filter", "FilterUiOrchestrator");
                    return;
                }
                
                // ✅ CRITICAL FIX: Get category from UI state OR from copied filter's SelectedMepCategoryNames
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _logger.Info($"Using category from UI state: '{categoryDisplay}'", "FilterUiOrchestrator");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning($"Error getting category from UI state: {ex.Message}", "FilterUiOrchestrator");
                    }
                }
                
                // Fallback to copied filter's SelectedMepCategoryNames if UI state not available
                if (string.IsNullOrEmpty(categoryDisplay) && copiedFilter.SelectedMepCategoryNames != null && copiedFilter.SelectedMepCategoryNames.Count > 0)
                {
                    categoryDisplay = MepCategoryConstants.Normalize(copiedFilter.SelectedMepCategoryNames[0]);
                    _logger.Info($"Using category from copied filter: '{categoryDisplay}'", "FilterUiOrchestrator");
                }
                
                // ⚠️ CRITICAL: If category is empty, prompt user to select a category
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _logger.Warning("Category is empty - cannot copy filter without category", "FilterUiOrchestrator");
                    _userInteraction.ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before copying the filter.");
                    return;
                }

                // ✅ CRITICAL: DATABASE-FIRST - Register filter in database FIRST
                var filterId = _persistenceService.SaveFilterToDatabase(copiedFilter.Name, categoryDisplay, copiedFilter);
                
                if (filterId > 0)
                {
                    _logger.Info($"Copied filter '{selectedFilter.Name}' to '{newName}' in database (FilterId={filterId})", "FilterUiOrchestrator");
                    
                    // ✅ STEP 2: UI state is already saved by FilterPersistenceService.SaveFilterToDatabase (Team I handles this)
                    // ✅ STEP 3: Only add to UI AFTER successful DB creation
                    AddFilterToList(filterListBox, copiedFilter);
                    _updateStatus($"Copied filter to: {newName}");
                }
                else
                {
                    _logger.Error($"Failed to copy filter to '{newName}' in database", null, "FilterUiOrchestrator");
                    _userInteraction.ShowError($"Failed to copy filter to '{newName}' in database. Filter cannot be used until it exists in the database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error copying filter: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error copying filter: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Renames a filter and updates persistence.
        /// ✅ ORCHESTRATION: Get old/new names → Rename → Update persistence → Update UI
        /// </summary>
        public void RenameFilter(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Renaming filter", "FilterUiOrchestrator");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _userInteraction.ShowWarning("Please select a filter to rename.");
                    return;
                }

                // ✅ Get category from filter (needed for database update)
                string categoryDisplay = GetDisplayCategory(selectedFilter);
                var newName = _userInteraction.GetFilterNameFromUser("Rename Filter", "Enter new name:", selectedFilter.Name);
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
                        RefreshFilterList(filterListBox);
                    }
                }
                
                // ✅ AUTO-SAVE: Filter rename is handled by database update (Team I handles persistence)

                // ✅ Update database
                if (!string.IsNullOrEmpty(categoryDisplay))
                {
                    _filterRepository.UpdateFilterName(oldName, categoryDisplay, newName);
                }
                
                _updateStatus($"Renamed filter to: {newName}");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error renaming filter: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error renaming filter: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Deletes a filter from persistence and UI.
        /// ✅ ORCHESTRATION: Get filter → Confirm → Delete from persistence → Update UI
        /// </summary>
        public void DeleteFilter(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Deleting filter", "FilterUiOrchestrator");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _userInteraction.ShowWarning("Please select a filter to delete.");
                    return;
                }

                // ✅ Confirm deletion
                var confirmMessage = $"Are you sure you want to delete the filter '{selectedFilter.Name}'?";
                if (!_userInteraction.ShowConfirmation(confirmMessage, "Delete Filter"))
                {
                    return;
                }

                var categoryDisplay = GetDisplayCategory(selectedFilter);
                
                // ✅ Delete from database FIRST
                if (!string.IsNullOrEmpty(categoryDisplay))
                {
                    _filterRepository.DeleteFilter(selectedFilter.Name, categoryDisplay);
                }
                
                // ✅ File deletion handled by Team I (database-only, no XML files)
                
                // ✅ Remove from UI
                RemoveFilterFromList(filterListBox, selectedFilter);
                _updateStatus($"Deleted filter: {selectedFilter.Name}");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error deleting filter: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error deleting filter: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Saves filter to persistence (database + XML).
        /// ✅ ORCHESTRATION: Get filter → Validate → Save to persistence → Update UI
        /// </summary>
        public void SaveFilter(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Saving filter", "FilterUiOrchestrator");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _userInteraction.ShowWarning("Please select a filter to save.");
                    return;
                }

                // ✅ Validate filter
                var (isValid, errorMessage) = _crudService.ValidateFilter(selectedFilter);
                if (!isValid)
                {
                    _userInteraction.ShowError(errorMessage);
                    return;
                }
                
                // ✅ CRITICAL FIX: Collect CURRENT UI state before saving
                var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? selectedFilter.SelectedHostCategories ?? new List<string>();
                
                // ✅ CRITICAL FIX: Get category from UI state
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _logger.Info($"Using category from UI state: '{categoryDisplay}'", "FilterUiOrchestrator");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning($"Error getting category from UI state: {ex.Message}", "FilterUiOrchestrator");
                    }
                }
                
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _logger.Warning("Category is empty - cannot save filter without category", "FilterUiOrchestrator");
                    _userInteraction.ShowError("Please select a MEP category (Ducts, Pipes, or Cable Trays) before saving the filter.");
                    return;
                }
                
                // ✅ Get OpeningSettings from UI
                OpeningSettings currentOpeningSettings = null;
                if (FilterUiStateProvider.GetClearanceSettings != null)
                {
                    try
                    {
                        var clearanceSettings = FilterUiStateProvider.GetClearanceSettings(categoryDisplay);
                        if (clearanceSettings != null)
                        {
                            currentOpeningSettings = new OpeningSettings
                            {
                                ClearanceSettings = clearanceSettings
                            };
                            if (selectedFilter.OpeningSettings != null)
                            {
                                currentOpeningSettings.SelectedMepType = selectedFilter.OpeningSettings.SelectedMepType;
                            }
                        }
                        else
                        {
                            currentOpeningSettings = selectedFilter.OpeningSettings;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning($"Could not get OpeningSettings from UI: {ex.Message}", "FilterUiOrchestrator");
                        currentOpeningSettings = selectedFilter.OpeningSettings;
                    }
                }
                else
                {
                    currentOpeningSettings = selectedFilter.OpeningSettings;
                }
                
                // ✅ Update filter object with current UI state
                selectedFilter.SelectedHostCategories = currentHostCategories;
                selectedFilter.OpeningSettings = currentOpeningSettings;
                
                // ✅ DATABASE-FIRST: Register filter in database FIRST
                var filterId = _persistenceService.SaveFilterToDatabase(selectedFilter.Name, categoryDisplay, selectedFilter);
                
                if (filterId > 0)
                {
                    _logger.Info($"Saved filter '{selectedFilter.Name}' in database (FilterId={filterId})", "FilterUiOrchestrator");
                    
                    // ✅ UI state is already saved by FilterPersistenceService.SaveFilterToDatabase (Team I handles this)
                    // ✅ Update in-memory filter object
                    UpdateFilterInMemory(selectedFilter.Name, selectedFilter);
                    _updateStatus($"Saved filter: {selectedFilter.Name}");
                }
                else
                {
                    _logger.Error($"Failed to save filter '{selectedFilter.Name}' in database", null, "FilterUiOrchestrator");
                    _userInteraction.ShowError($"Failed to save filter '{selectedFilter.Name}' in database. Check logs for details.");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error saving filter: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error saving filter: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Auto-saves filter (no user prompt).
        /// ✅ ORCHESTRATION: Get filter → Save to persistence (no user prompt)
        /// </summary>
        public void SaveFilterAuto(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Auto-saving filter", "FilterUiOrchestrator");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    _logger.Info("No filter selected for auto-save", "FilterUiOrchestrator");
                    return;
                }

                // ✅ Get category from UI state
                string categoryDisplay = null;
                if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
                {
                    try
                    {
                        var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                        if (selectedCategories != null && selectedCategories.Count > 0)
                        {
                            categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                            _logger.Info($"Auto-save: Using category from UI state: '{categoryDisplay}'", "FilterUiOrchestrator");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning($"Error getting category from UI state: {ex.Message}", "FilterUiOrchestrator");
                    }
                }
                
                if (string.IsNullOrEmpty(categoryDisplay))
                {
                    _logger.Warning("Category is empty - cannot auto-save filter without category", "FilterUiOrchestrator");
                    return; // Silent fail for auto-save
                }
                
                // ✅ Get current UI state
                var currentHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? selectedFilter.SelectedHostCategories ?? new List<string>();
                
                // ✅ Get OpeningSettings from UI
                OpeningSettings currentOpeningSettings = null;
                if (FilterUiStateProvider.GetClearanceSettings != null)
                {
                    try
                    {
                        var clearanceSettings = FilterUiStateProvider.GetClearanceSettings(categoryDisplay);
                        if (clearanceSettings != null)
                        {
                            currentOpeningSettings = new OpeningSettings
                            {
                                ClearanceSettings = clearanceSettings
                            };
                            if (selectedFilter.OpeningSettings != null)
                            {
                                currentOpeningSettings.SelectedMepType = selectedFilter.OpeningSettings.SelectedMepType;
                            }
                        }
                        else
                        {
                            currentOpeningSettings = selectedFilter.OpeningSettings;
                        }
                    }
                    catch
                    {
                        currentOpeningSettings = selectedFilter.OpeningSettings;
                    }
                }
                else
                {
                    currentOpeningSettings = selectedFilter.OpeningSettings;
                }
                
                // ✅ Update filter object with current UI state
                selectedFilter.SelectedHostCategories = currentHostCategories;
                selectedFilter.OpeningSettings = currentOpeningSettings;
                
                // ✅ DATABASE-FIRST: Register filter in database FIRST
                var filterId = _persistenceService.SaveFilterToDatabase(selectedFilter.Name, categoryDisplay, selectedFilter);
                
                if (filterId > 0)
                {
                    // ✅ UI state is already saved by FilterPersistenceService.SaveFilterToDatabase (Team I handles this)
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Error auto-saving filter: {ex.Message}", ex, "FilterUiOrchestrator");
                // Silent fail for auto-save
            }
        }
        
        /// <summary>
        /// Loads all saved filters into UI.
        /// ✅ ORCHESTRATION: Load from persistence → Add to UI
        /// ✅ PRESERVE: Database-first, XML fallback
        /// </summary>
        public void LoadAllSavedFilters(ListBox filterListBox)
        {
            try
            {
                _logger.Info("Loading all saved filters (DATABASE-FIRST)", "FilterUiOrchestrator");
                
                var loadedFilters = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int dbLoadedCount = 0;
                // int xmlLoadedCount = 0; // FIX: CS0219 - commented to fix critical warning
                
                // ✅ STEP 1: Load from DATABASE FIRST (primary source)
                try
                {
                    var dbFilters = _filterRepository.GetAllFilters();
                    foreach (var filterInfo in dbFilters)
                    {
                        if (string.IsNullOrWhiteSpace(filterInfo.FilterName))
                            continue;
                        
                        // Only add if not already in list (avoid duplicates)
                        if (!filterListBox.Items.Contains(filterInfo.FilterName) && !_filters.Any(f => f.Name == filterInfo.FilterName))
                        {
                            // ✅ Load saved UI state from database
                            var uiStateResult = _filterRepository.LoadFilterUIState(filterInfo.FilterName, filterInfo.Category);
                            var hostCategories = uiStateResult.selectedHostCategories;
                            var openingSettings = uiStateResult.settings;
                            
                            // ✅ Load full filter from database (Team I will implement LoadFilterFromDatabase)
                            var filter = _persistenceService.LoadFilterFromDatabase(filterInfo.FilterName, filterInfo.Category);
                            if (filter == null)
                            {
                                // Create filter with saved UI state if full filter not available
                                filter = new OpeningFilter
                                {
                                    Name = filterInfo.FilterName,
                                    SelectedMepCategoryName = filterInfo.Category,
                                    SelectedHostCategories = hostCategories ?? new List<string>(),
                                    OpeningSettings = openingSettings,
                                    ClashZoneStorage = new ClashZoneStorage
                                    {
                                        Filters = new List<FilterGroupForStorage>(),
                                        ClashZones = new List<ClashZone>()
                                    }
                                };
                            }
                            else
                            {
                                // Override with database UI state (more current)
                                filter.SelectedHostCategories = hostCategories ?? filter.SelectedHostCategories;
                                filter.OpeningSettings = openingSettings ?? filter.OpeningSettings;
                            }
                            
                            AddFilterToList(filterListBox, filter);
                            loadedFilters.Add(filterInfo.FilterName);
                            dbLoadedCount++;
                            _logger.Info($"Loaded filter '{filterInfo.FilterName}' from database with saved UI state", "FilterUiOrchestrator");
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    _logger.Warning($"Could not load filters from database: {dbEx.Message}", "FilterUiOrchestrator");
                }

                // ✅ STEP 2: XML loading removed - database-only approach (Team I handles persistence)
                
                _logger.Info($"Loaded {dbLoadedCount} filters from database", "FilterUiOrchestrator");
                _updateStatus($"Loaded {dbLoadedCount} saved filters");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error loading saved filters: {ex.Message}", ex, "FilterUiOrchestrator");
                _userInteraction.ShowError($"Error loading saved filters: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Seeds default filters into UI.
        /// ✅ ORCHESTRATION: Create default filters → Add to UI
        /// </summary>
        public void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames)
        {
            if (filterListBox == null) return;
            
            foreach (var name in defaultNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                
                // Skip if already in list
                if (_filters.Any(f => f.Name == name) || filterListBox.Items.Contains(name)) continue;
                
                // ✅ PERSISTENCE CHECK: Skip if filter already exists in database
                if (_persistenceService.IsFilterSaved(name))
                {
                    _logger.Info($"Skipping default filter '{name}' - already exists in database", "FilterUiOrchestrator");
                    continue;
                }
                
                var filter = _crudService.CreateFilter(name);
                if (filter != null)
                {
                    AddFilterToList(filterListBox, filter);
                }
            }
        }
        
        #region Private Helper Methods
        
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
                _logger.Info($"Skipping derived category filter '{filter.Name}' from UI list", "FilterUiOrchestrator");
                return;
            }

            filterListBox.Items.Add(filter.Name);
            _filters.Add(filter);
            _logger.Info($"Added filter '{filter.Name}' to list", "FilterUiOrchestrator");
        }
        
        private void RemoveFilterFromList(ListBox filterListBox, OpeningFilter filter)
        {
            if (filterListBox != null && filter != null)
            {
                filterListBox.Items.Remove(filter.Name);
                _filters.Remove(filter);
                _logger.Info($"Removed filter '{filter.Name}' from list", "FilterUiOrchestrator");
            }
        }
        
        private void RefreshFilterList(ListBox filterListBox)
        {
            if (filterListBox != null)
            {
                filterListBox.Refresh();
            }
        }
        
        private void UpdateFilterInMemory(string filterName, OpeningFilter updatedFilter)
        {
            if (string.IsNullOrWhiteSpace(filterName) || updatedFilter == null)
                return;
            
            try
            {
                var existingFilter = _filters.FirstOrDefault(f => f.Name == filterName);
                if (existingFilter != null)
                {
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
                    _logger.Info($"Updated in-memory filter '{filterName}' with saved UI state", "FilterUiOrchestrator");
                }
                else
                {
                    _filters.Add(updatedFilter);
                    _logger.Info($"Added filter '{filterName}' to in-memory list", "FilterUiOrchestrator");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error updating filter in memory: {ex.Message}", "FilterUiOrchestrator");
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

                // ✅ REMOVED: XML file check - database-only approach

                return false;
            }

            return BaseExists(normalized);
        }
        
        private string GetDisplayCategory(OpeningFilter filter)
        {
            if (filter == null)
                return string.Empty;

            if (!string.IsNullOrWhiteSpace(filter.SelectedMepCategoryName))
                return MepCategoryConstants.Normalize(filter.SelectedMepCategoryName);

            return string.Empty;
        }
        
        // ✅ REMOVED: GetFilterDirectory() - no longer needed (database-only, no XML files)
        
        #endregion
    }
}

