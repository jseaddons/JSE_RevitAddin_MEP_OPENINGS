using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

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
                AddFilterToList(filterListBox, newFilter);
                
                // ✅ AUTO-SAVE: Save filter automatically to persistent storage
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
                }
                var filePath = Path.Combine(filterDir, $"{filterName}.xml");
                SaveFilterToXmlFile(newFilter, filePath);
                
                _log($"[FILTER_MGMT] Created and auto-saved new filter: {filterName}");
                _updateStatus($"Created new filter: {filterName}");
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
                    SelectedHostElementTypes = selectedFilter.SelectedHostElementTypes != null ? new List<string>(selectedFilter.SelectedHostElementTypes) : null
                };

                AddFilterToList(filterListBox, copiedFilter);
                
                // ✅ AUTO-SAVE: Save copied filter automatically to persistent storage
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
                }
                var filePath = Path.Combine(filterDir, $"{newName}.xml");
                SaveFilterToXmlFile(copiedFilter, filePath);
                
                _log($"[FILTER_MGMT] Copied and auto-saved filter '{selectedFilter.Name}' to '{newName}'");
                _updateStatus($"Copied filter to: {newName}");
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
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
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

                // ✅ SILENT SAVE: Automatically save to default filter directory without showing dialog
                var filterDir = GetDefaultFilterDirectory();
                var filePath = System.IO.Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                
                SaveFilterToXmlFile(selectedFilter, filePath);
                
                _log($"[FILTER_MGMT] Saved filter '{selectedFilter.Name}' to: {filePath}");
                _updateStatus($"Saved filter: {selectedFilter.Name}");
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

                // Get default filter directory
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
                }

                var filePath = Path.Combine(filterDir, $"{selectedFilter.Name}.xml");
                SaveFilterToXmlFile(selectedFilter, filePath);
                
                _log($"[FILTER_MGMT] Auto-saved filter '{selectedFilter.Name}' to {filePath}");
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error auto-saving filter: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads all saved filters from the default filter directory and adds them to the list
        /// </summary>
        public void LoadAllSavedFilters(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Loading all saved filters from directory");
                
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    _log($"[FILTER_MGMT] Filter directory does not exist: {filterDir}");
                    return;
                }

                // Get all XML files in the filter directory
                var xmlFiles = Directory.GetFiles(filterDir, "*.xml");
                int loadedCount = 0;
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var filter = LoadFilterFromXmlFile(xmlFile);
                        if (filter != null)
                        {
                            // Only add if not already in list (avoid duplicates)
                            if (!filterListBox.Items.Contains(filter.Name) && !_filters.Any(f => f.Name == filter.Name))
                            {
                                AddFilterToList(filterListBox, filter);
                                loadedCount++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"[FILTER_MGMT] Error loading filter from {xmlFile}: {ex.Message}");
                        // Continue loading other filters even if one fails
                    }
                }
                
                _log($"[FILTER_MGMT] Loaded {loadedCount} saved filters from directory");
                _updateStatus($"Loaded {loadedCount} saved filters");
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
        /// Loads a filter automatically from default location
        /// </summary>
        public OpeningFilter LoadFilterAuto(string filterName)
        {
            try
            {
                _log($"[FILTER_MGMT] Auto-loading filter: {filterName}");
                
                var filterDir = GetDefaultFilterDirectory();
                var filePath = Path.Combine(filterDir, $"{filterName}.xml");
                
                if (!File.Exists(filePath))
                {
                    _log($"[FILTER_MGMT] Filter file not found: {filePath}");
                    return null;
                }

                var loadedFilter = LoadFilterFromXmlFile(filePath);
                _log($"[FILTER_MGMT] Auto-loaded filter '{filterName}' from {filePath}");
                return loadedFilter;
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error auto-loading filter '{filterName}': {ex.Message}");
                return null;
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
                
                // Get default filter directory
                var filterDir = GetDefaultFilterDirectory();
                if (!Directory.Exists(filterDir))
                {
                    Directory.CreateDirectory(filterDir);
                }
                
                var openDialog = new OpenFileDialog
                {
                    Title = "Load Filter",
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

        private OpeningFilter CreateFilterFromCurrentUIState(string filterName)
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

            // Pull current UI selections via provider if available
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

                var hostTypes = FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke();
                if (hostTypes != null) filter.SelectedHostElementTypes = new List<string>(hostTypes);
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
            if (filterListBox != null)
            {
                filterListBox.Items.Add(filter.Name);
                _filters.Add(filter);
                _log($"[FILTER_MGMT] Added filter '{filter.Name}' to list");
            }
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

        private void RefreshFilterList(ListBox filterListBox)
        {
            if (filterListBox != null)
            {
                filterListBox.Refresh();
                _log("[FILTER_MGMT] Refreshed filter list");
            }
        }

        /// <summary>
        /// Saves a filter to an XML file
        /// </summary>
        public void SaveFilterToXmlFile(OpeningFilter filter, string filePath)
        {
            try
            {
                if (filter == null)
                    throw new ArgumentNullException(nameof(filter));

                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));

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
                    serializer.Serialize(writer, filter);
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
                Log($"[FILTER_MGMT] Saved filter to XML: {filePath}");
            }
            catch (Exception ex)
            {
                Log($"[FILTER_MGMT] Error serializing filter to XML: {ex.Message}");
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

