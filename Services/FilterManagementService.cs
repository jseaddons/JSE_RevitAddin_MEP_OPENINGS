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
        private List<OpeningFilter> _filters;

        public FilterManagementService(Action<string> log, Action<string> updateStatus)
        {
            _log = log;
            _updateStatus = updateStatus;
            _filters = new List<OpeningFilter>();
        }

        public FilterManagementService(Document document, Action<string> log, Action<string> updateStatus)
        {
            _log = log;
            _updateStatus = updateStatus;
            _filters = new List<OpeningFilter>();
            // Document parameter is accepted but not used in this constructor
            // It's provided for compatibility with RefreshService
        }

        #region Public Methods

        /// <summary>
        /// Creates a new filter from current UI state
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
                
                _log($"[FILTER_MGMT] Created new filter: {filterName}");
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
                    LastModified = DateTime.Now
                };

                AddFilterToList(filterListBox, copiedFilter);
                
                _log($"[FILTER_MGMT] Copied filter '{selectedFilter.Name}' to '{newName}'");
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
                
                _log($"[FILTER_MGMT] Renamed filter '{oldName}' to '{newName}'");
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
        /// </summary>
        public void SeedDefaultFilters(ListBox filterListBox, IEnumerable<string> defaultNames)
        {
            if (filterListBox == null) return;
            foreach (var name in defaultNames)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (_filters.Any(f => f.Name == name) || filterListBox.Items.Contains(name)) continue;
                var filter = CreateFilterFromCurrentUIState(name);
                AddFilterToList(filterListBox, filter);
            }
        }

        /// <summary>
        /// Deletes the selected filter
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
        /// Saves the selected filter to XML file
        /// </summary>
        public void SaveFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Saving filter");
                
                var selectedFilter = GetSelectedFilter(filterListBox);
                if (selectedFilter == null)
                {
                    ShowWarning("Please select a filter to save.");
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Save Filter",
                    Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
                    DefaultExt = "xml",
                    FileName = $"{selectedFilter.Name}.xml",
                    InitialDirectory = GetDefaultFilterDirectory()
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    SaveFilterToXmlFile(selectedFilter, saveDialog.FileName);
                    
                    _log($"[FILTER_MGMT] Saved filter '{selectedFilter.Name}' to: {saveDialog.FileName}");
                    _updateStatus($"Saved filter to: {System.IO.Path.GetFileName(saveDialog.FileName)}");
                    
                    MessageBox.Show($"Filter '{selectedFilter.Name}' saved successfully!", "Success", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                // Normalize intersection coordinates to avoid 0,0,0 in XML
                if (filter?.ClashZoneStorage?.ClashZones != null)
                {
                    foreach (var z in filter.ClashZoneStorage.ClashZones)
                    {
                        if (z == null) continue;
                        bool hasIP = z.IntersectionPoint != null;
                        bool isIPZero = hasIP && Math.Abs(z.IntersectionPoint.X) < 1e-9 && Math.Abs(z.IntersectionPoint.Y) < 1e-9 && Math.Abs(z.IntersectionPoint.Z) < 1e-9;
                        bool hasSPP = z.SleevePlacementPoint != null;
                        bool isSPPZero = hasSPP && Math.Abs(z.SleevePlacementPoint.X) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Y) < 1e-9 && Math.Abs(z.SleevePlacementPoint.Z) < 1e-9;
                        
                        if (hasIP && !isIPZero)
                        {
                            z.IntersectionPointX = z.IntersectionPoint.X;
                            z.IntersectionPointY = z.IntersectionPoint.Y;
                            z.IntersectionPointZ = z.IntersectionPoint.Z;
                        }
                        else if (hasSPP && !isSPPZero)
                        {
                            z.IntersectionPointX = z.SleevePlacementPoint.X;
                            z.IntersectionPointY = z.SleevePlacementPoint.Y;
                            z.IntersectionPointZ = z.SleevePlacementPoint.Z;
                        }
                        else if (z.ClashBoundingBox != null)
                        {
                            var center = (z.ClashBoundingBox.Min + z.ClashBoundingBox.Max) / 2.0;
                            z.IntersectionPointX = center.X;
                            z.IntersectionPointY = center.Y;
                            z.IntersectionPointZ = center.Z;
                        }
                    }
                }
                using (var writer = new System.IO.StreamWriter(filePath))
                {
                    serializer.Serialize(writer, filter);
                }
                _log($"[FILTER_MGMT] Saved filter to XML: {filePath}");
            }
            catch (Exception ex)
            {
                _log($"[FILTER_MGMT] Error serializing filter to XML: {ex.Message}");
                throw;
            }
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
                    
                    // ✅ CRITICAL FIX: Reconstruct SleevePlacementPoint from XML-serializable properties
                    if (filter?.ClashZoneStorage?.ClashZones != null)
                    {
                        foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                        {
                            clashZone.EnsureSleevePlacementPointReconstructed();
                        }
                        _log($"[FILTER_MGMT] Reconstructed SleevePlacementPoint for {filter.ClashZoneStorage.ClashZones.Count} clash zones");
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

