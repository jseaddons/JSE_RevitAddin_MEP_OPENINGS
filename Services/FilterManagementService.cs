using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
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

                RefreshFilterList(filterListBox);
                
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
                    FileName = $"{selectedFilter.Name}.xml"
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
        /// Loads a filter from XML file
        /// </summary>
        public void LoadFilter(ListBox filterListBox)
        {
            try
            {
                _log("[FILTER_MGMT] Loading filter");
                
                var openDialog = new OpenFileDialog
                {
                    Title = "Load Filter",
                    Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
                    DefaultExt = "xml"
                };

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    var loadedFilter = LoadFilterFromXmlFile(openDialog.FileName);
                    
                    if (loadedFilter != null)
                    {
                        AddFilterToList(filterListBox, loadedFilter);
                        
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
            return new OpeningFilter
            {
                Name = filterName,
                Category = Models.MepCategory.Ducts, // Default category
                OpeningType = Models.OpeningType.RectangularSleeves, // Default type
                IsEnabled = true,
                LastModified = DateTime.Now,
                ClashZoneStorage = null // Will be populated when clash detection is run
            };
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

        private void SaveFilterToXmlFile(OpeningFilter filter, string filePath)
        {
            try
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
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

        private OpeningFilter LoadFilterFromXmlFile(string filePath)
        {
            try
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new System.IO.StreamReader(filePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    _log($"[FILTER_MGMT] Loaded filter from XML: {filePath}");
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
