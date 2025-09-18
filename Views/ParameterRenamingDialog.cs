using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Dialog for managing parameter value renaming conditions
    /// </summary>
    public partial class ParameterRenamingDialog : System.Windows.Forms.Form
    {
        private ParameterRenamingService _renamingService;
        private DataGridView _renamingDataGridView;
        private Button _addButton;
        private Button _editButton;
        private Button _removeButton;
        private Button _importCsvButton;
        private Button _exportCsvButton;
        private Button _loadPredefinedButton;
        private Button _clearAllButton;
        private Button _okButton;
        private Button _cancelButton;
        private ComboBox _parameterFilterComboBox;
        private Label _parameterFilterLabel;
        
        public ParameterRenamingDialog()
        {
            _renamingService = new ParameterRenamingService();
            InitializeComponent();
            LoadRenamingConditions();
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = "Parameter Value Renaming";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Parameter filter
            _parameterFilterLabel = new Label
            {
                Text = "Filter by Parameter:",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(100, 20)
            };
            this.Controls.Add(_parameterFilterLabel);
            
            _parameterFilterComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 8),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _parameterFilterComboBox.SelectedIndexChanged += ParameterFilterComboBox_SelectedIndexChanged;
            this.Controls.Add(_parameterFilterComboBox);
            
            // Data grid view
            _renamingDataGridView = new DataGridView
            {
                Location = new System.Drawing.Point(10, 40),
                Size = new Size(660, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                ReadOnly = true
            };
            
            // Add columns
            _renamingDataGridView.Columns.Add("OriginalValue", "Original Value");
            _renamingDataGridView.Columns.Add("NewValue", "New Value");
            _renamingDataGridView.Columns.Add("ParameterName", "Parameter Name");
            _renamingDataGridView.Columns.Add("IsEnabled", "Enabled");
            
            // Set column properties
            _renamingDataGridView.Columns["OriginalValue"].Width = 150;
            _renamingDataGridView.Columns["NewValue"].Width = 150;
            _renamingDataGridView.Columns["ParameterName"].Width = 200;
            _renamingDataGridView.Columns["IsEnabled"].Width = 80;
            
            _renamingDataGridView.CellDoubleClick += RenamingDataGridView_CellDoubleClick;
            this.Controls.Add(_renamingDataGridView);
            
            // Buttons
            CreateButtons();
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Not all parameter values may use the required naming conventions. For example, the System 'Plumbing' should be described with the abbreviation 'P'. To address this, you can tell conVoid how to rename specific values by adding the conditions. You can also create a *.csv template to import all the renaming conditions.",
                Location = new System.Drawing.Point(10, 350),
                Size = new Size(660, 40),
                ForeColor = System.Drawing.Color.Gray,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(descriptionLabel);
        }
        
        private void CreateButtons()
        {
            // Add button
            _addButton = new Button
            {
                Text = "Add",
                Location = new System.Drawing.Point(10, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _addButton.Click += AddButton_Click;
            this.Controls.Add(_addButton);
            
            // Edit button
            _editButton = new Button
            {
                Text = "Edit",
                Location = new System.Drawing.Point(100, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _editButton.Click += EditButton_Click;
            this.Controls.Add(_editButton);
            
            // Remove button
            _removeButton = new Button
            {
                Text = "Remove",
                Location = new System.Drawing.Point(190, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _removeButton.Click += RemoveButton_Click;
            this.Controls.Add(_removeButton);
            
            // Import CSV button
            _importCsvButton = new Button
            {
                Text = "Import CSV",
                Location = new System.Drawing.Point(280, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _importCsvButton.Click += ImportCsvButton_Click;
            this.Controls.Add(_importCsvButton);
            
            // Export CSV button
            _exportCsvButton = new Button
            {
                Text = "Export CSV",
                Location = new System.Drawing.Point(370, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _exportCsvButton.Click += ExportCsvButton_Click;
            this.Controls.Add(_exportCsvButton);
            
            // Load Predefined button
            _loadPredefinedButton = new Button
            {
                Text = "Load Predefined",
                Location = new System.Drawing.Point(460, 400),
                Size = new Size(100, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _loadPredefinedButton.Click += LoadPredefinedButton_Click;
            this.Controls.Add(_loadPredefinedButton);
            
            // Clear All button
            _clearAllButton = new Button
            {
                Text = "Clear All",
                Location = new System.Drawing.Point(570, 400),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _clearAllButton.Click += ClearAllButton_Click;
            this.Controls.Add(_clearAllButton);
            
            // OK button
            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(520, 400),
                Size = new Size(80, 25),
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            this.Controls.Add(_okButton);
            
            // Cancel button
            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(610, 400),
                Size = new Size(80, 25),
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            this.Controls.Add(_cancelButton);
        }
        
        private void LoadRenamingConditions()
        {
            try
            {
                // Load all renaming conditions
                var allConditions = _renamingService.GetAllRenamingConditions();
                
                // Populate parameter filter
                var parameterNames = allConditions
                    .Select(c => c.ParameterName)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
                
                _parameterFilterComboBox.Items.Clear();
                _parameterFilterComboBox.Items.Add("All Parameters");
                foreach (var paramName in parameterNames)
                {
                    _parameterFilterComboBox.Items.Add(paramName);
                }
                _parameterFilterComboBox.SelectedIndex = 0;
                
                // Refresh data grid
                RefreshDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading renaming conditions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void RefreshDataGridView()
        {
            try
            {
                _renamingDataGridView.Rows.Clear();
                
                var conditions = GetFilteredConditions();
                foreach (var condition in conditions)
                {
                    _renamingDataGridView.Rows.Add(
                        condition.OriginalValue,
                        condition.NewValue,
                        condition.ParameterName,
                        condition.IsEnabled
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing data grid: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private List<RenamingCondition> GetFilteredConditions()
        {
            var allConditions = _renamingService.GetAllRenamingConditions();
            
            if (_parameterFilterComboBox.SelectedItem == null || 
                _parameterFilterComboBox.SelectedItem.ToString() == "All Parameters")
            {
                return allConditions;
            }
            
            var selectedParameter = _parameterFilterComboBox.SelectedItem.ToString();
            return allConditions.Where(c => c.ParameterName == selectedParameter).ToList();
        }
        
        private void AddButton_Click(object sender, EventArgs e)
        {
            using (var addDialog = new AddRenamingConditionDialog())
            {
                if (addDialog.ShowDialog() == DialogResult.OK)
                {
                    var condition = addDialog.GetRenamingCondition();
                    _renamingService.AddRenamingCondition(condition);
                    LoadRenamingConditions(); // Reload to update parameter filter
                }
            }
        }
        
        private void EditButton_Click(object sender, EventArgs e)
        {
            if (_renamingDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a renaming condition to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            var selectedRow = _renamingDataGridView.SelectedRows[0];
            var originalValue = selectedRow.Cells["OriginalValue"].Value.ToString();
            var newValue = selectedRow.Cells["NewValue"].Value.ToString();
            var parameterName = selectedRow.Cells["ParameterName"].Value?.ToString() ?? string.Empty;
            var isEnabled = Convert.ToBoolean(selectedRow.Cells["IsEnabled"].Value);
            
            var condition = new RenamingCondition(originalValue, newValue, parameterName)
            {
                IsEnabled = isEnabled
            };
            
            using (var editDialog = new AddRenamingConditionDialog(condition))
            {
                if (editDialog.ShowDialog() == DialogResult.OK)
                {
                    // Remove old condition
                    _renamingService.RemoveRenamingCondition(originalValue, parameterName);
                    
                    // Add updated condition
                    var updatedCondition = editDialog.GetRenamingCondition();
                    _renamingService.AddRenamingCondition(updatedCondition);
                    LoadRenamingConditions(); // Reload to update parameter filter
                }
            }
        }
        
        private void RemoveButton_Click(object sender, EventArgs e)
        {
            if (_renamingDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a renaming condition to remove.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            var result = MessageBox.Show("Are you sure you want to remove the selected renaming condition?", 
                "Confirm Removal", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                var selectedRow = _renamingDataGridView.SelectedRows[0];
                var originalValue = selectedRow.Cells["OriginalValue"].Value.ToString();
                var parameterName = selectedRow.Cells["ParameterName"].Value?.ToString() ?? string.Empty;
                
                _renamingService.RemoveRenamingCondition(originalValue, parameterName);
                LoadRenamingConditions(); // Reload to update parameter filter
            }
        }
        
        private void ImportCsvButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.Title = "Import Renaming Conditions";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var importedConditions = _renamingService.ImportFromCsv(openFileDialog.FileName);
                        LoadRenamingConditions(); // Reload to update parameter filter
                        MessageBox.Show($"Imported {importedConditions.Count} renaming conditions.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error importing CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        
        private void ExportCsvButton_Click(object sender, EventArgs e)
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                saveFileDialog.Title = "Export Renaming Conditions";
                saveFileDialog.FileName = "renaming_conditions.csv";
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var conditions = GetFilteredConditions();
                        var success = _renamingService.ExportToCsv(conditions, saveFileDialog.FileName);
                        
                        if (success)
                        {
                            MessageBox.Show($"Exported {conditions.Count} renaming conditions.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Failed to export renaming conditions.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        
        private void LoadPredefinedButton_Click(object sender, EventArgs e)
        {
            try
            {
                _renamingService.LoadPredefinedRenamingConditions();
                LoadRenamingConditions(); // Reload to update parameter filter
                MessageBox.Show("Predefined renaming conditions loaded.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading predefined conditions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void ClearAllButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to clear all renaming conditions?", 
                "Confirm Clear All", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                _renamingService.ClearAllRenamingConditions();
                LoadRenamingConditions(); // Reload to update parameter filter
            }
        }
        
        private void ParameterFilterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshDataGridView();
        }
        
        private void RenamingDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                EditButton_Click(sender, e);
            }
        }
        
        public List<RenamingCondition> GetRenamingConditions()
        {
            return _renamingService.GetAllRenamingConditions();
        }
    }
    
    /// <summary>
    /// Dialog for adding or editing a renaming condition
    /// </summary>
    public partial class AddRenamingConditionDialog : System.Windows.Forms.Form
    {
        private TextBox _originalValueTextBox;
        private TextBox _newValueTextBox;
        private ComboBox _parameterNameComboBox;
        private CheckBox _isEnabledCheckBox;
        private Button _okButton;
        private Button _cancelButton;
        private RenamingCondition _condition;
        
        public AddRenamingConditionDialog() : this(null)
        {
        }
        
        public AddRenamingConditionDialog(RenamingCondition condition)
        {
            _condition = condition;
            InitializeComponent();
            
            if (condition != null)
            {
                LoadCondition(condition);
            }
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = _condition == null ? "Add Renaming Condition" : "Edit Renaming Condition";
            this.Size = new Size(400, 250);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Original Value
            var originalLabel = new Label
            {
                Text = "Original Value:",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(100, 20)
            };
            this.Controls.Add(originalLabel);
            
            _originalValueTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new Size(250, 20)
            };
            this.Controls.Add(_originalValueTextBox);
            
            // New Value
            var newLabel = new Label
            {
                Text = "New Value:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            this.Controls.Add(newLabel);
            
            _newValueTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(250, 20)
            };
            this.Controls.Add(_newValueTextBox);
            
            // Parameter Name
            var parameterLabel = new Label
            {
                Text = "Parameter Name:",
                Location = new System.Drawing.Point(10, 80),
                Size = new Size(100, 20)
            };
            this.Controls.Add(parameterLabel);
            
            _parameterNameComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new Size(250, 20),
                DropDownStyle = ComboBoxStyle.DropDown
            };
            
            // Add common parameter names
            var commonParameters = new List<string>
            {
                "System Abbreviation",
                "System Name",
                "System Type",
                "Fire Rating",
                "Material",
                "Level Name",
                "Comments",
                "Mark",
                "Type Name",
                "Family Name"
            };
            
            foreach (var param in commonParameters)
            {
                _parameterNameComboBox.Items.Add(param);
            }
            
            this.Controls.Add(_parameterNameComboBox);
            
            // Enabled checkbox
            _isEnabledCheckBox = new CheckBox
            {
                Text = "Enabled",
                Location = new System.Drawing.Point(10, 110),
                Size = new Size(100, 20),
                Checked = true
            };
            this.Controls.Add(_isEnabledCheckBox);
            
            // OK button
            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(200, 150),
                Size = new Size(80, 25),
                DialogResult = DialogResult.OK
            };
            _okButton.Click += OkButton_Click;
            this.Controls.Add(_okButton);
            
            // Cancel button
            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(290, 150),
                Size = new Size(80, 25),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(_cancelButton);
        }
        
        private void LoadCondition(RenamingCondition condition)
        {
            _originalValueTextBox.Text = condition.OriginalValue;
            _newValueTextBox.Text = condition.NewValue;
            _parameterNameComboBox.Text = condition.ParameterName;
            _isEnabledCheckBox.Checked = condition.IsEnabled;
        }
        
        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_originalValueTextBox.Text))
            {
                MessageBox.Show("Original value cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if (string.IsNullOrEmpty(_newValueTextBox.Text))
            {
                MessageBox.Show("New value cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if (_originalValueTextBox.Text.Equals(_newValueTextBox.Text, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Original and new values cannot be the same.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            this.DialogResult = DialogResult.OK;
        }
        
        public RenamingCondition GetRenamingCondition()
        {
            return new RenamingCondition(
                _originalValueTextBox.Text,
                _newValueTextBox.Text,
                _parameterNameComboBox.Text
            )
            {
                IsEnabled = _isEnabledCheckBox.Checked
            };
        }
    }
}
