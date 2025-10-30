using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Dialog for managing service type abbreviations
    /// </summary>
    public partial class ServiceTypeAbbreviationDialog : System.Windows.Forms.Form
    {
        private ServiceTypeAbbreviationService _abbreviationService;
        private DataGridView _abbreviationDataGridView;
        private Button _addButton;
        private Button _editButton;
        private Button _removeButton;
        private Button _importCsvButton;
        private Button _exportCsvButton;
        private Button _loadDefaultsButton;
        private Button _clearAllButton;
        private Button _okButton;
        private Button _cancelButton;
        private TextBox _searchTextBox;
        private Label _searchLabel;
        
        public ServiceTypeAbbreviationDialog()
        {
            _abbreviationService = new ServiceTypeAbbreviationService();
            InitializeComponent();
            LoadAbbreviations();
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = "Service Type Abbreviations";
            this.Size = new Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Search box
            _searchLabel = new Label
            {
                Text = "Search:",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(50, 20)
            };
            this.Controls.Add(_searchLabel);
            
            _searchTextBox = new TextBox
            {
                Location = new System.Drawing.Point(70, 8),
                Size = new Size(200, 20)
            };
            _searchTextBox.TextChanged += SearchTextBox_TextChanged;
            this.Controls.Add(_searchTextBox);
            
            // Data grid view
            _abbreviationDataGridView = new DataGridView
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
            _abbreviationDataGridView.Columns.Add("ServiceType", "Service Type");
            _abbreviationDataGridView.Columns.Add("Abbreviation", "Abbreviation");
            _abbreviationDataGridView.Columns.Add("ParameterName", "Parameter Name");
            _abbreviationDataGridView.Columns.Add("IsEnabled", "Enabled");
            
            // Set column properties
            _abbreviationDataGridView.Columns["ServiceType"].Width = 250;
            _abbreviationDataGridView.Columns["Abbreviation"].Width = 150;
            _abbreviationDataGridView.Columns["ParameterName"].Width = 150;
            _abbreviationDataGridView.Columns["IsEnabled"].Width = 80;
            
            _abbreviationDataGridView.CellDoubleClick += AbbreviationDataGridView_CellDoubleClick;
            this.Controls.Add(_abbreviationDataGridView);
            
            // Buttons
            CreateButtons();
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Define abbreviations for service types that will be used when transferring parameters from MEP elements to openings. " +
                       "For example, 'Electrical Distribution Board' can be abbreviated as 'EDB'.",
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
            
            // Load Defaults button
            _loadDefaultsButton = new Button
            {
                Text = "Load Defaults",
                Location = new System.Drawing.Point(460, 400),
                Size = new Size(100, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _loadDefaultsButton.Click += LoadDefaultsButton_Click;
            this.Controls.Add(_loadDefaultsButton);
            
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
        
        private void LoadAbbreviations()
        {
            try
            {
                RefreshDataGridView();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading abbreviations: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void RefreshDataGridView()
        {
            try
            {
                _abbreviationDataGridView.Rows.Clear();
                
                var abbreviations = GetFilteredAbbreviations();
                foreach (var abbreviation in abbreviations)
                {
                    _abbreviationDataGridView.AddRow(
                        abbreviation.ServiceType,
                        abbreviation.Abbreviation,
                        abbreviation.ParameterName,
                        abbreviation.IsEnabled
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing data grid: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ...existing code...
        
        private List<ServiceTypeAbbreviation> GetFilteredAbbreviations()
        {
            var allAbbreviations = GetAllAbbreviationsFromService();
            
            if (string.IsNullOrEmpty(_searchTextBox.Text))
            {
                return allAbbreviations;
            }
            
            var searchText = _searchTextBox.Text.ToLower();
            return allAbbreviations.Where(a => 
                a.ServiceType.ToLower().Contains(searchText) ||
                a.Abbreviation.ToLower().Contains(searchText) ||
                a.ParameterName.ToLower().Contains(searchText)
            ).ToList();
        }
        
        private List<ServiceTypeAbbreviation> GetAllAbbreviationsFromService()
        {
            var abbreviations = new List<ServiceTypeAbbreviation>();
            var serviceAbbreviations = _abbreviationService.GetAllAbbreviations();
            
            foreach (var kvp in serviceAbbreviations)
            {
                abbreviations.Add(new ServiceTypeAbbreviation(kvp.Key, kvp.Value));
            }
            
            return abbreviations.OrderBy(a => a.ServiceType).ToList();
        }
        
        private void AddButton_Click(object sender, EventArgs e)
        {
            using (var addDialog = new AddServiceTypeAbbreviationDialog())
            {
                if (addDialog.ShowDialog() == DialogResult.OK)
                {
                    var abbreviation = addDialog.GetServiceTypeAbbreviation();
                    _abbreviationService.AddAbbreviation(abbreviation.ServiceType, abbreviation.Abbreviation);
                    RefreshDataGridView();
                }
            }
        }
        
        private void EditButton_Click(object sender, EventArgs e)
        {
            if (_abbreviationDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select an abbreviation to edit.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            var selectedRow = _abbreviationDataGridView.SelectedRows[0];
            var serviceType = selectedRow.Cells["ServiceType"].Value.ToString();
            var abbreviation = selectedRow.Cells["Abbreviation"].Value.ToString();
            var parameterName = selectedRow.Cells["ParameterName"].Value?.ToString() ?? string.Empty;
            var isEnabled = Convert.ToBoolean(selectedRow.Cells["IsEnabled"].Value);
            
            var abbreviationObj = new ServiceTypeAbbreviation(serviceType, abbreviation, parameterName)
            {
                IsEnabled = isEnabled
            };
            
            using (var editDialog = new AddServiceTypeAbbreviationDialog(abbreviationObj))
            {
                if (editDialog.ShowDialog() == DialogResult.OK)
                {
                    // Remove old abbreviation
                    _abbreviationService.RemoveAbbreviation(serviceType);
                    
                    // Add updated abbreviation
                    var updatedAbbreviation = editDialog.GetServiceTypeAbbreviation();
                    _abbreviationService.AddAbbreviation(updatedAbbreviation.ServiceType, updatedAbbreviation.Abbreviation);
                    RefreshDataGridView();
                }
            }
        }
        
        private void RemoveButton_Click(object sender, EventArgs e)
        {
            if (_abbreviationDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select an abbreviation to remove.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            var result = MessageBox.Show("Are you sure you want to remove the selected abbreviation?", 
                "Confirm Removal", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                var selectedRow = _abbreviationDataGridView.SelectedRows[0];
                var serviceType = selectedRow.Cells["ServiceType"].Value.ToString();
                
                _abbreviationService.RemoveAbbreviation(serviceType);
                RefreshDataGridView();
            }
        }
        
        private void ImportCsvButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.Title = "Import Service Type Abbreviations";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var importedAbbreviations = _abbreviationService.ImportFromCsv(openFileDialog.FileName);
                        RefreshDataGridView();
                        MessageBox.Show($"Imported {importedAbbreviations.Count} service type abbreviations.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                saveFileDialog.Title = "Export Service Type Abbreviations";
                saveFileDialog.FileName = "service_type_abbreviations.csv";
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var abbreviations = GetFilteredAbbreviations();
                        var success = _abbreviationService.ExportToCsv(abbreviations, saveFileDialog.FileName);
                        
                        if (success)
                        {
                            MessageBox.Show($"Exported {abbreviations.Count} service type abbreviations.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Failed to export service type abbreviations.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting CSV: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        
        private void LoadDefaultsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Clear existing and reload defaults
                _abbreviationService.ClearAllAbbreviations();
                var newService = new ServiceTypeAbbreviationService(); // This loads defaults
                var defaultAbbreviations = newService.GetAllAbbreviations();
                
                foreach (var kvp in defaultAbbreviations)
                {
                    _abbreviationService.AddAbbreviation(kvp.Key, kvp.Value);
                }
                
                RefreshDataGridView();
                MessageBox.Show("Default service type abbreviations loaded.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading default abbreviations: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void ClearAllButton_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Are you sure you want to clear all service type abbreviations?", 
                "Confirm Clear All", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                _abbreviationService.ClearAllAbbreviations();
                RefreshDataGridView();
            }
        }
        
        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            RefreshDataGridView();
        }
        
        private void AbbreviationDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                EditButton_Click(sender, e);
            }
        }
        
        public ServiceTypeAbbreviationService GetAbbreviationService()
        {
            return _abbreviationService;
        }
    }
    
    /// <summary>
    /// Dialog for adding or editing a service type abbreviation
    /// </summary>
    public partial class AddServiceTypeAbbreviationDialog : System.Windows.Forms.Form
    {
        private TextBox _serviceTypeTextBox;
        private TextBox _abbreviationTextBox;
        private ComboBox _parameterNameComboBox;
        private CheckBox _isEnabledCheckBox;
        private Button _okButton;
        private Button _cancelButton;
        private ServiceTypeAbbreviation _abbreviation;
        
        public AddServiceTypeAbbreviationDialog() : this(null)
        {
        }
        
        public AddServiceTypeAbbreviationDialog(ServiceTypeAbbreviation abbreviation)
        {
            _abbreviation = abbreviation;
            InitializeComponent();
            
            if (abbreviation != null)
            {
                LoadAbbreviation(abbreviation);
            }
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = _abbreviation == null ? "Add Service Type Abbreviation" : "Edit Service Type Abbreviation";
            this.Size = new Size(400, 250);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Service Type
            var serviceTypeLabel = new Label
            {
                Text = "Service Type:",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(100, 20)
            };
            this.Controls.Add(serviceTypeLabel);
            
            _serviceTypeTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 18),
                Size = new Size(250, 20)
            };
            this.Controls.Add(_serviceTypeTextBox);
            
            // Abbreviation
            var abbreviationLabel = new Label
            {
                Text = "Abbreviation:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            this.Controls.Add(abbreviationLabel);
            
            _abbreviationTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(250, 20)
            };
            this.Controls.Add(_abbreviationTextBox);
            
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
                "Family Name",
                "Type Name",
                "Comments"
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
        
        private void LoadAbbreviation(ServiceTypeAbbreviation abbreviation)
        {
            _serviceTypeTextBox.Text = abbreviation.ServiceType;
            _abbreviationTextBox.Text = abbreviation.Abbreviation;
            _parameterNameComboBox.Text = abbreviation.ParameterName;
            _isEnabledCheckBox.Checked = abbreviation.IsEnabled;
        }
        
        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_serviceTypeTextBox.Text))
            {
                MessageBox.Show("Service type cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if (string.IsNullOrEmpty(_abbreviationTextBox.Text))
            {
                MessageBox.Show("Abbreviation cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            if (_serviceTypeTextBox.Text.Equals(_abbreviationTextBox.Text, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Service type and abbreviation cannot be the same.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            this.DialogResult = DialogResult.OK;
        }
        
        public ServiceTypeAbbreviation GetServiceTypeAbbreviation()
        {
            return new ServiceTypeAbbreviation(
                _serviceTypeTextBox.Text,
                _abbreviationTextBox.Text,
                _parameterNameComboBox.Text
            )
            {
                IsEnabled = _isEnabledCheckBox.Checked
            };
        }
    }
}
