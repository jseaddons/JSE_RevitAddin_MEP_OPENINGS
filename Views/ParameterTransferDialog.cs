using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Dialog for configuring parameter transfer operations
    /// </summary>
    public partial class ParameterTransferDialog : System.Windows.Forms.Form
    {
        private Document _document;
        private ParameterTransferService _transferService;
        private ParameterMappingService _mappingService;
        private ParameterRenamingService _renamingService;
        private ParameterTransferConfiguration _configuration;
        
        // UI Controls
        private TabControl _mainTabControl;
        private TabPage _referenceTab;
        private TabPage _hostTab;
        private TabPage _levelTab;
        private TabPage _modelTab;
        private TabPage _renamingTab;
        private TabPage _serviceSizeTab;
        
        // Reference Element to Opening controls
        private CheckBox _referenceEnabledCheckBox;
        private ComboBox _referenceSourceComboBox;
        private ComboBox _referenceTargetComboBox;
        
        // Host to Opening controls
        private CheckBox _hostEnabledCheckBox;
        private ComboBox _hostSourceComboBox;
        private ComboBox _hostTargetComboBox;
        private TextBox _hostSeparatorTextBox;
        
        // Level to Opening controls
        private CheckBox _levelEnabledCheckBox;
        private ComboBox _levelSourceComboBox;
        private ComboBox _levelTargetComboBox;
        
        // Model Name controls
        private CheckBox _modelEnabledCheckBox;
        private ComboBox _modelTargetComboBox;
        
        // Renaming controls
        private DataGridView _renamingDataGridView;
        private Button _addRenamingButton;
        private Button _removeRenamingButton;
        private Button _importCsvButton;
        private Button _exportCsvButton;
        private Button _loadPredefinedButton;
        private Button _manageAbbreviationsButton;
        private Button _configureOpeningParametersButton;
        
        // Service Size Calculation controls
        private CheckBox _serviceSizeEnabledCheckBox;
        private ComboBox _serviceSizeTargetComboBox;
        private TextBox _clearanceTextBox;
        private TextBox _clearanceSuffixTextBox;
        
        // Action buttons
        private Button _okButton;
        private Button _cancelButton;
        private Button _previewButton;
        private Button _helpButton;
        
        public ParameterTransferDialog(Document document)
        {
            _document = document;
            _transferService = new ParameterTransferService();
            _mappingService = new ParameterMappingService();
            _renamingService = new ParameterRenamingService();
            _configuration = new ParameterTransferConfiguration();
            
            InitializeComponent();
            LoadParameterData();
            LoadPredefinedMappings();
        }
        
        private void InitializeComponent()
        {
            // Form properties
            this.Text = "Parameter Transfer Settings";
            this.Size = new Size(600, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            
            // Main tab control
            _mainTabControl = new TabControl
            {
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(560, 400),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(_mainTabControl);
            
            // Create tabs
            CreateReferenceTab();
            CreateHostTab();
            CreateLevelTab();
            CreateModelTab();
            CreateRenamingTab();
            
            // Action buttons
            CreateActionButtons();
        }
        
        private void CreateReferenceTab()
        {
            _referenceTab = new TabPage("Reference Element to Openings");
            _mainTabControl.TabPages.Add(_referenceTab);
            
            // Reference Element to Opening section
            var referenceGroupBox = new GroupBox
            {
                Text = "Reference Element to Openings",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(520, 120),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _referenceTab.Controls.Add(referenceGroupBox);
            
            // Enabled checkbox
            _referenceEnabledCheckBox = new CheckBox
            {
                Text = "Transfer from Reference Elements",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(200, 20),
                Checked = true
            };
            referenceGroupBox.Controls.Add(_referenceEnabledCheckBox);
            
            // Source parameter
            var sourceLabel = new Label
            {
                Text = "Source Parameter:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            referenceGroupBox.Controls.Add(sourceLabel);
            
            _referenceSourceComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            referenceGroupBox.Controls.Add(_referenceSourceComboBox);
            
            // Target parameter
            var targetLabel = new Label
            {
                Text = "Target Parameter:",
                Location = new System.Drawing.Point(10, 80),
                Size = new Size(100, 20)
            };
            referenceGroupBox.Controls.Add(targetLabel);
            
            _referenceTargetComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            referenceGroupBox.Controls.Add(_referenceTargetComboBox);
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Transfer parameter values from MEP elements (Provision for Voids) to the openings.",
                Location = new System.Drawing.Point(10, 110),
                Size = new Size(500, 40),
                ForeColor = System.Drawing.Color.Gray
            };
            _referenceTab.Controls.Add(descriptionLabel);
        }
        
        private void CreateHostTab()
        {
            _hostTab = new TabPage("Host to Opening");
            _mainTabControl.TabPages.Add(_hostTab);
            
            // Host to Opening section
            var hostGroupBox = new GroupBox
            {
                Text = "Host to Opening",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(520, 150),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _hostTab.Controls.Add(hostGroupBox);
            
            // Enabled checkbox
            _hostEnabledCheckBox = new CheckBox
            {
                Text = "Transfer from Host Elements",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(200, 20),
                Checked = true
            };
            hostGroupBox.Controls.Add(_hostEnabledCheckBox);
            
            // Source parameter
            var sourceLabel = new Label
            {
                Text = "Source Parameter:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            hostGroupBox.Controls.Add(sourceLabel);
            
            _hostSourceComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            hostGroupBox.Controls.Add(_hostSourceComboBox);
            
            // Target parameter
            var targetLabel = new Label
            {
                Text = "Target Parameter:",
                Location = new System.Drawing.Point(10, 80),
                Size = new Size(100, 20)
            };
            hostGroupBox.Controls.Add(targetLabel);
            
            _hostTargetComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            hostGroupBox.Controls.Add(_hostTargetComboBox);
            
            // Separator
            var separatorLabel = new Label
            {
                Text = "Separator:",
                Location = new System.Drawing.Point(10, 110),
                Size = new Size(100, 20)
            };
            hostGroupBox.Controls.Add(separatorLabel);
            
            _hostSeparatorTextBox = new TextBox
            {
                Location = new System.Drawing.Point(120, 108),
                Size = new Size(50, 20),
                Text = ";"
            };
            hostGroupBox.Controls.Add(_hostSeparatorTextBox);
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Transfer Host properties such as fire rating from the wall to the openings. If the opening penetrates multiple Hosts, the separator combines multiple values in one parameter.",
                Location = new System.Drawing.Point(10, 140),
                Size = new Size(500, 40),
                ForeColor = System.Drawing.Color.Gray
            };
            _hostTab.Controls.Add(descriptionLabel);
        }
        
        private void CreateLevelTab()
        {
            _levelTab = new TabPage("Level to Openings");
            _mainTabControl.TabPages.Add(_levelTab);
            
            // Level to Opening section
            var levelGroupBox = new GroupBox
            {
                Text = "Level to Openings",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(520, 120),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _levelTab.Controls.Add(levelGroupBox);
            
            // Enabled checkbox
            _levelEnabledCheckBox = new CheckBox
            {
                Text = "Transfer from Level",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(200, 20),
                Checked = true
            };
            levelGroupBox.Controls.Add(_levelEnabledCheckBox);
            
            // Source parameter
            var sourceLabel = new Label
            {
                Text = "Source Parameter:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            levelGroupBox.Controls.Add(sourceLabel);
            
            _levelSourceComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            levelGroupBox.Controls.Add(_levelSourceComboBox);
            
            // Target parameter
            var targetLabel = new Label
            {
                Text = "Target Parameter:",
                Location = new System.Drawing.Point(10, 80),
                Size = new Size(100, 20)
            };
            levelGroupBox.Controls.Add(targetLabel);
            
            _levelTargetComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 78),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            levelGroupBox.Controls.Add(_levelTargetComboBox);
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Transfer level parameter values such as the level name to the openings.",
                Location = new System.Drawing.Point(10, 110),
                Size = new Size(500, 40),
                ForeColor = System.Drawing.Color.Gray
            };
            _levelTab.Controls.Add(descriptionLabel);
        }
        
        private void CreateModelTab()
        {
            _modelTab = new TabPage("Model Information");
            _mainTabControl.TabPages.Add(_modelTab);
            
            // Model Name section
            var modelGroupBox = new GroupBox
            {
                Text = "Model Information",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(520, 120),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _modelTab.Controls.Add(modelGroupBox);
            
            // Enabled checkbox
            _modelEnabledCheckBox = new CheckBox
            {
                Text = "Transfer Model Name to Opening",
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(200, 20),
                Checked = true
            };
            modelGroupBox.Controls.Add(_modelEnabledCheckBox);
            
            // Target parameter
            var targetLabel = new Label
            {
                Text = "Target Parameter:",
                Location = new System.Drawing.Point(10, 50),
                Size = new Size(100, 20)
            };
            modelGroupBox.Controls.Add(targetLabel);
            
            _modelTargetComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(120, 48),
                Size = new Size(200, 20),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            modelGroupBox.Controls.Add(_modelTargetComboBox);
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Additionally, we offer a built-in parameter called 'Model Name' that enables you to transfer the name of the models to the openings.",
                Location = new System.Drawing.Point(10, 80),
                Size = new Size(500, 40),
                ForeColor = System.Drawing.Color.Gray
            };
            _modelTab.Controls.Add(descriptionLabel);
        }
        
        private void CreateRenamingTab()
        {
            _renamingTab = new TabPage("Parameter Value Renaming");
            _mainTabControl.TabPages.Add(_renamingTab);
            
            // Renaming section
            var renamingGroupBox = new GroupBox
            {
                Text = "Renaming Conditions",
                Location = new System.Drawing.Point(10, 10),
                Size = new Size(520, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            _renamingTab.Controls.Add(renamingGroupBox);
            
            // Data grid view
            _renamingDataGridView = new DataGridView
            {
                Location = new System.Drawing.Point(10, 20),
                Size = new Size(500, 200),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            
            // Add columns
            _renamingDataGridView.Columns.Add("OriginalValue", "Original Value");
            _renamingDataGridView.Columns.Add("NewValue", "New Value");
            _renamingDataGridView.Columns.Add("ParameterName", "Parameter Name");
            _renamingDataGridView.Columns.Add("IsEnabled", "Enabled");
            
            // Set column properties
            _renamingDataGridView.Columns["OriginalValue"].Width = 150;
            _renamingDataGridView.Columns["NewValue"].Width = 150;
            _renamingDataGridView.Columns["ParameterName"].Width = 150;
            _renamingDataGridView.Columns["IsEnabled"].Width = 80;
            
            renamingGroupBox.Controls.Add(_renamingDataGridView);
            
            // Buttons
            _addRenamingButton = new Button
            {
                Text = "Add",
                Location = new System.Drawing.Point(10, 230),
                Size = new Size(80, 25)
            };
            _addRenamingButton.Click += AddRenamingButton_Click;
            renamingGroupBox.Controls.Add(_addRenamingButton);
            
            _removeRenamingButton = new Button
            {
                Text = "Remove",
                Location = new System.Drawing.Point(100, 230),
                Size = new Size(80, 25)
            };
            _removeRenamingButton.Click += RemoveRenamingButton_Click;
            renamingGroupBox.Controls.Add(_removeRenamingButton);
            
            _importCsvButton = new Button
            {
                Text = "Import CSV",
                Location = new System.Drawing.Point(190, 230),
                Size = new Size(80, 25)
            };
            _importCsvButton.Click += ImportCsvButton_Click;
            renamingGroupBox.Controls.Add(_importCsvButton);
            
            _exportCsvButton = new Button
            {
                Text = "Export CSV",
                Location = new System.Drawing.Point(280, 230),
                Size = new Size(80, 25)
            };
            _exportCsvButton.Click += ExportCsvButton_Click;
            renamingGroupBox.Controls.Add(_exportCsvButton);
            
            _loadPredefinedButton = new Button
            {
                Text = "Load Predefined",
                Location = new System.Drawing.Point(370, 230),
                Size = new Size(100, 25)
            };
            _loadPredefinedButton.Click += LoadPredefinedButton_Click;
            renamingGroupBox.Controls.Add(_loadPredefinedButton);
            
            // Manage Abbreviations button
            _manageAbbreviationsButton = new Button
            {
                Text = "Manage Abbreviations",
                Location = new System.Drawing.Point(10, 300),
                Size = new Size(120, 25)
            };
            _manageAbbreviationsButton.Click += ManageAbbreviationsButton_Click;
            renamingGroupBox.Controls.Add(_manageAbbreviationsButton);
            
            // Configure Opening Parameters button
            _configureOpeningParametersButton = new Button
            {
                Text = "Configure Opening Parameters",
                Location = new System.Drawing.Point(140, 300),
                Size = new Size(150, 25)
            };
            _configureOpeningParametersButton.Click += ConfigureOpeningParametersButton_Click;
            renamingGroupBox.Controls.Add(_configureOpeningParametersButton);
            
            // Description
            var descriptionLabel = new Label
            {
                Text = "Not all parameter values may use the required naming conventions. For example, the System 'Plumbing' should be described with the abbreviation 'P'. To address this, you can tell conVoid how to rename specific values by adding the conditions.",
                Location = new System.Drawing.Point(10, 260),
                Size = new Size(500, 40),
                ForeColor = System.Drawing.Color.Gray
            };
            _renamingTab.Controls.Add(descriptionLabel);
        }
        
        private void CreateActionButtons()
        {
            // OK button
            _okButton = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(350, 420),
                Size = new Size(80, 25),
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _okButton.Click += OkButton_Click;
            this.Controls.Add(_okButton);
            
            // Cancel button
            _cancelButton = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(440, 420),
                Size = new Size(80, 25),
                DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            this.Controls.Add(_cancelButton);
            
            // Preview button
            _previewButton = new Button
            {
                Text = "Preview",
                Location = new System.Drawing.Point(260, 420),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _previewButton.Click += PreviewButton_Click;
            this.Controls.Add(_previewButton);
            
            // Help button
            _helpButton = new Button
            {
                Text = "Help",
                Location = new System.Drawing.Point(170, 420),
                Size = new Size(80, 25),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _helpButton.Click += HelpButton_Click;
            this.Controls.Add(_helpButton);
        }
        
        private void LoadParameterData()
        {
            try
            {
                // Use the same working parameter extraction logic as EmergencyMainDialog
                LoadMepParameters();
                LoadHostParameters();
                LoadLevelParameters();
                LoadOpeningParameters();

                // Set default selections
                if (_referenceSourceComboBox.Items.Count > 0)
                    _referenceSourceComboBox.SelectedIndex = 0;
                if (_referenceTargetComboBox.Items.Count > 0)
                    _referenceTargetComboBox.SelectedIndex = 0;
                if (_hostSourceComboBox.Items.Count > 0)
                    _hostSourceComboBox.SelectedIndex = 0;
                if (_hostTargetComboBox.Items.Count > 0)
                    _hostTargetComboBox.SelectedIndex = 0;
                if (_levelSourceComboBox.Items.Count > 0)
                    _levelSourceComboBox.SelectedIndex = 0;
                if (_levelTargetComboBox.Items.Count > 0)
                    _levelTargetComboBox.SelectedIndex = 0;
                if (_modelTargetComboBox.Items.Count > 0)
                    _modelTargetComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading parameter data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Load MEP parameters using the same logic as EmergencyMainDialog
        /// </summary>
        private void LoadMepParameters()
        {
            try
            {
                // Get selected MEP categories (for now, get all common categories)
                var selectedCategories = new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };

                // Use ParameterExtractionService to get parameters from linked files (same as EmergencyMainDialog)
                var parameterService = new Services.ParameterExtractionService();
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(_document);

                var mepParameters = new HashSet<string>();

                if (linkedFiles.Count > 0)
                {
                    // Get parameters from linked files
                    foreach (var linkedFile in linkedFiles)
                    {
                        try
                        {
                            var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                // Convert string categories to enum
                                var mepCategories = selectedCategories
                                    .Select(cat => GetMepCategoryFromName(cat))
                                    .Where(cat => cat.HasValue)
                                    .Select(cat => cat.Value)
                                    .Cast<Services.MepCategory>()
                                    .ToList();

                                var parameters = parameterService.GetParametersForMepCategories(linkedDoc, mepCategories);
                                var parameterNames = parameters.Select(p => p.Name).ToList();

                                // Add unique parameters
                                foreach (var paramName in parameterNames)
                                {
                                    if (!string.IsNullOrEmpty(paramName))
                                    {
                                        mepParameters.Add(paramName);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                        }
                    }
                }
                else
                {
                    // Fallback to current document if no linked files
                    var mepCategories = selectedCategories
                        .Select(cat => GetMepCategoryFromName(cat))
                        .Where(cat => cat.HasValue)
                        .Select(cat => cat.Value)
                        .Cast<Services.MepCategory>()
                        .ToList();

                    var parameters = parameterService.GetParametersForMepCategories(_document, mepCategories);
                    var parameterNames = parameters.Select(p => p.Name).ToList();

                    foreach (var paramName in parameterNames)
                    {
                        if (!string.IsNullOrEmpty(paramName))
                        {
                            mepParameters.Add(paramName);
                        }
                    }
                }

                // Add to comboboxes
                foreach (var param in mepParameters.OrderBy(p => p))
                {
                    _referenceSourceComboBox.Items.Add(param);
                }

                System.Diagnostics.Debug.WriteLine($"Loaded {mepParameters.Count} MEP parameters for parameter transfer");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading MEP parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Load host parameters using the same logic as EmergencyMainDialog
        /// </summary>
        private void LoadHostParameters()
        {
            try
            {
                var hostParameters = new HashSet<string>();

                // Use the same logic as EmergencyMainDialog - check linked files first
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(_document);

                if (linkedFiles.Count > 0)
                {
                    // Get parameters from linked files (same as EmergencyMainDialog)
                    foreach (var linkedFile in linkedFiles)
                    {
                        try
                        {
                            var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                            if (linkedDoc != null)
                            {
                                // Get host parameters from linked document
                                var hostParams = _mappingService.GetHostElementParameters(linkedDoc);
                                foreach (var param in hostParams)
                                {
                                    if (!string.IsNullOrEmpty(param.Name))
                                    {
                                        hostParameters.Add(param.Name);
                                    }
                                }
                                System.Diagnostics.Debug.WriteLine($"Got {hostParams.Count} host parameters from linked file '{linkedFile.FileName}'");
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting host parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                        }
                    }
                }
                else
                {
                    // Fallback to current document if no linked files
                    var hostParams = _mappingService.GetHostElementParameters(_document);
                    foreach (var param in hostParams)
                    {
                        if (!string.IsNullOrEmpty(param.Name))
                        {
                            hostParameters.Add(param.Name);
                        }
                    }
                    System.Diagnostics.Debug.WriteLine($"Got {hostParams.Count} host parameters from current document");
                }

                // If still no parameters found, add common host parameters as fallback
                if (hostParameters.Count == 0)
                {
                    var fallbackParams = new List<string>
                    {
                        "Fire Rating",
                        "Material",
                        "Wall Type",
                        "Floor Type",
                        "Ceiling Type",
                        "Thickness",
                        "Width",
                        "Height",
                        "Area",
                        "Volume",
                        "Comments",
                        "Mark",
                        "Type Name",
                        "Family Name"
                    };

                    foreach (var param in fallbackParams)
                    {
                        hostParameters.Add(param);
                    }

                    System.Diagnostics.Debug.WriteLine("No host parameters found from documents - using fallback parameter list");
                }

                // Add to combobox
                foreach (var param in hostParameters.OrderBy(p => p))
                {
                    _hostSourceComboBox.Items.Add(param);
                }

                System.Diagnostics.Debug.WriteLine($"Loaded {hostParameters.Count} host parameters for parameter transfer from {linkedFiles.Count} linked files");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading host parameters: {ex.Message}");

                // Ultimate fallback - add basic parameters
                var basicParams = new List<string> { "Fire Rating", "Material", "Wall Type", "Comments", "Mark" };
                foreach (var param in basicParams)
                {
                    _hostSourceComboBox.Items.Add(param);
                }
            }
        }

        /// <summary>
        /// Load level parameters
        /// </summary>
        private void LoadLevelParameters()
        {
            try
            {
                var levelParams = _mappingService.GetLevelParameters(_document);
                foreach (var param in levelParams)
                {
                    _levelSourceComboBox.Items.Add(param.Name);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading level parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Load opening parameters from the 4 specific opening families only
        /// </summary>
        private void LoadOpeningParameters()
        {
            try
            {
                var openingParameters = new HashSet<string>();

                // Use the specific method to get parameters from the 4 opening families only
                var openingParams = _mappingService.GetOpeningParametersFromSpecificFamilies(_document);
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] GetOpeningParametersFromSpecificFamilies returned {openingParams.Count} parameters");

                foreach (var param in openingParams)
                {
                    if (!string.IsNullOrEmpty(param.Name))
                    {
                        openingParameters.Add(param.Name);
                        System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Added opening parameter: '{param.Name}' (Shared: {param.IsShared})");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Total unique opening parameters: {openingParameters.Count}");

                // Add to all target comboboxes
                foreach (var param in openingParameters.OrderBy(p => p))
                {
                    _referenceTargetComboBox.Items.Add(param);
                    _hostTargetComboBox.Items.Add(param);
                    _levelTargetComboBox.Items.Add(param);
                    _modelTargetComboBox.Items.Add(param);
                    System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Added to target comboboxes: '{param}'");
                }

                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Loaded {openingParameters.Count} opening parameters from the 4 specific opening families for parameter transfer");

                // Log all parameters for debugging
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Opening parameters list: {string.Join(", ", openingParameters.OrderBy(p => p))}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Error loading opening parameters: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_TRANSFER_DEBUG] Stack trace: {ex.StackTrace}");

                // No fallback - only use the 4 specific families
                System.Diagnostics.Debug.WriteLine("[PARAMETER_TRANSFER_DEBUG] No opening parameters loaded - the 4 specific opening families may not be present in the document");
            }
        }



        /// <summary>
        /// Convert category name string to MepCategory enum
        /// </summary>
        private Models.MepCategory? GetMepCategoryFromName(string categoryName)
        {
            return categoryName.ToLower() switch
            {
                "pipes" or "pipe" => Models.MepCategory.Pipes,
                "ducts" or "duct" => Models.MepCategory.Ducts,
                "cable trays" or "cable tray" => Models.MepCategory.CableTrays,
                "conduits" or "conduit" => Models.MepCategory.CableTrays, // Conduits not available, use CableTrays instead
                "duct accessories" or "duct accessory" => Models.MepCategory.DuctAccessories,
                "duct fittings" or "duct fitting" => Models.MepCategory.DuctAccessories, // DuctFittings not available, use DuctAccessories instead
                _ => null
            };
        }
        
        private void LoadPredefinedMappings()
        {
            try
            {
                var predefinedMappings = _mappingService.GetPredefinedMappings();
                
                // Set predefined mappings
                foreach (var mapping in predefinedMappings)
                {
                    switch (mapping.TransferType)
                    {
                        case TransferType.ReferenceToOpening:
                            if (_referenceSourceComboBox.Items.Contains(mapping.SourceParameter))
                                _referenceSourceComboBox.SelectedItem = mapping.SourceParameter;
                            if (_referenceTargetComboBox.Items.Contains(mapping.TargetParameter))
                                _referenceTargetComboBox.SelectedItem = mapping.TargetParameter;
                            break;
                        case TransferType.HostToOpening:
                            if (_hostSourceComboBox.Items.Contains(mapping.SourceParameter))
                                _hostSourceComboBox.SelectedItem = mapping.SourceParameter;
                            if (_hostTargetComboBox.Items.Contains(mapping.TargetParameter))
                                _hostTargetComboBox.SelectedItem = mapping.TargetParameter;
                            break;
                        case TransferType.LevelToOpening:
                            if (_levelSourceComboBox.Items.Contains(mapping.SourceParameter))
                                _levelSourceComboBox.SelectedItem = mapping.SourceParameter;
                            if (_levelTargetComboBox.Items.Contains(mapping.TargetParameter))
                                _levelTargetComboBox.SelectedItem = mapping.TargetParameter;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading predefined mappings: {ex.Message}");
            }
        }
        
        private void AddRenamingButton_Click(object sender, EventArgs e)
        {
            // TODO: Implement add renaming condition dialog
            MessageBox.Show("Add renaming condition functionality will be implemented.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        private void RemoveRenamingButton_Click(object sender, EventArgs e)
        {
            if (_renamingDataGridView.SelectedRows.Count > 0)
            {
                _renamingDataGridView.Rows.RemoveAt(_renamingDataGridView.SelectedRows[0].Index);
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
                        RefreshRenamingDataGridView();
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
                        var conditions = GetRenamingConditionsFromDataGridView();
                        var success = _renamingService.ExportToCsv(conditions, saveFileDialog.FileName);
                        
                        if (success)
                        {
                            MessageBox.Show("Renaming conditions exported successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                RefreshRenamingDataGridView();
                MessageBox.Show("Predefined renaming conditions loaded.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading predefined conditions: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void ManageAbbreviationsButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var abbreviationDialog = new ServiceTypeAbbreviationDialog())
                {
                    abbreviationDialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening abbreviation dialog: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void ConfigureOpeningParametersButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (var configDialog = new OpeningParameterConfigurationDialog())
                {
                    if (configDialog.ShowDialog() == DialogResult.OK)
                    {
                        var config = configDialog.GetConfiguration();
                        // TODO: Save configuration or apply to service
                        MessageBox.Show("Opening parameter configuration saved.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening parameter configuration dialog: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void OkButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Build configuration from UI
                BuildConfigurationFromUI();
                
                // Validate configuration
                var validationErrors = ValidateConfiguration();
                if (validationErrors.Count > 0)
                {
                    var errorMessage = "Configuration validation errors:\n" + string.Join("\n", validationErrors);
                    MessageBox.Show(errorMessage, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void PreviewButton_Click(object sender, EventArgs e)
        {
            // TODO: Implement preview functionality
            MessageBox.Show("Preview functionality will be implemented.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        private void HelpButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Parameter Transfer Help:\n\n" +
                           "1. Reference Element to Openings: Transfer parameters from MEP elements to openings\n" +
                           "2. Host to Opening: Transfer parameters from walls/floors/ceilings to openings\n" +
                           "3. Level to Openings: Transfer parameters from levels to openings\n" +
                           "4. Model Information: Transfer model name to openings\n" +
                           "5. Parameter Value Renaming: Rename parameter values during transfer\n\n" +
                           "Use the Renaming Conditions to map values like 'Plumbing' to 'P'.",
                           "Help", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        
        private void RefreshRenamingDataGridView()
        {
            _renamingDataGridView.Rows.Clear();
            
            var conditions = _renamingService.GetAllRenamingConditions();
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
        
        private List<RenamingCondition> GetRenamingConditionsFromDataGridView()
        {
            var conditions = new List<RenamingCondition>();
            
            foreach (DataGridViewRow row in _renamingDataGridView.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    var condition = new RenamingCondition(
                        row.Cells[0].Value.ToString(),
                        row.Cells[1].Value.ToString(),
                        row.Cells[2].Value?.ToString() ?? string.Empty
                    );
                    condition.IsEnabled = Convert.ToBoolean(row.Cells[3].Value ?? true);
                    conditions.Add(condition);
                }
            }
            
            return conditions;
        }
        
        private void BuildConfigurationFromUI()
        {
            _configuration.Mappings.Clear();
            
            // Reference Element to Opening mapping
            if (_referenceEnabledCheckBox.Checked && 
                _referenceSourceComboBox.SelectedItem != null && 
                _referenceTargetComboBox.SelectedItem != null)
            {
                var mapping = new ParameterMapping(
                    _referenceSourceComboBox.SelectedItem.ToString(),
                    _referenceTargetComboBox.SelectedItem.ToString(),
                    TransferType.ReferenceToOpening
                );
                _configuration.Mappings.Add(mapping);
            }
            
            // Host to Opening mapping
            if (_hostEnabledCheckBox.Checked && 
                _hostSourceComboBox.SelectedItem != null && 
                _hostTargetComboBox.SelectedItem != null)
            {
                var mapping = new ParameterMapping(
                    _hostSourceComboBox.SelectedItem.ToString(),
                    _hostTargetComboBox.SelectedItem.ToString(),
                    TransferType.HostToOpening
                );
                mapping.Separator = _hostSeparatorTextBox.Text;
                _configuration.Mappings.Add(mapping);
            }
            
            // Level to Opening mapping
            if (_levelEnabledCheckBox.Checked && 
                _levelSourceComboBox.SelectedItem != null && 
                _levelTargetComboBox.SelectedItem != null)
            {
                var mapping = new ParameterMapping(
                    _levelSourceComboBox.SelectedItem.ToString(),
                    _levelTargetComboBox.SelectedItem.ToString(),
                    TransferType.LevelToOpening
                );
                _configuration.Mappings.Add(mapping);
            }
            
            // Model name transfer
            _configuration.TransferModelNames = _modelEnabledCheckBox.Checked;
            if (_modelTargetComboBox.SelectedItem != null)
            {
                _configuration.ModelNameParameter = _modelTargetComboBox.SelectedItem.ToString();
            }
            
            // Renaming conditions
            _configuration.RenamingConditions = GetRenamingConditionsFromDataGridView();
        }
        
        private List<string> ValidateConfiguration()
        {
            var errors = new List<string>();
            
            // Check if at least one mapping is enabled
            if (_configuration.Mappings.Count == 0 && !_configuration.TransferModelNames)
            {
                errors.Add("At least one parameter transfer must be enabled.");
            }
            
            // Validate renaming conditions
            var renamingErrors = _renamingService.ValidateRenamingConditions(_configuration.RenamingConditions);
            errors.AddRange(renamingErrors);
            
            return errors;
        }
        
        public ParameterTransferConfiguration GetConfiguration()
        {
            return _configuration;
        }
    }
}
