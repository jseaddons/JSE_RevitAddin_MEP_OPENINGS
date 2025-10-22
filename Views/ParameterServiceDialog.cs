using System;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Standalone Parameter Service Dialog - provides parameter transfer functionality
    /// </summary>
    public partial class ParameterServiceDialog : WinForms.Form
    {
        private readonly Document? _document;
        private readonly UIDocument? _uiDocument;
        
        // Parameter service UI components
        private WinForms.Panel _parameterFilterPanel = null!;
        private List<WinForms.Panel> _parameterRows = new List<WinForms.Panel>();
        private WinForms.Button _addParameterButton = null!;
        private WinForms.Button _transferAllHeaderBtn = null!;
        
        // Master parameter tabs (Reference Elements vs Host Elements)
        private WinForms.TabControl _masterParameterTabs = null!;
        
        // Reference Elements sub-tabs (MEP categories)
        private WinForms.TabControl _referenceParameterTabs = null!;
        
        // Host Elements sub-tabs (Host categories)
        private WinForms.TabControl _hostParameterTabs = null!;
        
        // Parameter marking configuration
        private WinForms.TextBox _sleeveParameterPrefixTextBox = null!;
        private WinForms.TextBox _projectPrefixTextBox = null!;
        private WinForms.TextBox _disciplinePrefixTextBox = null!;
        private WinForms.TextBox _ductPrefixTextBox = null!;
        private WinForms.TextBox _pipePrefixTextBox = null!;
        private WinForms.TextBox _cableTrayPrefixTextBox = null!;
        private WinForms.TextBox _damperPrefixTextBox = null!;
        private WinForms.CheckBox _remarkAllCheckBox = null!;

        private Services.HostParameterService _hostParameterService;
        
        // Selection tracking for parameter filtering
        private List<string> _selectedReferenceFiles = new List<string>();
        private List<string> _selectedCategories = new List<string>();
        private List<string> _selectedHostFiles = new List<string>();

        public ParameterServiceDialog(Document? document = null, UIDocument? uiDocument = null)
        {
            _document = document;
            _uiDocument = uiDocument;
            
            InitializeComponent();
            
            // Load saved prefix settings
            LoadPrefixSettings();
            
            // Automatically populate parameters from linked files after initialization
            PopulateParametersFromLinkedFiles();
        }

        private void InitializeComponent()
        {
            this.Text = "Parameter Service";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = WinForms.FormStartPosition.CenterParent;
            this.FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            this.MinimumSize = new System.Drawing.Size(600, 400);

            // Create main layout
            CreateMainLayout();
        }

        private void CreateMainLayout()
        {
            // Create parameter filter panel
            _parameterFilterPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Padding = new WinForms.Padding(10)
            };
            this.Controls.Add(_parameterFilterPanel);

            // Create title
            var title = new WinForms.Label
            {
                Text = "Parameter Transfer Service",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(300, 25),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51)
            };
            _parameterFilterPanel.Controls.Add(title);

            // Add Transfer All button
            _transferAllHeaderBtn = new WinForms.Button
            {
                Text = "Transfer All →",
                Size = new System.Drawing.Size(110, 30),
                Location = new System.Drawing.Point(_parameterFilterPanel.Width - 120, 10),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right,
                BackColor = System.Drawing.Color.FromArgb(100, 150, 200),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Enabled = true // Always enabled in standalone dialog
            };
            _transferAllHeaderBtn.Click += OnTransferAllClick;
            _parameterFilterPanel.Controls.Add(_transferAllHeaderBtn);

            // Add Parameter button
            _addParameterButton = new WinForms.Button
            {
                Text = "+ Add Parameter",
                Size = new System.Drawing.Size(120, 30),
                Location = new System.Drawing.Point(_parameterFilterPanel.Width - 250, 10),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right,
                BackColor = System.Drawing.Color.FromArgb(230, 255, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Enabled = true // Always enabled in standalone dialog
            };
            _addParameterButton.Click += OnAddParameterClick;
            _parameterFilterPanel.Controls.Add(_addParameterButton);

            // Create master tabs
            _masterParameterTabs = new WinForms.TabControl
            {
                Location = new System.Drawing.Point(10, 50),
                Size = new System.Drawing.Size(_parameterFilterPanel.Width - 20, _parameterFilterPanel.Height - 170), // Adjusted for larger marking panel
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right | WinForms.AnchorStyles.Bottom
            };
            _parameterFilterPanel.Controls.Add(_masterParameterTabs);

            // Create tabs
            CreateReferenceElementsMasterTab();
            CreateHostElementsMasterTab();

            // Create parameter marking section
            CreateParameterMarkingSection();

            // Handle resize
            _parameterFilterPanel.Resize += (_, __) =>
            {
                _transferAllHeaderBtn.Location = new System.Drawing.Point(_parameterFilterPanel.Width - 120, 10);
                _addParameterButton.Location = new System.Drawing.Point(_parameterFilterPanel.Width - 250, 10);
                _masterParameterTabs.Size = new System.Drawing.Size(_parameterFilterPanel.Width - 20, _parameterFilterPanel.Height - 170);
            };
        }

        private void CreateReferenceElementsMasterTab()
        {
            var referenceTabPage = new WinForms.TabPage("Reference Elements");
            
            // Create sub-tabs for MEP categories
            _referenceParameterTabs = new WinForms.TabControl
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.White
            };
            referenceTabPage.Controls.Add(_referenceParameterTabs);

            // Create MEP service tabs
            CreateServiceTab("Ducts", "DUCTS");
            CreateServiceTab("Damper", "DUCT_ACCESSORIES");
            CreateServiceTab("Cable Trays", "CABLE_TRAYS");
            CreateServiceTab("Pipes", "PIPES");

            _masterParameterTabs.TabPages.Add(referenceTabPage);
        }

        private void CreateHostElementsMasterTab()
        {
            var hostTabPage = new WinForms.TabPage("Host Elements");
            
            // Create sub-tabs for Host categories
            _hostParameterTabs = new WinForms.TabControl
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.White
            };
            hostTabPage.Controls.Add(_hostParameterTabs);

            // Create Host element tabs using a shared service instance
            if (_hostParameterService == null)
            {
                _hostParameterService = new Services.HostParameterService();
            }
            _hostParameterService.CreateHostTabs(_hostParameterTabs);

            _masterParameterTabs.TabPages.Add(hostTabPage);
        }

        private void CreateServiceTab(string tabName, string serviceCode)
        {
            var tabPage = new WinForms.TabPage(tabName);
            
            var servicePanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.White,
                Padding = new WinForms.Padding(3)
            };

            // Add default essential parameter rows with proper mappings
            AddDefaultParameterRows(servicePanel, serviceCode);

            tabPage.Controls.Add(servicePanel);
            _referenceParameterTabs.TabPages.Add(tabPage);
        }

        /// <summary>
        /// Add default parameter rows with essential parameters and proper MEP-to-Opening mappings
        /// </summary>
        private void AddDefaultParameterRows(WinForms.Panel servicePanel, string serviceCode)
        {
            // Define default parameter mappings based on service type
            var defaultMappings = GetDefaultParameterMappings(serviceCode);
            
            foreach (var mapping in defaultMappings)
            {
                AddServiceParameterRow(servicePanel, mapping.MepParameter, mapping.OpeningParameter);
            }
        }

        /// <summary>
        /// Get default parameter mappings for each service type - ONLY 4 ESSENTIAL PARAMETERS
        /// </summary>
        private List<(string MepParameter, string OpeningParameter)> GetDefaultParameterMappings(string serviceCode)
        {
            var mappings = new List<(string MepParameter, string OpeningParameter)>();
            
            // ONLY 4 ESSENTIAL PARAMETERS FOR ALL SERVICE TYPES (Option A)
            var systemTypeParam = serviceCode.ToUpper() == "CABLE_TRAYS" ? "Service Type" : "System Type";
            
            mappings.AddRange(new[]
            {
                ("Size", "MEP Size"),
                (systemTypeParam, "MEP System Type"),
                ("Reference Level", "Level"),
                ("MEP System Name", "System Name")
            });
            
            // No additional service-specific parameters - only the 4 essential ones
            
            return mappings;
        }

        private void AddServiceParameterRow(WinForms.Panel servicePanel, string parameterName, string value)
        {
            int rowHeight = 24;
            int top = 25 + (servicePanel.Controls.OfType<WinForms.Panel>().Count() * (rowHeight + 3));

            var row = new WinForms.Panel
            {
                Location = new System.Drawing.Point(8, top),
                Size = new System.Drawing.Size(servicePanel.Width - 16, rowHeight + 10),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            servicePanel.Controls.Add(row);

            // MEP Parameter ComboBox (left side) - Use optimized loading
            var nameCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(0, 2),
                Size = new System.Drawing.Size(120, 20),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                Tag = "mep"
            };
            
            // Get essential MEP parameters only for better performance
            var parameterService = new Services.ParameterExtractionService();
            var mepParameters = GetEssentialMepParameters();
            
            nameCombo.Items.AddRange(mepParameters.ToArray());
            nameCombo.SelectedItem = parameterName;

            row.Controls.Add(nameCombo);

            // Opening Parameter ComboBox (right side) - Use optimized loading
            var openingParamCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(130, 2),
                Size = new System.Drawing.Size(row.Width - 130 - 60 - 80, 30), // Reduced width to make room for text input
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                Tag = "opening"
            };

            // Get essential opening parameters only
            var openingParameters = GetEssentialOpeningParameters();
            openingParamCombo.Items.AddRange(openingParameters.ToArray());
            openingParamCombo.SelectedItem = value;

            row.Controls.Add(openingParamCombo);

            // Add text input for custom parameter names
            var customParamTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(row.Width - 140, 2),
                Size = new System.Drawing.Size(80, 20),
                Text = "", // Start with empty text
                Tag = "custom_opening_param",
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            customParamTextBox.TextChanged += (_, __) => {
                // When user types in custom parameter, clear the dropdown selection
                if (!string.IsNullOrEmpty(customParamTextBox.Text))
                {
                    openingParamCombo.SelectedItem = null;
                }
            };
            row.Controls.Add(customParamTextBox);

            // Add remove button
            var removeBtn = new WinForms.Button
            {
                Text = "×",
                Location = new System.Drawing.Point(row.Width - 25, 1),
                Size = new System.Drawing.Size(20, 20),
                BackColor = System.Drawing.Color.FromArgb(255, 230, 230),
                FlatStyle = WinForms.FlatStyle.Flat,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Right
            };
            removeBtn.Click += (_, __) => {
                servicePanel.Controls.Remove(row);
                RepositionServiceParameterRows(servicePanel);
            };
            row.Controls.Add(removeBtn);
        }

        private void RepositionServiceParameterRows(WinForms.Panel servicePanel)
        {
            var rows = servicePanel.Controls.OfType<WinForms.Panel>().ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].Location = new System.Drawing.Point(8, 25 + (i * 27));
            }
        }

        private void CreateParameterMarkingSection()
        {
            // Create marking panel at bottom with more height to accommodate all prefixes
            var markingPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, _parameterFilterPanel.Height - 150),
                Size = new System.Drawing.Size(_parameterFilterPanel.Width - 20, 140),
                BackColor = System.Drawing.Color.FromArgb(240, 248, 255),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _parameterFilterPanel.Controls.Add(markingPanel);

            // Project Prefix
            var projectPrefixLabel = new WinForms.Label
            {
                Text = "Project Prefix:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 8),
                Size = new System.Drawing.Size(80, 18)
            };
            markingPanel.Controls.Add(projectPrefixLabel);

            _projectPrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(95, 6),
                Size = new System.Drawing.Size(70, 20),
                Text = "SLEEVE_",
                Tag = "project_prefix"
            };
            _projectPrefixTextBox.TextChanged += OnPrefixChanged;
            markingPanel.Controls.Add(_projectPrefixTextBox);

            // Duct Prefix
            var ductPrefixLabel = new WinForms.Label
            {
                Text = "Duct Prefix:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(80, 18)
            };
            markingPanel.Controls.Add(ductPrefixLabel);

            _ductPrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(95, 33),
                Size = new System.Drawing.Size(50, 20),
                Text = "D",
                Tag = "duct_prefix"
            };
            _ductPrefixTextBox.TextChanged += OnPrefixChanged;
            markingPanel.Controls.Add(_ductPrefixTextBox);

            // Pipe Prefix
            var pipePrefixLabel = new WinForms.Label
            {
                Text = "Pipe Prefix:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 62),
                Size = new System.Drawing.Size(80, 18)
            };
            markingPanel.Controls.Add(pipePrefixLabel);

            _pipePrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(95, 60),
                Size = new System.Drawing.Size(50, 20),
                Text = "P",
                Tag = "pipe_prefix"
            };
            _pipePrefixTextBox.TextChanged += OnPrefixChanged;
            markingPanel.Controls.Add(_pipePrefixTextBox);

            // Cable Tray Prefix
            var cableTrayPrefixLabel = new WinForms.Label
            {
                Text = "Cable Tray Prefix:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 89),
                Size = new System.Drawing.Size(80, 18)
            };
            markingPanel.Controls.Add(cableTrayPrefixLabel);

            _cableTrayPrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(95, 87),
                Size = new System.Drawing.Size(50, 20),
                Text = "E",
                Tag = "cabletray_prefix"
            };
            _cableTrayPrefixTextBox.TextChanged += OnPrefixChanged;
            markingPanel.Controls.Add(_cableTrayPrefixTextBox);

            // Damper Prefix (for duct accessories/dampers)
            var damperPrefixLabel = new WinForms.Label
            {
                Text = "Damper:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                Location = new System.Drawing.Point(10, 116),
                Size = new System.Drawing.Size(60, 18)
            };
            markingPanel.Controls.Add(damperPrefixLabel);

            _damperPrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(95, 114),
                Size = new System.Drawing.Size(50, 20),
                Text = "DMP",
                Tag = "damper_prefix"
            };
            _damperPrefixTextBox.TextChanged += OnPrefixChanged;
            markingPanel.Controls.Add(_damperPrefixTextBox);

            // Re-mark all checkbox - positioned to the right of the prefix textboxes
            _remarkAllCheckBox = new WinForms.CheckBox
            {
                Location = new System.Drawing.Point(160, 6),
                Size = new System.Drawing.Size(110, 20),
                Text = "Re-mark all",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular),
                Checked = false,
                Tag = "remark_all"
            };
            markingPanel.Controls.Add(_remarkAllCheckBox);

            // Apply Marks button - positioned below the prefix textboxes
            var applyMarksButton = new WinForms.Button
            {
                Text = "Apply Marks",
                Location = new System.Drawing.Point(160, 35),
                Size = new System.Drawing.Size(100, 25),
                BackColor = System.Drawing.Color.FromArgb(100, 150, 200),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            applyMarksButton.Click += OnApplyMarksClick;
            markingPanel.Controls.Add(applyMarksButton);

            // Legacy hidden textbox
            _sleeveParameterPrefixTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(0, 0),
                Size = new System.Drawing.Size(1, 1),
                Visible = false,
                Text = "SLEEVE_",
                Tag = "sleeve_parameter_prefix"
            };
            markingPanel.Controls.Add(_sleeveParameterPrefixTextBox);

            // Wire up tab selection change to update discipline prefix
            _referenceParameterTabs.SelectedIndexChanged += OnMepTabSelectionChanged;
            _masterParameterTabs.SelectedIndexChanged += OnMasterTabSelectionChanged;
            
            // Update Transfer All button status based on sleeve state
            UpdateTransferAllButtonStatus();
        }

        /// <summary>
        /// Check sleeve status in the model and update Transfer All button accordingly
        /// </summary>
        private void UpdateTransferAllButtonStatus()
        {
            try
            {
                var openingIds = GetAllOpeningInstanceIds(_document);
                var sleevesWithPrefixes = 0;
                
                foreach (var openingId in openingIds)
                {
                    var opening = _document.GetElement(openingId);
                    if (opening != null)
                    {
                        var markParam = opening.LookupParameter("Mark");
                        if (markParam != null && markParam.StorageType == StorageType.String)
                        {
                            var markValue = markParam.AsString();
                            if (!string.IsNullOrEmpty(markValue))
                            {
                                sleevesWithPrefixes++;
                            }
                        }
                    }
                }
                
                // Always keep Transfer All button enabled and visible
                _transferAllHeaderBtn.Enabled = true;
                _transferAllHeaderBtn.BackColor = System.Drawing.Color.FromArgb(100, 150, 200);
                _transferAllHeaderBtn.ForeColor = System.Drawing.Color.White;
                
                if (sleevesWithPrefixes > 0)
                {
                    _transferAllHeaderBtn.Text = $"Transfer All → ({sleevesWithPrefixes})";
                }
                else
                {
                    _transferAllHeaderBtn.Text = "Transfer All →";
                }
                
                DebugLogger.Info($"[UI] Transfer All button updated: {openingIds.Count} sleeves, {sleevesWithPrefixes} with prefixes");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UI] Error updating Transfer All button status: {ex.Message}");
                _transferAllHeaderBtn.Enabled = true;
                _transferAllHeaderBtn.Text = "Transfer All →";
                _transferAllHeaderBtn.BackColor = System.Drawing.Color.FromArgb(100, 150, 200);
            }
        }

        private string GetCurrentServiceCategory()
        {
            try
            {
                var activeMasterTab = _masterParameterTabs?.SelectedTab;
                if (activeMasterTab == null) return "Unknown";
                
                var subTabs = activeMasterTab.Controls.OfType<WinForms.TabControl>().FirstOrDefault();
                var activeServiceTab = subTabs?.SelectedTab;
                
                if (activeServiceTab != null)
                {
                    var tabName = activeServiceTab.Text.ToLower();
                    if (tabName.Contains("duct")) return "Ducts";
                    if (tabName.Contains("pipe")) return "Pipes";
                    if (tabName.Contains("cable") || tabName.Contains("tray")) return "Cable Trays";
                    if (tabName.Contains("accessory") || tabName.Contains("damper")) return "Duct Accessories";
                }
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        private void OnTransferAllClick(object sender, EventArgs e)
        {
            try
            {
                // Write directly to file for debugging
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Transfer All button clicked\n");
                
                DebugLogger.Info("[TRANSFER_ALL] Transfer All button clicked");
                
                var activeMasterTab = _masterParameterTabs?.SelectedTab;
                if (activeMasterTab == null) 
                {
                    DebugLogger.Warning("[TRANSFER_ALL] No active master tab found");
                    return;
                }
                
                DebugLogger.Info($"[TRANSFER_ALL] Active master tab: {activeMasterTab.Text}");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Active master tab: {activeMasterTab.Text}\n");
                
                var subTabs = activeMasterTab.Controls.OfType<WinForms.TabControl>().FirstOrDefault();
                var activeServiceTab = subTabs?.SelectedTab;
                var servicePanel = activeServiceTab?.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                
                DebugLogger.Info($"[TRANSFER_ALL] Active service tab: {activeServiceTab?.Text}");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Active service tab: {activeServiceTab?.Text}\n");
                
                if (servicePanel != null)
                {
                    DebugLogger.Info("[TRANSFER_ALL] Calling TransferAllMappingsFromPanel");
                    TransferAllMappingsFromPanel(servicePanel);
                }
                else
                {
                    DebugLogger.Warning("[TRANSFER_ALL] No service panel found");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[TRANSFER_ALL] Error in Transfer All click: {ex.Message}");
                WinForms.MessageBox.Show($"Error transferring parameters: {ex.Message}", "Transfer Error",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnAddParameterClick(object sender, EventArgs e)
        {
            try
            {
                var activeMasterTab = _masterParameterTabs?.SelectedTab;
                if (activeMasterTab == null) return;
                
                var subTabs = activeMasterTab.Controls.OfType<WinForms.TabControl>().FirstOrDefault();
                var activeServiceTab = subTabs?.SelectedTab;
                var servicePanel = activeServiceTab?.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                
                if (servicePanel != null)
                {
                    AddServiceParameterRow(servicePanel, "<Select>", "");
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error adding parameter row: {ex.Message}", "Add Parameter Error",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void TransferAllMappingsFromPanel(WinForms.Panel servicePanel)
        {
            try
            {
                DebugLogger.Info("[TRANSFER_ALL] Starting TransferAllMappingsFromPanel");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Starting TransferAllMappingsFromPanel\n");
                
                // Get all opening instances
                var openingIds = GetAllOpeningInstanceIds(_document);
                DebugLogger.Info($"[TRANSFER_ALL] Found {openingIds.Count} opening instances");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Found {openingIds.Count} opening instances\n");
                
                // NO VALIDATION - proceed with transfer regardless
                DebugLogger.Info("[TRANSFER_ALL] Proceeding with transfer - no validation checks");

                // Get parameter mappings from the panel
                var mappings = GetParameterMappingsFromPanel(servicePanel);
                DebugLogger.Info($"[TRANSFER_ALL] Found {mappings.Count} parameter mappings");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Found {mappings.Count} parameter mappings\n");
                
                if (mappings.Count == 0)
                {
                    DebugLogger.Warning("[TRANSFER_ALL] No parameter mappings found");
                    WinForms.MessageBox.Show("No parameter mappings found. Please add parameter mappings first.",
                        "No Mappings", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                    return;
                }

                // Log the mappings
                foreach (var mapping in mappings)
                {
                    DebugLogger.Info($"[TRANSFER_ALL] Mapping: {mapping.SourceParameter} -> {mapping.TargetParameter}");
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                        $"[{DateTime.Now}] [TRANSFER_ALL] Mapping: {mapping.SourceParameter} -> {mapping.TargetParameter}\n");
                }

                // Transfer parameters
                var transferService = new ParameterTransferService();
                var config = new ParameterTransferConfiguration
                {
                    Mappings = mappings,
                    SourceCategoryName = GetCurrentServiceCategory() // Add the current service category
                };

                DebugLogger.Info($"[TRANSFER_ALL] Calling ExecuteTransferConfiguration with category: {config.SourceCategoryName}");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Calling ExecuteTransferConfiguration with category: {config.SourceCategoryName}\n");

                var result = transferService.ExecuteTransferConfiguration(_document, openingIds, config);
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] ExecuteTransferConfiguration completed\n");
                
                DebugLogger.Info($"[TRANSFER_ALL] Transfer result: Success={result.Success}, TransferredCount={result.TransferredCount}, FailedCount={result.FailedCount}");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\transfer_debug.log", 
                    $"[{DateTime.Now}] [TRANSFER_ALL] Transfer result: Success={result.Success}, TransferredCount={result.TransferredCount}, FailedCount={result.FailedCount}\n");
                
                if (result.Success)
                {
                    DebugLogger.Info($"[TRANSFER_ALL] Transfer completed successfully: {result.TransferredCount} parameters transferred");
                    WinForms.MessageBox.Show($"Successfully transferred {result.TransferredCount} parameters.",
                        "Transfer Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
                else
                {
                    DebugLogger.Error($"[TRANSFER_ALL] Transfer failed: {string.Join(", ", result.Errors)}");
                    WinForms.MessageBox.Show($"Transfer failed: {string.Join(", ", result.Errors)}",
                        "Transfer Failed", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[TRANSFER_ALL] Exception in TransferAllMappingsFromPanel: {ex.Message}");
                DebugLogger.Error($"[TRANSFER_ALL] Stack trace: {ex.StackTrace}");
                WinForms.MessageBox.Show($"Error in parameter transfer: {ex.Message}", "Transfer Error",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private List<ElementId> GetAllOpeningInstanceIds(Document doc)
        {
            if (doc == null) return new List<ElementId>();

            var collector = new FilteredElementCollector(doc);
            var openingInstances = collector
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .Select(fi => fi.Id)
                    .ToList();

            return openingInstances;
        }

        private List<ParameterMapping> GetParameterMappingsFromPanel(WinForms.Panel servicePanel)
        {
            var mappings = new List<ParameterMapping>();
            
            foreach (var row in servicePanel.Controls.OfType<WinForms.Panel>())
            {
                var mepCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "mep");
                var openingCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "opening");
                var customTextBox = row.Controls.OfType<WinForms.TextBox>().FirstOrDefault(t => t.Tag?.ToString() == "custom_opening_param");
                
                if (mepCombo?.SelectedItem != null)
                {
                    string targetParameter;
                    
                    // Use custom text input if provided, otherwise use dropdown selection
                    if (!string.IsNullOrEmpty(customTextBox?.Text?.Trim()))
                    {
                        targetParameter = customTextBox.Text.Trim();
                    }
                    else if (openingCombo?.SelectedItem != null)
                    {
                        targetParameter = openingCombo.SelectedItem.ToString();
                    }
                    else
                    {
                        continue; // Skip if no target parameter is selected or entered
                    }
                    
                    mappings.Add(new ParameterMapping
                    {
                        SourceParameter = mepCombo.SelectedItem.ToString(),
                        TargetParameter = targetParameter,
                        TransferType = TransferType.ReferenceToOpening
                    });
                }
            }
            
            return mappings;
        }

        /// <summary>
        /// Get essential MEP parameters only for better performance - FIXED SYSTEM TYPE ISSUE
        /// </summary>
        private List<string> GetEssentialMepParameters()
        {
            var essentialParams = new List<string>();
            
            // Always include these core parameters first - these are guaranteed to be available
            var coreParameters = new[]
            {
                "Size", "System Type", "Service Type", "System Abbreviation", "Height", "Reference Level", 
                "Level", "Diameter", "Width", "Schedule Level", "Reference Level Elevation",
                "MEP Size", "MEP System Type", "MEP System Name", "MEP System Abbreviation"
            };
            
            essentialParams.AddRange(coreParameters);
            
            try
            {
                var parameterService = new Services.ParameterExtractionService();
                
                // DEBUG: Check what parameters are actually returned
                var mepParameters = parameterService.HarvestMepParametersFromDocument(_document);
                System.Diagnostics.Debug.WriteLine($"[DEBUG_MEP] HarvestMepParametersFromDocument returned {mepParameters.Count} parameters: {string.Join(", ", mepParameters.Take(10))}");
                
                // Add ALL parameters from HarvestMepParametersFromDocument (don't filter)
                foreach (var param in mepParameters)
                {
                    if (!essentialParams.Contains(param))
                    {
                        essentialParams.Add(param);
                    }
                }
                
                // Also check current document directly for System Type
                try
                {
                    var collector = new FilteredElementCollector(_document);
                    var ductElements = collector.OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().Take(5);
                    var pipeElements = collector.OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().Take(5);
                    
                    foreach (var element in ductElements.Concat(pipeElements))
                    {
                        foreach (Parameter param in element.Parameters)
                        {
                            if (param?.Definition?.Name != null && !essentialParams.Contains(param.Definition.Name))
                            {
                                essentialParams.Add(param.Definition.Name);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error checking current document for System Type: {ex.Message}");
                }
                
                System.Diagnostics.Debug.WriteLine($"[ESSENTIAL_MEP] Final list has {essentialParams.Count} parameters: {string.Join(", ", essentialParams.Take(15))}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting essential MEP parameters: {ex.Message}");
            }
            
            return essentialParams.OrderBy(p => p).ToList();
        }

        /// <summary>
        /// Get essential opening parameters only for better performance - NO DESTRUCTIVE OPERATIONS
        /// </summary>
        private List<string> GetEssentialOpeningParameters()
        {
            var essentialParams = new List<string>();
            
            // Always include these core opening parameters first - these are guaranteed to be available
            var coreOpeningParameters = new[]
            {
                "MEP Size", "MEP System Name", "MEP System Abbreviation", "MEP System Type",
                "Height", "Reference Level", "Level", "Diameter", "Width", "Schedule Level", 
                "Reference Level Elevation", "Mark", "Comments", "Size", "System Type", 
                "System Abbreviation", "System Name"
            };
            
            essentialParams.AddRange(coreOpeningParameters);
            
            try
            {
                // ONLY use the optimized method - NO full parameter loading to avoid destructive operations
                var parameterService = new Services.ParameterExtractionService();
                var additionalParams = parameterService.GetEssentialOpeningParameters(_document);
                
                // Add additional parameters that aren't already in the core list
                foreach (var param in additionalParams)
                {
                    if (!essentialParams.Contains(param))
                    {
                        essentialParams.Add(param);
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[ESSENTIAL_OPENING] Found {essentialParams.Count} essential opening parameters: {string.Join(", ", essentialParams.Take(10))}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting essential opening parameters: {ex.Message}");
            }
            
            return essentialParams.OrderBy(p => p).ToList();
        }

        private List<string> GetCurrentOpeningParameters()
        {
            try
            {
                var openingParameters = new List<string>();
                
                if (_document == null) return openingParameters;
                
                // Get opening families
                var openingFamilies = GetOpeningFamilies(_document);
                if (openingFamilies.Count == 0) return openingParameters;
                
                // Get ALL parameters from ALL opening families, including shared parameters
                foreach (var family in openingFamilies)
                {
                    var parameters = family.Parameters;
                    foreach (Parameter param in parameters)
                    {
                        if (!openingParameters.Contains(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                
                // Also get parameters from opening family symbols (types) for additional parameters
                var openingSymbols = new List<FamilySymbol>();
                var collector = new FilteredElementCollector(_document);
                var familySymbols = collector.OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
                
                foreach (var symbol in familySymbols)
                {
                    if (symbol.Family?.Name?.Contains("Opening") == true && !openingSymbols.Contains(symbol))
                    {
                        openingSymbols.Add(symbol);
                        
                        // Get parameters from the symbol as well
                        var symbolParameters = symbol.Parameters;
                        foreach (Parameter param in symbolParameters)
                        {
                            if (!openingParameters.Contains(param.Definition.Name))
                            {
                                openingParameters.Add(param.Definition.Name);
                            }
                        }
                    }
                }
                
                // Also get parameters from actual opening instances in the document
                var openingInstances = collector.OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .Take(10); // Limit to first 10 instances to avoid performance issues
                
                foreach (var instance in openingInstances)
                {
                    var instanceParameters = instance.Parameters;
                    foreach (Parameter param in instanceParameters)
                    {
                        if (!openingParameters.Contains(param.Definition.Name))
                        {
                            openingParameters.Add(param.Definition.Name);
                        }
                    }
                }
                
                // Sort parameters for better user experience
                openingParameters.Sort();
                
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SERVICE] Found {openingParameters.Count} opening parameters including shared parameters");
                
                return openingParameters;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening parameters: {ex.Message}");
                return new List<string>();
            }
        }

        private List<Family> GetOpeningFamilies(Document document)
        {
            var families = new List<Family>();
            
            try
            {
                var collector = new FilteredElementCollector(document);
                var familySymbols = collector.OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>();
                
                foreach (var symbol in familySymbols)
                {
                    if (symbol.Family?.Name?.Contains("Opening") == true && !families.Contains(symbol.Family))
                    {
                        families.Add(symbol.Family);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening families: {ex.Message}");
        }

            return families;
        }

        private void OnMepTabSelectionChanged(object sender, EventArgs e)
        {
            UpdateDisciplinePrefixBasedOnActiveTab();
            UpdateTransferAllButtonStatus();
        }

        private void OnMasterTabSelectionChanged(object sender, EventArgs e)
        {
            UpdateDisciplinePrefixBasedOnActiveTab();
            UpdateTransferAllButtonStatus();
        }

        private void UpdateDisciplinePrefixBasedOnActiveTab()
        {
            try
            {
                // Only update if we're on the Reference Elements tab
                if (_masterParameterTabs.SelectedTab?.Text != "Reference Elements")
                {
                    return;
                }

                var activeMepTab = _referenceParameterTabs.SelectedTab;
                if (activeMepTab == null)
                {
                    return;
                }

                // Only auto-update if the textbox is empty or contains default values
                var currentText = _disciplinePrefixTextBox.Text?.Trim();
                if (string.IsNullOrEmpty(currentText) || currentText == "D" || currentText == "P" || currentText == "C")
                {
                    // Map MEP service categories to discipline prefixes
                    var disciplinePrefix = GetDisciplinePrefixForMepService(activeMepTab.Text);
                    if (!string.IsNullOrEmpty(disciplinePrefix))
                    {
                        _disciplinePrefixTextBox.Text = disciplinePrefix;
                    }
                }
                // If user has entered custom text, preserve it
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating discipline prefix: {ex.Message}");
            }
        }

        private string GetDisciplinePrefixForMepService(string mepServiceName)
        {
            // Map MEP service names to discipline prefixes
            switch (mepServiceName.ToUpper())
            {
                case "DUCTS":
                    return "D"; // Duct
                case "DUCT ACCESSORIES":
                    return "D"; // Duct
                case "PIPES":
                    return "P"; // Pipe
                case "CABLE TRAYS":
                    return "C"; // Cable
                default:
                    return "D"; // Default to Duct
            }
        }

        private void OnApplyMarksClick(object sender, EventArgs e)
        {
            try
            {
                // Get mark prefixes from UI
                var projectPrefix = _projectPrefixTextBox.Text.Trim();
                var ductPrefix = _ductPrefixTextBox.Text.Trim();
                var pipePrefix = _pipePrefixTextBox.Text.Trim();
                var cableTrayPrefix = _cableTrayPrefixTextBox.Text.Trim();
                var damperPrefix = _damperPrefixTextBox.Text.Trim();
                var remarkAll = _remarkAllCheckBox.Checked;

                if (string.IsNullOrEmpty(projectPrefix))
                {
                    WinForms.MessageBox.Show("Please enter a Project Prefix.", "Missing Project Prefix",
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(ductPrefix) && string.IsNullOrEmpty(pipePrefix) && 
                    string.IsNullOrEmpty(cableTrayPrefix) && string.IsNullOrEmpty(damperPrefix))
                {
                    WinForms.MessageBox.Show("Please enter at least one Discipline Prefix.", "Missing Discipline Prefix",
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                    return;
                }

                // Execute mark MEP command with type-specific prefixes
                var markPrefixes = new MarkPrefixSettings
                {
                    ProjectPrefix = projectPrefix,
                    DuctPrefix = ductPrefix,
                    PipePrefix = pipePrefix,
                    CableTrayPrefix = cableTrayPrefix,
                    DamperPrefix = damperPrefix,
                    RemarkAll = remarkAll
                };

                var markMepCommand = new Commands.MarkParameterCommand("ALL", projectPrefix, "", remarkAll, markPrefixes);
                
                // Get UIApplication from UIDocument
                if (_uiDocument?.Application == null)
                {
                    WinForms.MessageBox.Show("Unable to access Revit application.", "Application Error",
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }
                
                try
                {
                    markMepCommand.Execute(_uiDocument.Application);
                    WinForms.MessageBox.Show("Successfully applied marks to MEP elements.",
                        "Mark Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
                catch (Exception markEx)
                {
                    WinForms.MessageBox.Show($"Mark operation failed: {markEx.Message}",
                        "Mark Failed", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error applying marks: {ex.Message}", "Mark Error",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Save prefix values to user settings for persistence
        /// </summary>
        private void SavePrefixSettings()
        {
            try
            {
                var settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", "prefix_settings.xml");
                
                var doc = new System.Xml.XmlDocument();
                var root = doc.CreateElement("PrefixSettings");
                doc.AppendChild(root);
                
                var projectPrefix = _projectPrefixTextBox?.Text?.Trim() ?? "SLEEVE_";
                var ductPrefix = _ductPrefixTextBox?.Text?.Trim() ?? "D";
                var pipePrefix = _pipePrefixTextBox?.Text?.Trim() ?? "P";
                var cableTrayPrefix = _cableTrayPrefixTextBox?.Text?.Trim() ?? "E";
                var damperPrefix = _damperPrefixTextBox?.Text?.Trim() ?? "DMP";
                var remarkAll = (_remarkAllCheckBox?.Checked ?? false).ToString();
                
                root.SetAttribute("ProjectPrefix", projectPrefix);
                root.SetAttribute("DuctPrefix", ductPrefix);
                root.SetAttribute("PipePrefix", pipePrefix);
                root.SetAttribute("CableTrayPrefix", cableTrayPrefix);
                root.SetAttribute("DamperPrefix", damperPrefix);
                root.SetAttribute("RemarkAll", remarkAll);
                
                doc.Save(settingsPath);
                
                DebugLogger.Info($"[SavePrefixSettings] Saved prefix settings to: {settingsPath}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SavePrefixSettings] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Load prefix values from user settings
        /// </summary>
        private void LoadPrefixSettings()
        {
            try
            {
                var settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", "prefix_settings.xml");
                
                if (!File.Exists(settingsPath)) return;
                
                var doc = new System.Xml.XmlDocument();
                doc.Load(settingsPath);
                
                var root = doc.DocumentElement;
                if (root != null)
                {
                    if (root.HasAttribute("ProjectPrefix"))
                        _projectPrefixTextBox.Text = root.GetAttribute("ProjectPrefix");
                    if (root.HasAttribute("DuctPrefix"))
                        _ductPrefixTextBox.Text = root.GetAttribute("DuctPrefix");
                    if (root.HasAttribute("PipePrefix"))
                        _pipePrefixTextBox.Text = root.GetAttribute("PipePrefix");
                    if (root.HasAttribute("CableTrayPrefix"))
                        _cableTrayPrefixTextBox.Text = root.GetAttribute("CableTrayPrefix");
                    if (root.HasAttribute("DamperPrefix"))
                        _damperPrefixTextBox.Text = root.GetAttribute("DamperPrefix");
                    if (root.HasAttribute("RemarkAll") && bool.TryParse(root.GetAttribute("RemarkAll"), out bool remarkAllValue))
                        _remarkAllCheckBox.Checked = remarkAllValue;
                }
                
                DebugLogger.Info($"[LoadPrefixSettings] Loaded prefix settings from: {settingsPath}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[LoadPrefixSettings] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Event handler for prefix textbox changes - updates the global MarkPrefixService
        /// </summary>
        private void OnPrefixChanged(object sender, EventArgs e)
        {
            try
            {
                // Update the global service whenever prefixes change
                GetCurrentMarkPrefixes();
                
                // Save settings for persistence
                SavePrefixSettings();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OnPrefixChanged] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Get current mark prefixes from UI for use by other dialogs
        /// </summary>
        public MarkPrefixSettings GetCurrentMarkPrefixes()
        {
            try
            {
                var projectPrefix = _projectPrefixTextBox?.Text?.Trim() ?? "SLEEVE_";
                var ductPrefix = _ductPrefixTextBox?.Text?.Trim() ?? "D";
                var pipePrefix = _pipePrefixTextBox?.Text?.Trim() ?? "P";
                var cableTrayPrefix = _cableTrayPrefixTextBox?.Text?.Trim() ?? "E";
                var damperPrefix = _damperPrefixTextBox?.Text?.Trim() ?? "DMP";
                var remarkAll = _remarkAllCheckBox?.Checked ?? false;

                var markPrefixes = new MarkPrefixSettings
                {
                    ProjectPrefix = projectPrefix,
                    DuctPrefix = ductPrefix,
                    PipePrefix = pipePrefix,
                    CableTrayPrefix = cableTrayPrefix,
                    DamperPrefix = damperPrefix,
                    RemarkAll = remarkAll
                };

                // Update the global service so EmergencyMainDialog can access these prefixes
                MarkPrefixService.SetCurrentPrefixes(markPrefixes);

                DebugLogger.Info($"[GetCurrentMarkPrefixes] Project: '{projectPrefix}', Duct: '{ductPrefix}', Pipe: '{pipePrefix}', CableTray: '{cableTrayPrefix}', Damper: '{damperPrefix}', RemarkAll: {remarkAll}");
                return markPrefixes;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GetCurrentMarkPrefixes] Error: {ex.Message}");
                return new MarkPrefixSettings(); // Return defaults on error
            }
        }

        /// <summary>
        /// Automatically populate parameters from linked files for both Reference and Host elements - OPTIMIZED
        /// </summary>
        private void PopulateParametersFromLinkedFiles()
        {
            try
            {
                if (_document == null) return;

                // Get all linked files
                var linkedFileService = new Services.LinkedFileService();
                var linkedFiles = linkedFileService.GetLinkedFiles(_document);
                
                if (linkedFiles.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[PARAMETER_SERVICE] No linked files found for parameter population");
                    return;
                }

                // Get MEP categories and host files from linked files
                var mepCategories = new List<string> { "DUCTS", "DUCT_ACCESSORIES", "CABLE_TRAYS", "PIPES" };
                var hostFiles = linkedFiles.Where(lf => lf.FileType == Services.LinkedFileType.Structural || lf.FileType == Services.LinkedFileType.Architectural).Select(lf => lf.FileName).ToList();
                var referenceFiles = linkedFiles.Where(lf => lf.FileType == Services.LinkedFileType.Mechanical || lf.FileType == Services.LinkedFileType.Electrical || lf.FileType == Services.LinkedFileType.Plumbing || lf.FileType == Services.LinkedFileType.FireProtection).Select(lf => lf.FileName).ToList();

                // Populate host parameters if we have host files - but limit to first 3 files for performance
                if (hostFiles.Count > 0 && _hostParameterTabs != null)
                {
                    var limitedHostFiles = hostFiles.Take(3).ToList();
                    _hostParameterService.PopulateHostParameters(_hostParameterTabs, limitedHostFiles, _document);
                    System.Diagnostics.Debug.WriteLine($"[PARAMETER_SERVICE] Populated host parameters from {limitedHostFiles.Count} host files");
                }

                // Skip reference parameter refresh to avoid performance issues - default rows are already populated
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SERVICE] Skipping reference parameter refresh - using default essential parameters only");

                System.Diagnostics.Debug.WriteLine("[PARAMETER_SERVICE] Parameter population from linked files completed - OPTIMIZED");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SERVICE] Error populating parameters from linked files: {ex.Message}");
            }
        }

        /// <summary>
        /// Refresh existing reference parameter rows with updated parameters from linked files
        /// </summary>
        private void RefreshReferenceParameterRows()
        {
            try
            {
                foreach (WinForms.TabPage tabPage in _referenceParameterTabs.TabPages)
                {
                    var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                    if (servicePanel != null)
                    {
                        // Update each parameter row's MEP parameter dropdown
                        foreach (var row in servicePanel.Controls.OfType<WinForms.Panel>())
                        {
                            var mepCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "mep");
                            if (mepCombo != null)
                            {
                                var currentSelection = mepCombo.SelectedItem?.ToString();
                                
                                // Clear and repopulate with essential parameters only
                                mepCombo.Items.Clear();
                                var essentialMepParameters = GetEssentialMepParameters();
                                mepCombo.Items.AddRange(essentialMepParameters.ToArray());
                                
                                // Restore selection if it still exists
                                if (!string.IsNullOrEmpty(currentSelection) && essentialMepParameters.Contains(currentSelection))
                                {
                                    mepCombo.SelectedItem = currentSelection;
                                }
                                else if (essentialMepParameters.Count > 0)
                                {
                                    mepCombo.SelectedIndex = 0;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PARAMETER_SERVICE] Error refreshing reference parameter rows: {ex.Message}");
            }
        }
    }
}