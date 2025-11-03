using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Commands;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// ✅ NEW UI V2: Two-panel layout (Left: Prefixes, Right: Parameter Mapping)
    /// Skeleton version - functionality to be added later
    /// </summary>
    public partial class ParameterServiceDialogV2 : WinForms.Form
    {
        private readonly Document? _document;
        private readonly UIDocument? _uiDocument;

        // Two-panel containers
        private WinForms.Panel _leftPrefixPanel = null!;
        private WinForms.Panel _rightMappingPanel = null!;

        // Left Panel: Prefix Configuration
        private WinForms.TextBox _projectPrefixTextBox = null!;
        private WinForms.TextBox _ductPrefixTextBox = null!;
        private WinForms.TextBox _pipePrefixTextBox = null!;
        private WinForms.TextBox _cableTrayPrefixTextBox = null!;
        private WinForms.TextBox _damperPrefixTextBox = null!;
        private WinForms.Panel _systemTypeOverridesPanel = null!;
        private List<WinForms.Panel> _systemTypeRows = new List<WinForms.Panel>();
        
        // Number format and remark checkboxes
        private WinForms.ComboBox _numberFormatCombo = null!;
        private WinForms.CheckBox _remarkProjectCheckBox = null!;
        private WinForms.CheckBox _remarkDuctCheckBox = null!;
        private WinForms.CheckBox _remarkPipeCheckBox = null!;
        private WinForms.CheckBox _remarkCableTrayCheckBox = null!;
        private WinForms.CheckBox _remarkDamperCheckBox = null!;

        // Right Panel: Parameter Mapping
        private WinForms.TabControl _masterParameterTabs = null!;
        private WinForms.TabControl _referenceParameterTabs = null!;
        private WinForms.TabControl _hostParameterTabs = null!;
        private WinForms.Button _addParameterButton = null!;

        // Top bar
        private WinForms.Button _applyMarksButton = null!;
        private WinForms.Button _remarkSelectedButton = null!;
        private WinForms.Button _transferParametersButton = null!;
        private WinForms.Button _closeButton = null!;

        public ParameterServiceDialogV2(Document document, UIDocument uiDocument)
        {
            _document = document;
            _uiDocument = uiDocument;

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Parameter Service V2";
            this.Size = new Size(1000, 700);
            this.MinimumSize = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Create UI sections
            CreateTopBar();
            CreateTwoPanelLayout();
            CreatePrefixConfigurationPanel();
            CreateParameterMappingPanel();
        }

        /// <summary>
        /// Top Bar: Transfer Parameters button
        /// </summary>
        private void CreateTopBar()
        {
            var topBar = new WinForms.Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(240, 248, 255)
            };
            this.Controls.Add(topBar);

            // Position 4 buttons aligned to the right with 5px spacing
            int buttonWidth = 140;
            int buttonSpacing = 5;
            int buttonsStartX = topBar.Width - (4 * buttonWidth + 3 * buttonSpacing); // 4 buttons + 3 gaps
            
            // Button 1: Apply Marks
            _applyMarksButton = new WinForms.Button
            {
                Text = "Apply Marks",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX, 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(100, 150, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _applyMarksButton.Click += OnApplyMarksClick;
            topBar.Controls.Add(_applyMarksButton);
            
            // Button 2: Remark Selected
            _remarkSelectedButton = new WinForms.Button
            {
                Text = "Remark Selected",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(150, 100, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _remarkSelectedButton.Click += OnRemarkSelectedClick;
            topBar.Controls.Add(_remarkSelectedButton);
            
            // Button 3: Transfer Parameters
            _transferParametersButton = new WinForms.Button
            {
                Text = "Transfer Parameters",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(100, 150, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _transferParametersButton.Click += OnTransferParametersClick;
            topBar.Controls.Add(_transferParametersButton);
            
            // Button 4: Close
            _closeButton = new WinForms.Button
            {
                Text = "✕ Close",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(200, 100, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _closeButton.Click += (s, e) => this.Close();
            topBar.Controls.Add(_closeButton);
        }

        /// <summary>
        /// Two-panel layout: Left (35% - Prefixes) + Right (65% - Mapping)
        /// </summary>
        private void CreateTwoPanelLayout()
        {
            // Left Panel: 35% width - Prefix Configuration
            _leftPrefixPanel = new WinForms.Panel
            {
                Location = new Point(5, 55),
                Size = new Size((int)(this.Width * 0.35) - 10, this.Height - 100),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left,
                BackColor = Color.FromArgb(250, 250, 255),
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(_leftPrefixPanel);

            // Right Panel: 65% width - Parameter Mapping
            _rightMappingPanel = new WinForms.Panel
            {
                Location = new Point((int)(this.Width * 0.35), 55),
                Size = new Size((int)(this.Width * 0.65) - 10, this.Height - 100),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(_rightMappingPanel);
        }

        /// <summary>
        /// Left Panel: Prefix Configuration (4 rows x 1 column)
        /// </summary>
        private void CreatePrefixConfigurationPanel()
        {
            var yPos = 10;

            // Title
            var title = new WinForms.Label
            {
                Text = "Prefix Configuration",
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(200, 20),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            _leftPrefixPanel.Controls.Add(title);
            yPos += 30;

            // Project Prefix
            var projectPrefixLabel = new WinForms.Label
            {
                Text = "Project Prefix:",
                Location = new Point(10, yPos),
                Size = new Size(100, 20)
            };
            _leftPrefixPanel.Controls.Add(projectPrefixLabel);

            _projectPrefixTextBox = new WinForms.TextBox
            {
                Location = new Point(120, yPos - 2),
                Size = new Size(140, 22),
                Text = "SLEEVE_"
            };
            _leftPrefixPanel.Controls.Add(_projectPrefixTextBox);
            
            // Remark checkbox for Project Prefix
            _remarkProjectCheckBox = new WinForms.CheckBox
            {
                Text = "Remark",
                Location = new Point(270, yPos - 2),
                Size = new Size(80, 22),
                Checked = false
            };
            _leftPrefixPanel.Controls.Add(_remarkProjectCheckBox);
            yPos += 35;

            // Number Format selector
            var numberFormatLabel = new WinForms.Label
            {
                Text = "Number Format:",
                Location = new Point(10, yPos),
                Size = new Size(100, 20)
            };
            _leftPrefixPanel.Controls.Add(numberFormatLabel);
            
            _numberFormatCombo = new WinForms.ComboBox
            {
                Location = new Point(120, yPos - 2),
                Size = new Size(140, 22),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _numberFormatCombo.Items.AddRange(new[] { "00 (01, 02...)", "000 (001, 002...)", "0000 (0001, 0002...)" });
            _numberFormatCombo.SelectedIndex = 1; // Default to "000"
            _leftPrefixPanel.Controls.Add(_numberFormatCombo);
            yPos += 35;

            // Section Title
            var disciplineLabel = new WinForms.Label
            {
                Text = "Discipline Prefixes:",
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(150, 20)
            };
            _leftPrefixPanel.Controls.Add(disciplineLabel);
            yPos += 25;

            // Row 1: Duct
            var ductLabel = new WinForms.Label { Text = "Duct:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(ductLabel);
            _ductPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "M" };
            _leftPrefixPanel.Controls.Add(_ductPrefixTextBox);
            _remarkDuctCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = false };
            _leftPrefixPanel.Controls.Add(_remarkDuctCheckBox);
            yPos += 30;

            // Row 2: Pipe
            var pipeLabel = new WinForms.Label { Text = "Pipe:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(pipeLabel);
            _pipePrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "P" };
            _leftPrefixPanel.Controls.Add(_pipePrefixTextBox);
            _remarkPipeCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = false };
            _leftPrefixPanel.Controls.Add(_remarkPipeCheckBox);
            yPos += 30;

            // Row 3: Cable Tray
            var cableTrayLabel = new WinForms.Label { Text = "Cable Tray:", Location = new Point(10, yPos), Size = new Size(70, 20) };
            _leftPrefixPanel.Controls.Add(cableTrayLabel);
            _cableTrayPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "E" };
            _leftPrefixPanel.Controls.Add(_cableTrayPrefixTextBox);
            _remarkCableTrayCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = false };
            _leftPrefixPanel.Controls.Add(_remarkCableTrayCheckBox);
            yPos += 30;

            // Row 4: Damper
            var damperLabel = new WinForms.Label { Text = "Damper:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(damperLabel);
            _damperPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "DMP" };
            _leftPrefixPanel.Controls.Add(_damperPrefixTextBox);
            _remarkDamperCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = false };
            _leftPrefixPanel.Controls.Add(_remarkDamperCheckBox);
            yPos += 40;

            // ✅ NEW: System Type Overrides section (always visible, not collapsible)
            var systemTypeLabel = new WinForms.Label
            {
                Text = "System Type Overrides:",
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(150, 20)
            };
            _leftPrefixPanel.Controls.Add(systemTypeLabel);
            yPos += 25;

            // Container panel for system type overrides (no scroll, expand freely)
            _systemTypeOverridesPanel = new WinForms.Panel
            {
                Location = new Point(10, yPos),
                Size = new Size(_leftPrefixPanel.Width - 25, 0),
                BackColor = Color.FromArgb(240, 248, 255),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = false // No scrollbars - free flow
            };
            _leftPrefixPanel.Controls.Add(_systemTypeOverridesPanel);

            // Add header row with "+" button
            CreateSystemTypeHeader();

            // Add 4 default rows
            for (int i = 0; i < 4; i++)
            {
                AddSystemTypeRow("<Select>", "");
            }
        }

        /// <summary>
        /// Right Panel: Parameter Mapping (Tabs + Grid)
        /// </summary>
        private void CreateParameterMappingPanel()
        {
            var yPos = 10;

            // Title
            var title = new WinForms.Label
            {
                Text = "Parameter Mapping",
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(200, 20)
            };
            _rightMappingPanel.Controls.Add(title);
            yPos += 30;

            // Master Tabs
            _masterParameterTabs = new WinForms.TabControl
            {
                Location = new Point(5, yPos),
                Size = new Size(_rightMappingPanel.Width - 12, _rightMappingPanel.Height - yPos - 50),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _rightMappingPanel.Controls.Add(_masterParameterTabs);

            // Reference Elements Tab
            var referenceTab = new WinForms.TabPage("Reference Elements");
            _masterParameterTabs.TabPages.Add(referenceTab);

            _referenceParameterTabs = new WinForms.TabControl { Dock = DockStyle.Fill };
            referenceTab.Controls.Add(_referenceParameterTabs);

            // Add sub-tabs with parameter rows
            CreateCategorySubTab("Ducts", _referenceParameterTabs);
            CreateCategorySubTab("Pipes", _referenceParameterTabs);
            CreateCategorySubTab("Cable Trays", _referenceParameterTabs);
            CreateCategorySubTab("Duct Accessories", _referenceParameterTabs);

            // Host Elements Tab
            var hostTab = new WinForms.TabPage("Host Elements");
            _masterParameterTabs.TabPages.Add(hostTab);

            _hostParameterTabs = new WinForms.TabControl { Dock = DockStyle.Fill };
            hostTab.Controls.Add(_hostParameterTabs);

            // Add sub-tabs with parameter rows
            CreateCategorySubTab("Floors", _hostParameterTabs);
            CreateCategorySubTab("Walls", _hostParameterTabs);
            CreateCategorySubTab("Structural Framing", _hostParameterTabs);

            // ✅ REMOVED: Big "+ Add Parameter" button
            // Note: Small "+" buttons for adding rows will be added per-tab later
        }

        /// <summary>
        /// Create a category sub-tab with parameter mapping rows
        /// </summary>
        private void CreateCategorySubTab(string categoryName, WinForms.TabControl parentTabs)
        {
            var tabPage = new WinForms.TabPage(categoryName);
            parentTabs.TabPages.Add(tabPage);

            var servicePanel = new WinForms.Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            // ✅ Add header with "+" button aligned with close buttons
            var headerPanel = new WinForms.Panel
            {
                Location = new Point(0, 0),
                Size = new Size(servicePanel.Width - 15, 25),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            servicePanel.Controls.Add(headerPanel);

            // "+" button at far right, aligned with close buttons (same X position)
            var addButton = new WinForms.Button
            {
                Text = "+",
                Location = new Point(headerPanel.Width - 30, 2), // Aligned with close button X position
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(200, 255, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right // Keep aligned to right edge
            };
            addButton.Click += (s, e) => {
                AddParameterRow(servicePanel, categoryName, "<Select>", "");
            };
            headerPanel.Controls.Add(addButton);

            // Add default parameter rows
            AddDefaultParameterRows(servicePanel, categoryName);
            tabPage.Controls.Add(servicePanel);
        }

        /// <summary>
        /// Add default parameter rows for a category - ONLY 4 ESSENTIAL PARAMETERS on startup
        /// From Linked MEP Files: System Type, System Size, System Abbreviation, Reference Level
        /// </summary>
        private void AddDefaultParameterRows(WinForms.Panel panel, string category)
        {
            // ✅ PERFORMANCE FIX: Only 4 essential parameters from linked MEP files on startup
            // Rest load when user clicks "+" button
            var systemTypeParam = category == "Cable Trays" ? "Service Type" : "System Type";
            
            var mappings = new List<(string Mep, string Opening)>
            {
                ("Size", "MEP Size"),  // ✅ FIX: Use "Size" as default (more commonly used parameter name)
                (systemTypeParam, "MEP System Type"),
                ("System Abbreviation", "MEP System Abbreviation"),
                ("Reference Level", "Level")
            };
            
            // ✅ NOTE: Dropdown will include both "Size" and "System Size" if both exist in linked files
            // Default selection uses "Size" as it's more commonly used

            if (!DeploymentConfiguration.DeploymentMode && _document != null)
            {
                string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                string projectPath = ProjectPathService.GetProjectRoot(_document);
                string filtersPath = ProjectPathService.GetFiltersDirectory(_document);
                System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Adding default parameter rows for category '{category}'\n");
                System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Project path: {projectPath}\n");
                System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Filters directory (XML path): {filtersPath}\n");
            }

            foreach (var mapping in mappings)
            {
                AddParameterRow(panel, category, mapping.Mep, mapping.Opening);
            }
        }

        /// <summary>
        /// Get MEP/Host parameters for a specific category - ONLY ESSENTIAL PARAMETERS on startup
        /// ✅ PERFORMANCE FIX: Returns only 4 essential parameters for MEP categories on startup
        /// Rest load when user clicks "+" button
        /// </summary>
        private string[] GetMepParametersForCategory(string category)
        {
            // Handle host elements (Floors, Walls, Structural Framing) - load from architectural/structural linked files
            if (category == "Floors" || category == "Walls" || category == "Structural Framing")
            {
                return GetHostParametersForCategory(category);
            }
            
            // ✅ PERFORMANCE FIX: Only return 4 essential parameters for MEP categories on startup
            // 1. Size (or "System Size" - include both as some files use different names)
            // 2. System Type (or "Service Type" for Cable Trays)
            // 3. System Abbreviation
            // 4. Reference Level
            var systemTypeParam = category == "Cable Trays" ? "Service Type" : "System Type";
            // ✅ FIX: Include both "Size" and "System Size" as different linked files might use different parameter names
            var essentialParams = new List<string> { "Size", "System Size", systemTypeParam, "System Abbreviation", "Reference Level" };
            // Remove duplicates while preserving order
            essentialParams = essentialParams.Distinct().ToList();
            
            if (_document == null)
            {
                // Fallback to essentials if no document available
                return essentialParams.ToArray();
            }
            
            // ✅ PERFORMANCE FIX: Just return the 4 essential parameters on startup
            // Don't scan all parameters - they'll load when user clicks "+"
            DebugLogger.Info($"[ParameterServiceDialogV2] Returning {essentialParams.Count} essential MEP parameters for category '{category}' (startup mode)");
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Returning essential MEP parameters for '{category}': {string.Join(", ", essentialParams)}\n");
            }
            
            return essentialParams.ToArray();
        }
        
        /// <summary>
        /// Get Host parameters for a specific category - loads from architectural/structural linked files
        /// ✅ PERFORMANCE FIX: Only loads "Fire Rating" on startup, rest load when user clicks "+"
        /// </summary>
        private string[] GetHostParametersForCategory(string category)
        {
            // ✅ PERFORMANCE FIX: Only return "Fire Rating" parameter on startup
            // Rest load when user clicks "+" button
            var essentialParams = new[] { "Fire Rating" };
            
            if (_document == null)
            {
                // Fallback to essentials if no document available
                return essentialParams;
            }
            
            try
            {
                var hostParameters = new HashSet<string>();
                
                // Map category to BuiltInCategory
                BuiltInCategory? targetCategory = category switch
                {
                    "Walls" => BuiltInCategory.OST_Walls,
                    "Floors" => BuiltInCategory.OST_Floors,
                    "Structural Framing" => BuiltInCategory.OST_StructuralFraming,
                    _ => null
                };
                
                if (targetCategory == null)
                {
                    DebugLogger.Warning($"[ParameterServiceDialogV2] Unknown host category: {category}");
                    return essentialParams; // Return essentials instead of wrong defaults
                }
                
                // Get linked files from document
                var linkedFileService = new Services.LinkedFileService();
                var allLinkedFiles = linkedFileService.GetLinkedFiles(_document);
                
                // Get architectural and structural linked files (host element files)
                var hostLinkedFiles = linkedFileService.GetHostElementFiles(allLinkedFiles);
                
                DebugLogger.Info($"[ParameterServiceDialogV2] Found {hostLinkedFiles.Count} host-linked files for category '{category}'");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Found {hostLinkedFiles.Count} host-linked files for category '{category}'\n");
                }
                
                // If no linked files, fall back to current document
                if (hostLinkedFiles.Count == 0)
                {
                    DebugLogger.Info($"[ParameterServiceDialogV2] No host-linked files found, using current document for host parameters");
                    hostLinkedFiles.Add(new Services.LinkedFileInfo 
                    { 
                        FileName = "Current Document",
                        LinkInstance = null 
                    });
                }
                
                // ✅ PERFORMANCE FIX: Only load "Fire Rating" parameter on startup (rest load when user clicks "+")
                // Check if Fire Rating parameter exists in any of the host categories
                foreach (var linkedFile in hostLinkedFiles)
                {
                    Document targetDocument = _document;
                    
                    // If it's a linked file, get the linked document
                    if (linkedFile.FileName != "Current Document" && linkedFile.LinkInstance != null)
                    {
                        var linkDoc = linkedFile.LinkInstance.GetLinkDocument();
                        if (linkDoc != null)
                        {
                            targetDocument = linkDoc;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                                System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Loading host parameters from linked file: {linkedFile.FileName}\n");
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                    
                    // Check if Fire Rating parameter exists in the target category
                    var collector = new FilteredElementCollector(targetDocument)
                        .OfCategory(targetCategory.Value)
                        .WhereElementIsNotElementType()
                        .Take(1);
                    
                    foreach (Element element in collector)
                    {
                        foreach (Parameter param in element.Parameters)
                        {
                            if (string.Equals(param.Definition?.Name, "Fire Rating", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!hostParameters.Contains("Fire Rating"))
                                {
                                    hostParameters.Add("Fire Rating");
                                    break;
                                }
                            }
                        }
                        if (hostParameters.Contains("Fire Rating")) break;
                    }
                    if (hostParameters.Contains("Fire Rating")) break; // Found it, no need to check more files
                }
                
                DebugLogger.Info($"[ParameterServiceDialogV2] Found {hostParameters.Count} host parameters for category '{category}' (startup: only Fire Rating)");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Found {hostParameters.Count} host parameters for category '{category}' (startup: only Fire Rating)\n");
                }
                
                // Return Fire Rating if found, otherwise return empty array (user can add via "+" button)
                return hostParameters.Count > 0 
                    ? hostParameters.OrderBy(p => p).ToArray()
                    : essentialParams; // Return essentials even if not found (user can add it manually)
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[ParameterServiceDialogV2] Error loading host parameters for category '{category}': {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] ERROR loading host parameters for category '{category}': {ex.Message}\n");
                }
                // Fallback to essentials on error
                return essentialParams;
            }
        }
        
        /// <summary>
        /// Get Opening parameters for a specific category - loads from active document FamilySymbol objects
        /// ✅ PERFORMANCE FIX: Only loads 4 essential parameters on startup: MEP System Type, MEP Size, MEP System Abbreviation, Level
        /// Rest load when user clicks "+"
        /// </summary>
        private string[] GetOpeningParametersForCategory(string category)
        {
            // ✅ PERFORMANCE FIX: Only return 4 essential parameters on startup
            // 1. MEP System Type
            // 2. MEP Size  
            // 3. MEP System Abbreviation
            // 4. Level
            var essentialParams = new[] { "MEP System Type", "MEP Size", "MEP System Abbreviation", "Level" };
            
            if (_document == null)
            {
                // Fallback to essentials if no document available
                return essentialParams;
            }
            
            try
            {
                // Get opening FamilySymbol objects from active document (not instances, not linked files)
                var openingFamilies = GetOpeningFamilies(_document);
                
                if (openingFamilies.Count == 0)
                {
                    DebugLogger.Warning($"[ParameterServiceDialogV2] No opening families found in active document");
                    // Return essentials even if no families found
                    return essentialParams;
                }
                
                DebugLogger.Info($"[ParameterServiceDialogV2] Found {openingFamilies.Count} opening families in active document (startup: returning only 4 essential parameters)");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Found {openingFamilies.Count} opening families in active document (startup: returning only 4 essential parameters)\n");
                }
                
                // ✅ PERFORMANCE FIX: Only check if the 4 essential parameters exist, then return them
                // Don't scan all parameters - they'll load when user clicks "+"
                // Just verify the essential params exist in at least one family symbol
                var foundParams = new HashSet<string>();
                
                foreach (var familySymbol in openingFamilies.Take(1)) // Only check first family
                {
                    foreach (Parameter param in familySymbol.Parameters)
                    {
                        var paramName = param.Definition?.Name;
                        if (!string.IsNullOrEmpty(paramName) && 
                            essentialParams.Any(req => string.Equals(paramName, req, StringComparison.OrdinalIgnoreCase)))
                        {
                            foundParams.Add(paramName);
                        }
                    }
                    break; // Only check first family
                }
                
                // Return the essential params (even if not all found - user can add them later)
                DebugLogger.Info($"[ParameterServiceDialogV2] Returning {essentialParams.Length} essential opening parameters (startup mode)");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Returning {essentialParams.Length} essential opening parameters: {string.Join(", ", essentialParams)}\n");
                }
                return essentialParams;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[ParameterServiceDialogV2] Error loading opening parameters for category '{category}': {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] ERROR loading opening parameters for category '{category}': {ex.Message}\n");
                }
                // Fallback to essentials on error
                return essentialParams;
            }
        }
        
        /// <summary>
        /// Get opening FamilySymbol objects from active document (matching EmergencyMainDialog pattern)
        /// </summary>
        private List<FamilySymbol> GetOpeningFamilies(Document document)
        {
            var openingFamilies = new List<FamilySymbol>();

            try
            {
                // Specific opening family names to filter by
                var targetFamilyNames = new List<string>
                {
                    "RectangularOpeningOnWall",
                    "RectangularOpeningOnSlab",
                    "CircularOpeningOnWall",
                    "CircularOpeningOnSlab",
                    "Opening" // Also check for families containing "Opening" in name
                };

                // FamilySymbol is an ElementType; do NOT filter with WhereElementIsNotElementType
                var collector = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilySymbol));

                foreach (Element element in collector)
                {
                    if (element is FamilySymbol familySymbol)
                    {
                        var familyName = familySymbol.Family?.Name ?? "";
                        var symbolName = familySymbol.Name ?? "";

                        // Check if this family matches our target families or contains "Opening"
                        bool isTargetFamily = targetFamilyNames.Any(targetName =>
                            familyName.Contains(targetName, StringComparison.OrdinalIgnoreCase) ||
                            symbolName.Contains(targetName, StringComparison.OrdinalIgnoreCase) ||
                            $"{familyName} {symbolName}".Contains(targetName, StringComparison.OrdinalIgnoreCase)) ||
                            familyName.Contains("Opening", StringComparison.OrdinalIgnoreCase);

                        if (isTargetFamily && !openingFamilies.Contains(familySymbol))
                        {
                            openingFamilies.Add(familySymbol);
                        }
                    }
                }

                DebugLogger.Info($"[ParameterServiceDialogV2] GetOpeningFamilies found {openingFamilies.Count} opening families in active document");
                if (!DeploymentConfiguration.DeploymentMode && _document != null)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    string projectPath = ProjectPathService.GetProjectRoot(_document);
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] GetOpeningFamilies found {openingFamilies.Count} opening families\n");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Project path: {projectPath}\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error getting opening families: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] ERROR getting opening families: {ex.Message}\n");
                }
            }

            return openingFamilies;
        }

        /// <summary>
        /// Add a parameter mapping row
        /// </summary>
        private void AddParameterRow(WinForms.Panel panel, string category, string mepParam, string openingParam)
        {
            int rowHeight = 28;
            // Skip header panel (first control), count only row panels
            var rowCount = panel.Controls.OfType<WinForms.Panel>().Count() - 1;
            int top = 30 + (rowCount * rowHeight); // Start below header

            var row = new WinForms.Panel
            {
                Location = new Point(5, top),
                Size = new Size(panel.Width - 15, rowHeight),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            panel.Controls.Add(row);

            // MEP Parameter dropdown
            var mepCombo = new WinForms.ComboBox
            {
                Location = new Point(0, 3),
                Size = new Size(140, 22),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Tag = "mep" // Tag for identification when collecting mappings
            };
            var mepParams = GetMepParametersForCategory(category);
            mepCombo.Items.AddRange(mepParams);
            if (mepCombo.Items.Contains(mepParam))
                mepCombo.SelectedItem = mepParam;
            row.Controls.Add(mepCombo);

            // Arrow
            var arrow = new WinForms.Label { Text = "→", Location = new Point(145, 6), Size = new Size(20, 16) };
            row.Controls.Add(arrow);

            // Opening Parameter dropdown (width reduced by 20px, maintains Right anchor for proper sizing)
            var openingCombo = new WinForms.ComboBox
            {
                Location = new Point(170, 3),
                Size = new Size(row.Width - 200, 22), // Reduced by 20px (was 180, now 200)
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Tag = "opening" // Tag for identification when collecting mappings
            };
            var openingParams = GetOpeningParametersForCategory(category);
            openingCombo.Items.AddRange(openingParams);
            if (openingCombo.Items.Contains(openingParam))
                openingCombo.SelectedItem = openingParam;
            row.Controls.Add(openingCombo);

            // ✅ FIX: Explicitly set width AFTER adding to prevent Right anchor expansion
            openingCombo.Width = row.Width - 200; // Reduced by 20px (was 180, now 200)
            
            // ✅ FIX: Handle resize to maintain width
            row.Resize += (s, e) => {
                openingCombo.Width = row.Width - 200;
            };

            // Close button at far right
            var deleteBtn = new WinForms.Button
            {
                Text = "×",
                Location = new Point(row.Width - 30, 2), // Standard position at far right
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(255, 200, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            deleteBtn.Click += (s, e) => {
                panel.Controls.Remove(row);
                RepositionRows(panel);
            };
            row.Controls.Add(deleteBtn);
        }

        /// <summary>
        /// Reposition rows after deletion
        /// </summary>
        private void RepositionRows(WinForms.Panel panel)
        {
            var rows = panel.Controls.OfType<WinForms.Panel>().ToList();
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].Location = new Point(5, 5 + (i * 28));
            }
        }

        /// <summary>
        /// Create header row for System Type Overrides
        /// </summary>
        private void CreateSystemTypeHeader()
        {
            var header = new WinForms.Panel
            {
                Location = new Point(2, 2),
                Size = new Size(_systemTypeOverridesPanel.Width - 6, 25),
                BackColor = Color.FromArgb(220, 235, 250)
            };
            _systemTypeOverridesPanel.Controls.Add(header);

            // Column headers
            var valueLabel = new WinForms.Label
            {
                Text = "System Type",
                Location = new Point(3, 5),
                Size = new Size(100, 16),
                Font = new Font("Microsoft Sans Serif", 8.5F, FontStyle.Bold)
            };
            header.Controls.Add(valueLabel);

            var renameLabel = new WinForms.Label
            {
                Text = "Prefix",
                Location = new Point(140, 5),
                Size = new Size(60, 16),
                Font = new Font("Microsoft Sans Serif", 8.5F, FontStyle.Bold)
            };
            header.Controls.Add(renameLabel);
            
            // Remark column header
            var remarkLabel = new WinForms.Label
            {
                Text = "Remark",
                Location = new Point(215, 5),
                Size = new Size(70, 16),
                Font = new Font("Microsoft Sans Serif", 8.5F, FontStyle.Bold)
            };
            header.Controls.Add(remarkLabel);

            // "+" button aligned with close buttons (same X position)
            var addBtn = new WinForms.Button
            {
                Text = "+",
                Location = new Point(_systemTypeOverridesPanel.Width - 30, 2), // Aligned with close button X position
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(200, 255, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold),
                Anchor = AnchorStyles.Top | AnchorStyles.Right // Keep aligned to right edge
            };
            addBtn.Click += (s, e) => AddSystemTypeRow("<Select>", "");
            header.Controls.Add(addBtn);
        }

        /// <summary>
        /// Add a System Type Override row
        /// </summary>
        private void AddSystemTypeRow(string systemType, string prefix)
        {
            int rowHeight = 26;
            int top = 30 + (_systemTypeRows.Count * rowHeight);

            var row = new WinForms.Panel
            {
                Location = new Point(2, top),
                Size = new Size(_systemTypeOverridesPanel.Width - 6, rowHeight),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White
            };
            _systemTypeOverridesPanel.Controls.Add(row);

            // System Type dropdown (like ConVoid)
            var systemTypeCombo = new WinForms.ComboBox
            {
                Location = new Point(3, 3),
                Size = new Size(130, 22),
                DropDownStyle = ComboBoxStyle.DropDown
            };
            systemTypeCombo.Items.AddRange(new[] { "<Select>", "Exhaust Air", "Supply Air", "Return Air", "Sanitary", "Hydronic", "Fire Protection" });
            if (systemTypeCombo.Items.Contains(systemType))
                systemTypeCombo.SelectedItem = systemType;
            else
                systemTypeCombo.Text = systemType;
            row.Controls.Add(systemTypeCombo);

            // Arrow
            var arrow = new WinForms.Label { Text = "→", Location = new Point(138, 6), Size = new Size(15, 16) };
            row.Controls.Add(arrow);

            // Prefix textbox
            var prefixTextBox = new WinForms.TextBox
            {
                Location = new Point(158, 3),
                Size = new Size(50, 22),
                Text = prefix
            };
            row.Controls.Add(prefixTextBox);
            
            // Remark checkbox
            var remarkCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new Point(218, 5),
                Size = new Size(20, 20),
                Checked = false
            };
            row.Controls.Add(remarkCheckBox);

            // Close button (aligned with parameter mapping close buttons)
            var closeBtn = new WinForms.Button
            {
                Text = "×",
                Location = new Point(row.Width - 30, 2), // Same position as parameter rows
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(255, 200, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            closeBtn.Click += (s, e) => {
                _systemTypeOverridesPanel.Controls.Remove(row);
                _systemTypeRows.Remove(row);
                RepositionSystemTypeRows();
            };
            row.Controls.Add(closeBtn);

            _systemTypeRows.Add(row);

            // Update panel height (freely expand, no max limit)
            _systemTypeOverridesPanel.Height = _systemTypeRows.Count * rowHeight + 32;
        }

        /// <summary>
        /// Reposition System Type rows after deletion
        /// </summary>
        private void RepositionSystemTypeRows()
        {
            int rowHeight = 26;
            for (int i = 0; i < _systemTypeRows.Count; i++)
            {
                _systemTypeRows[i].Location = new Point(2, 30 + (i * rowHeight));
            }
        }
        
        /// <summary>
        /// Handle Transfer Parameters button click
        /// Transfers parameters to ALL sleeves (individual and cluster) based on configured mappings
        /// </summary>
        private void OnTransferParametersClick(object sender, EventArgs e)
        {
            try
            {
                if (_document == null || _uiDocument == null)
                {
                    WinForms.MessageBox.Show("Document not available.", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }
                
                // ✅ NEW: Check if filter XML files exist for all categories before processing
                // NOTE: Only Transfer Parameters needs XML validation because it reads parameter values FROM XML files
                // Apply Marks and Remark Selected work directly with Revit elements and don't need XML files
                var missingCategories = CheckFilterFilesForCategories();
                if (missingCategories.Count > 0)
                {
                    var categoryList = string.Join("\n• ", missingCategories);
                    var result = WinForms.MessageBox.Show(
                        $"⚠️ FILTER DATA NOT FOUND\n\n" +
                        $"The following categories cannot be processed because no filter data is found:\n\n" +
                        $"• {categoryList}\n\n" +
                        $"These categories require XML filter files (e.g., '*_ducts.xml', '*_pipes.xml') in the Filters directory.\n\n" +
                        $"Would you like to continue with available categories only?",
                        "Missing Filter Data",
                        WinForms.MessageBoxButtons.YesNo,
                        WinForms.MessageBoxIcon.Warning);
                    
                    if (result != WinForms.DialogResult.Yes)
                    {
                        return; // User cancelled
                    }
                }
                
                // Show progress dialog
                using (var progressForm = new WinForms.Form())
                {
                    progressForm.Text = "Transferring Parameters...";
                    progressForm.Size = new Size(400, 100);
                    progressForm.StartPosition = FormStartPosition.CenterScreen;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.MaximizeBox = false;
                    progressForm.MinimizeBox = false;
                    
                    var progressLabel = new WinForms.Label
                    {
                        Text = "Transferring parameters to all sleeves...",
                        Location = new Point(10, 30),
                        Size = new Size(380, 20),
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                    };
                    progressForm.Controls.Add(progressLabel);
                    
                    progressForm.Show();
                    progressForm.Refresh();
                    
                    // Create parameter transfer service
                    var transferService = new ParameterTransferService();
                    
                    // Get all openings (individual + cluster) in the document
                    // Only match the 4 specific opening families: RectangularOpeningOnWall, RectangularOpeningOnSlab, CircularOpeningOnWall, CircularOpeningOnSlab
                    var allOpeningInstances = new FilteredElementCollector(_document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => {
                            var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                            // Match only the 4 specific opening families
                            return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                                || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                        })
                        .ToList();
                    
                    // Log family names for debugging
                    var familyNames = allOpeningInstances
                        .Select(fi => fi.Symbol?.Family?.Name ?? "Unknown")
                        .Distinct()
                        .ToList();
                    DebugLogger.Info($"[ParameterServiceDialogV2] Found opening families: {string.Join(", ", familyNames)}");
                    
                    var openings = allOpeningInstances.Select(fi => fi.Id).ToList();
                    
                    DebugLogger.Info($"[ParameterServiceDialogV2] Found {openings.Count} opening sleeves in document for parameter transfer");
                    string transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
                    System.IO.File.AppendAllText(transferDebugLogPath, $"[{DateTime.Now}] [ParameterServiceDialogV2] Found {openings.Count} opening sleeves: {string.Join(", ", familyNames)}\n");
                    
                    if (openings.Count == 0)
                    {
                        progressForm.Close();
                        WinForms.MessageBox.Show("No openings found in the document. Please place sleeves first.", 
                            "No Openings", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                        return;
                    }
                    
                    // Collect parameter mappings from all category tabs (Reference Elements + Host Elements)
                    var allMappings = GetAllParameterMappingsFromUI();
                    
                    if (allMappings.Count == 0)
                    {
                        progressForm.Close();
                        WinForms.MessageBox.Show("No parameter mappings found. Please add parameter mappings first.",
                            "No Mappings", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
                        return;
                    }
                    
                    DebugLogger.Info($"[ParameterServiceDialogV2] Collected {allMappings.Count} parameter mappings from UI");
                    
                    // Create configuration with all mappings
                    var config = new Models.ParameterTransferConfiguration
                    {
                        SourceCategoryName = "All", // Transfer from all categories
                        Mappings = allMappings
                    };
                    
                    // ExecuteTransferConfiguration creates its own transaction, so we don't need to wrap it
                    var result = transferService.ExecuteTransferConfiguration(_document, openings, config);
                    
                    progressForm.Close();
                    
                    if (result.Success)
                    {
                        WinForms.MessageBox.Show(
                            $"Parameters transferred successfully!\nProcessed: {result.TransferredCount} sleeves\nErrors: {result.FailedCount}",
                            "Transfer Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    }
                    else
                    {
                        WinForms.MessageBox.Show($"Transfer failed: {result.Message}", 
                            "Transfer Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error during transfer: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Collect all parameter mappings from all category tabs (Reference Elements + Host Elements)
        /// </summary>
        private List<Models.ParameterMapping> GetAllParameterMappingsFromUI()
        {
            var allMappings = new List<Models.ParameterMapping>();
            
            try
            {
                // Get mappings from Reference Elements tabs
                if (_referenceParameterTabs != null)
                {
                    foreach (WinForms.TabPage tabPage in _referenceParameterTabs.TabPages)
                    {
                        var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                        if (servicePanel != null)
                        {
                            var mappings = GetParameterMappingsFromPanel(servicePanel);
                            allMappings.AddRange(mappings);
                            DebugLogger.Info($"[ParameterServiceDialogV2] Found {mappings.Count} mappings from Reference Elements tab: {tabPage.Text}");
                        }
                    }
                }
                
                // Get mappings from Host Elements tabs
                if (_hostParameterTabs != null)
                {
                    foreach (WinForms.TabPage tabPage in _hostParameterTabs.TabPages)
                    {
                        var servicePanel = tabPage.Controls.OfType<WinForms.Panel>().FirstOrDefault();
                        if (servicePanel != null)
                        {
                            var mappings = GetParameterMappingsFromPanel(servicePanel);
                            allMappings.AddRange(mappings);
                            DebugLogger.Info($"[ParameterServiceDialogV2] Found {mappings.Count} mappings from Host Elements tab: {tabPage.Text}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error collecting parameter mappings: {ex.Message}");
            }
            
            return allMappings;
        }
        
        /// <summary>
        /// Get parameter mappings from a service panel (same logic as old UI)
        /// </summary>
        private List<Models.ParameterMapping> GetParameterMappingsFromPanel(WinForms.Panel servicePanel)
        {
            var mappings = new List<Models.ParameterMapping>();
            
            try
            {
                // Skip header panel - get only parameter row panels
                var rows = servicePanel.Controls.OfType<WinForms.Panel>()
                    .Where(p => p.Controls.Count > 0) // Has controls (skip header or empty panels)
                    .ToList();
                
                foreach (var row in rows)
                {
                    var mepCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "mep");
                    var openingCombo = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault(c => c.Tag?.ToString() == "opening");
                    
                    if (mepCombo?.SelectedItem != null && openingCombo?.SelectedItem != null)
                    {
                        // Skip if either dropdown is empty or shows placeholder
                        string sourceParam = mepCombo.SelectedItem.ToString();
                        string targetParam = openingCombo.SelectedItem.ToString();
                        
                        if (!string.IsNullOrEmpty(sourceParam) && 
                            !string.IsNullOrEmpty(targetParam) &&
                            !sourceParam.Equals("<Select>", StringComparison.OrdinalIgnoreCase) &&
                            !targetParam.Equals("<Select>", StringComparison.OrdinalIgnoreCase))
                        {
                            mappings.Add(new Models.ParameterMapping
                            {
                                IsEnabled = true,
                                SourceParameter = sourceParam,
                                TargetParameter = targetParam,
                                TransferType = Models.TransferType.ReferenceToOpening
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error getting mappings from panel: {ex.Message}");
            }
            
            return mappings;
        }
        
        /// <summary>
        /// Handle Apply Marks button click
        /// Applies marks to sleeves based on selected remark checkboxes
        /// NOTE: This needs XML files to determine which category each sleeve belongs to (Ducts/Pipes/etc.)
        /// </summary>
        private void OnApplyMarksClick(object sender, EventArgs e)
        {
            try
            {
                if (_document == null || _uiDocument == null)
                {
                    WinForms.MessageBox.Show("Document not available.", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }
                
                // ✅ NEW: Check if filter XML files exist for all categories before processing
                // Apply Marks needs XML to determine which category each sleeve belongs to
                var missingCategories = CheckFilterFilesForCategories();
                if (missingCategories.Count > 0)
                {
                    var categoryList = string.Join("\n• ", missingCategories);
                    var result = WinForms.MessageBox.Show(
                        $"⚠️ FILTER DATA NOT FOUND\n\n" +
                        $"The following categories cannot be processed because no filter data is found:\n\n" +
                        $"• {categoryList}\n\n" +
                        $"Apply Marks requires XML filter files to determine sleeve categories (e.g., '*_ducts.xml', '*_pipes.xml') in the Filters directory.\n\n" +
                        $"Would you like to continue with available categories only?",
                        "Missing Filter Data",
                        WinForms.MessageBoxButtons.YesNo,
                        WinForms.MessageBoxIcon.Warning);
                    
                    if (result != WinForms.DialogResult.Yes)
                    {
                        return; // User cancelled
                    }
                }
                
                // Collect settings from UI
                var projectPrefix = _projectPrefixTextBox.Text.Trim();
                var numberFormat = _numberFormatCombo.SelectedIndex switch
                {
                    0 => "00",
                    1 => "000",
                    2 => "0000",
                    _ => "000"
                };
                
                var remarkProject = _remarkProjectCheckBox.Checked;
                var remarkDuct = _remarkDuctCheckBox.Checked;
                var remarkPipe = _remarkPipeCheckBox.Checked;
                var remarkCableTray = _remarkCableTrayCheckBox.Checked;
                var remarkDamper = _remarkDamperCheckBox.Checked;
                
                // Show progress dialog
                using (var progressForm = new WinForms.Form())
                {
                    progressForm.Text = "Applying Marks...";
                    progressForm.Size = new Size(400, 100);
                    progressForm.StartPosition = FormStartPosition.CenterScreen;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.MaximizeBox = false;
                    progressForm.MinimizeBox = false;
                    
                    var progressLabel = new WinForms.Label
                    {
                        Text = "Applying marks to sleeves with remark checkboxes enabled...",
                        Location = new Point(10, 30),
                        Size = new Size(380, 20),
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                    };
                    progressForm.Controls.Add(progressLabel);
                    
                    progressForm.Show();
                    progressForm.Refresh();
                    
                    // Prepare mark prefixes settings
                    var ductPrefix = _ductPrefixTextBox.Text.Trim();
                    var pipePrefix = _pipePrefixTextBox.Text.Trim();
                    var cableTrayPrefix = _cableTrayPrefixTextBox.Text.Trim();
                    var damperPrefix = _damperPrefixTextBox.Text.Trim();
                    
                    var markPrefixes = new Models.MarkPrefixSettings
                    {
                        ProjectPrefix = projectPrefix,
                        DuctPrefix = ductPrefix,
                        PipePrefix = pipePrefix,
                        CableTrayPrefix = cableTrayPrefix,
                        DamperPrefix = damperPrefix,
                        NumberFormat = numberFormat
                    };
                    
                    // Debug: Log the prefix values being used
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Apply Marks - Project: '{projectPrefix}', Duct: '{ductPrefix}', Pipe: '{pipePrefix}', CableTray: '{cableTrayPrefix}', Damper: '{damperPrefix}', Format: '{numberFormat}'\n");
                    
                    // Apply Marks: Mark ALL sleeves regardless of checkbox state
                    // remarkAll=false means skip sleeves that already have marks
                    var cmd = new MarkParameterCommand("ALL", projectPrefix, "", false, markPrefixes);
                    cmd.Execute(_uiDocument.Application);
                    
                    progressForm.Close();
                    
                    var prefixSummary = string.IsNullOrWhiteSpace(projectPrefix) 
                        ? $"Duct:{ductPrefix}, Pipe:{pipePrefix}, CableTray:{cableTrayPrefix}, Damper:{damperPrefix}"
                        : $"Project:{projectPrefix}, Duct:{ductPrefix}, Pipe:{pipePrefix}, CableTray:{cableTrayPrefix}, Damper:{damperPrefix}";
                    
                    WinForms.MessageBox.Show($"Marks applied to all categories using:\n{prefixSummary}", 
                        "Apply Marks Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error during mark application: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// Handle Remark Selected button click
        /// Re-marks only the categories with remark checkboxes checked
        /// NOTE: This DOES need XML files to determine which sleeves belong to each category
        /// </summary>
        private void OnRemarkSelectedClick(object sender, EventArgs e)
        {
            try
            {
                if (_document == null || _uiDocument == null)
                {
                    WinForms.MessageBox.Show("Document not available.", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }
                
                // ✅ NEW: Check if filter XML files exist for selected categories before processing
                // Remark Selected needs XML to determine which sleeves belong to each category
                var checkedCategories = new List<string>();
                if (_remarkDuctCheckBox.Checked) checkedCategories.Add("Ducts");
                if (_remarkPipeCheckBox.Checked) checkedCategories.Add("Pipes");
                if (_remarkCableTrayCheckBox.Checked) checkedCategories.Add("Cable Trays");
                if (_remarkDamperCheckBox.Checked) checkedCategories.Add("Duct Accessories");
                
                if (checkedCategories.Count > 0)
                {
                    var missingCategories = CheckFilterFilesForCategories()
                        .Where(cat => checkedCategories.Contains(cat))
                        .ToList();
                    
                    if (missingCategories.Count > 0)
                    {
                        var categoryList = string.Join("\n• ", missingCategories);
                        var result = WinForms.MessageBox.Show(
                            $"⚠️ FILTER DATA NOT FOUND\n\n" +
                            $"The following selected categories cannot be processed because no filter data is found:\n\n" +
                            $"• {categoryList}\n\n" +
                            $"Remark Selected requires XML filter files to determine sleeve categories (e.g., '*_ducts.xml', '*_pipes.xml') in the Filters directory.\n\n" +
                            $"Would you like to continue with available categories only?",
                            "Missing Filter Data",
                            WinForms.MessageBoxButtons.YesNo,
                            WinForms.MessageBoxIcon.Warning);
                        
                        if (result != WinForms.DialogResult.Yes)
                        {
                            return; // User cancelled
                        }
                    }
                }
                
                // Collect remark checkbox states
                var remarkProject = _remarkProjectCheckBox.Checked;
                var remarkDuct = _remarkDuctCheckBox.Checked;
                var remarkPipe = _remarkPipeCheckBox.Checked;
                var remarkCableTray = _remarkCableTrayCheckBox.Checked;
                var remarkDamper = _remarkDamperCheckBox.Checked;
                
                // Collect prefix values
                var projectPrefix = _projectPrefixTextBox.Text.Trim();
                var ductPrefix = _ductPrefixTextBox.Text.Trim();
                var pipePrefix = _pipePrefixTextBox.Text.Trim();
                var cableTrayPrefix = _cableTrayPrefixTextBox.Text.Trim();
                var damperPrefix = _damperPrefixTextBox.Text.Trim();
                
                var numberFormat = _numberFormatCombo.SelectedIndex switch
                {
                    0 => "00",
                    1 => "000",
                    2 => "0000",
                    _ => "000"
                };
                
                // Collect checked System Type Overrides with remark checkboxes
                var systemTypeOverrides = new List<(string systemType, string prefix)>();
                foreach (var row in _systemTypeRows)
                {
                    // Find checkbox in row
                    var checkbox = row.Controls.OfType<WinForms.CheckBox>().FirstOrDefault();
                    if (checkbox != null && checkbox.Checked)
                    {
                        var comboBox = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault();
                        var textBox = row.Controls.OfType<WinForms.TextBox>().FirstOrDefault();
                        
                        if (comboBox != null && textBox != null)
                        {
                            systemTypeOverrides.Add((comboBox.Text, textBox.Text));
                        }
                    }
                }
                
                // Show progress dialog
                using (var progressForm = new WinForms.Form())
                {
                    progressForm.Text = "Remarking Selected Categories...";
                    progressForm.Size = new Size(400, 100);
                    progressForm.StartPosition = FormStartPosition.CenterScreen;
                    progressForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    progressForm.MaximizeBox = false;
                    progressForm.MinimizeBox = false;
                    
                    var progressLabel = new WinForms.Label
                    {
                        Text = "Remarking selected categories...",
                        Location = new Point(10, 30),
                        Size = new Size(380, 20),
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                    };
                    progressForm.Controls.Add(progressLabel);
                    
                    progressForm.Show();
                    progressForm.Refresh();
                    
                    // Prepare mark prefixes settings
                    var markPrefixes = new Models.MarkPrefixSettings
                    {
                        ProjectPrefix = projectPrefix,
                        DuctPrefix = ductPrefix,
                        PipePrefix = pipePrefix,
                        CableTrayPrefix = cableTrayPrefix,
                        DamperPrefix = damperPrefix,
                        NumberFormat = numberFormat
                    };
                    
                    // Remark only checked categories with remarkAll=true
                    int totalProcessed = 0;
                    var categoriesProcessed = new List<string>();
                    
                    // If Project Prefix remark is checked, re-mark ALL categories with new project prefix
                    if (remarkProject)
                    {
                        var cmd = new MarkParameterCommand("ALL", projectPrefix, "", true, markPrefixes);
                        cmd.Execute(_uiDocument.Application);
                        categoriesProcessed.Add("All Categories (Project Prefix)");
                        totalProcessed++;
                    }
                    
                    // Process individual discipline categories if their checkboxes are checked
                    if (remarkDuct)
                    {
                        var cmd = new MarkParameterCommand("Ducts", projectPrefix, ductPrefix, true, markPrefixes);
                        cmd.Execute(_uiDocument.Application);
                        totalProcessed++;
                        categoriesProcessed.Add("Ducts");
                    }
                    if (remarkPipe)
                    {
                        var cmd = new MarkParameterCommand("Pipes", projectPrefix, pipePrefix, true, markPrefixes);
                        cmd.Execute(_uiDocument.Application);
                        totalProcessed++;
                        categoriesProcessed.Add("Pipes");
                    }
                    if (remarkCableTray)
                    {
                        var cmd = new MarkParameterCommand("Cable Trays", projectPrefix, cableTrayPrefix, true, markPrefixes);
                        cmd.Execute(_uiDocument.Application);
                        totalProcessed++;
                        categoriesProcessed.Add("Cable Trays");
                    }
                    if (remarkDamper)
                    {
                        var cmd = new MarkParameterCommand("Duct Accessories", projectPrefix, damperPrefix, true, markPrefixes);
                        cmd.Execute(_uiDocument.Application);
                        totalProcessed++;
                        categoriesProcessed.Add("Duct Accessories");
                    }
                    
                    progressForm.Close();
                    
                    if (totalProcessed > 0)
                    {
                        WinForms.MessageBox.Show($"Re-marked {totalProcessed} categories:\n{string.Join(", ", categoriesProcessed)}", 
                            "Remark Selected Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    }
                    else
                    {
                        WinForms.MessageBox.Show("No categories selected for re-marking. Please check remark checkboxes.", 
                            "No Selection", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error during remark: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }
        
        /// <summary>
        /// ✅ NEW: Check if filter XML files exist for all MEP categories
        /// Returns list of categories that don't have filter data
        /// </summary>
        private List<string> CheckFilterFilesForCategories()
        {
            var missingCategories = new List<string>();
            
            try
            {
                if (_document == null) return missingCategories;
                
                // Get filters directory
                ProjectPathService.EnsureFiltersDirectory(_document);
                string filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    // No filters directory - all categories missing
                    return new List<string> { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                }
                
                // Check each MEP category (only Reference Elements categories need filter files)
                var categoriesToCheck = new Dictionary<string, string[]>
                {
                    { "Ducts", new[] { "*_ducts.xml", "*ducts*.xml", "ducts*.xml" } },
                    { "Pipes", new[] { "*_pipes.xml", "*pipes*.xml", "pipes*.xml" } },
                    { "Cable Trays", new[] { "*_cable_trays.xml", "*cable_trays*.xml", "*cable_tray*.xml", "*cabletray*.xml" } },
                    { "Duct Accessories", new[] { "*_duct_accessories.xml", "*duct_accessories*.xml", "*damper*.xml" } }
                };
                
                foreach (var categoryPair in categoriesToCheck)
                {
                    string category = categoryPair.Key;
                    string[] patterns = categoryPair.Value;
                    
                    bool found = false;
                    foreach (var pattern in patterns)
                    {
                        var files = Directory.GetFiles(filtersDirectory, pattern);
                        if (files.Length > 0)
                        {
                            // Found at least one file - check if it has clash zones
                            foreach (var file in files)
                            {
                                try
                                {
                                    // Quick check: deserialize and see if it has clash zones
                                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                                    using (var reader = new System.IO.StreamReader(file))
                                    {
                                        var filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                                        if (filter?.ClashZoneStorage?.ClashZones != null && filter.ClashZoneStorage.ClashZones.Count > 0)
                                        {
                                            found = true;
                                            break;
                                        }
                                    }
                                }
                                catch
                                {
                                    // If deserialization fails, skip this file
                                    continue;
                                }
                            }
                        }
                        
                        if (found) break;
                    }
                    
                    if (!found)
                    {
                        missingCategories.Add(category);
                    }
                }
                
                DebugLogger.Info($"[ParameterServiceDialogV2] Filter file check: {missingCategories.Count} categories missing filter data");
                if (missingCategories.Count > 0)
                {
                    DebugLogger.Warning($"[ParameterServiceDialogV2] Missing filter files for: {string.Join(", ", missingCategories)}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error checking filter files: {ex.Message}");
                // On error, don't block processing - return empty list
            }
            
            return missingCategories;
        }
    }
}

