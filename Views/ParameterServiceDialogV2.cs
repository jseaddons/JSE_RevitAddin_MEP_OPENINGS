using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;

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

        // Right Panel: Parameter Mapping
        private WinForms.TabControl _masterParameterTabs = null!;
        private WinForms.TabControl _referenceParameterTabs = null!;
        private WinForms.TabControl _hostParameterTabs = null!;
        private WinForms.Button _addParameterButton = null!;

        // Top bar
        private WinForms.Button _transferParametersButton = null!;

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

            // Transfer Parameters button
            _transferParametersButton = new WinForms.Button
            {
                Text = "Transfer Parameters",
                Size = new Size(150, 32),
                Location = new Point(this.Width - 170, 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(100, 150, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _transferParametersButton.Click += (s, e) => WinForms.MessageBox.Show("Transfer Parameters clicked!");
            topBar.Controls.Add(_transferParametersButton);
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
                Size = new Size(200, 22),
                Text = "SLEEVE_"
            };
            _leftPrefixPanel.Controls.Add(_projectPrefixTextBox);
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
            yPos += 30;

            // Row 2: Pipe
            var pipeLabel = new WinForms.Label { Text = "Pipe:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(pipeLabel);
            _pipePrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "P" };
            _leftPrefixPanel.Controls.Add(_pipePrefixTextBox);
            yPos += 30;

            // Row 3: Cable Tray
            var cableTrayLabel = new WinForms.Label { Text = "Cable Tray:", Location = new Point(10, yPos), Size = new Size(70, 20) };
            _leftPrefixPanel.Controls.Add(cableTrayLabel);
            _cableTrayPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "E" };
            _leftPrefixPanel.Controls.Add(_cableTrayPrefixTextBox);
            yPos += 30;

            // Row 4: Damper
            var damperLabel = new WinForms.Label { Text = "Damper:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(damperLabel);
            _damperPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "DMP" };
            _leftPrefixPanel.Controls.Add(_damperPrefixTextBox);
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
                Location = new Point(headerPanel.Width - 30, 2), // Same X as close buttons
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(200, 255, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
            };
            addButton.Click += (s, e) => {
                AddParameterRow(servicePanel, "<Select>", "");
            };
            headerPanel.Controls.Add(addButton);

            // Add default parameter rows
            AddDefaultParameterRows(servicePanel, categoryName);
            tabPage.Controls.Add(servicePanel);
        }

        /// <summary>
        /// Add default parameter rows for a category
        /// </summary>
        private void AddDefaultParameterRows(WinForms.Panel panel, string category)
        {
            // Define default mappings
            var mappings = new List<(string Mep, string Opening)>
            {
                ("Size", "MEP Size"),
                ("System Type", "MEP System Type"),
                ("Reference Level", "Level"),
                ("MEP System Name", "System Name")
            };

            foreach (var mapping in mappings)
            {
                AddParameterRow(panel, mapping.Mep, mapping.Opening);
            }
        }

        /// <summary>
        /// Add a parameter mapping row
        /// </summary>
        private void AddParameterRow(WinForms.Panel panel, string mepParam, string openingParam)
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
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            mepCombo.Items.AddRange(new[] { "Size", "System Type", "Service Type", "Reference Level", "MEP System Name" });
            if (mepCombo.Items.Contains(mepParam))
                mepCombo.SelectedItem = mepParam;
            row.Controls.Add(mepCombo);

            // Arrow
            var arrow = new WinForms.Label { Text = "→", Location = new Point(145, 6), Size = new Size(20, 16) };
            row.Controls.Add(arrow);

            // Opening Parameter dropdown (reduced width to create gap before close button)
            var openingCombo = new WinForms.ComboBox
            {
                Location = new Point(170, 3),
                Size = new Size(row.Width - 245, 22), // Reduce width to leave clear gap
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            openingCombo.Items.AddRange(new[] { "MEP Size", "MEP System Type", "Level", "System Name", "Mark" });
            if (openingCombo.Items.Contains(openingParam))
                openingCombo.SelectedItem = openingParam;
            row.Controls.Add(openingCombo);

            // ✅ FIXED: Close button with clear gap and bold font
            var deleteBtn = new WinForms.Button
            {
                Text = "×",
                Location = new Point(row.Width - 30, 2), // Leave 5px gap
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(255, 200, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
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

            // "+" button aligned with close buttons (same X position)
            var addBtn = new WinForms.Button
            {
                Text = "+",
                Location = new Point(_systemTypeOverridesPanel.Width - 30, 2), // Same X as close
                Size = new Size(22, 22),
                BackColor = Color.FromArgb(200, 255, 200),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 10F, FontStyle.Bold)
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
    }
}

