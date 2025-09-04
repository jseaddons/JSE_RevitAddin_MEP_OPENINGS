using System;
using System.Drawing;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Emergency WinForms MainDialog - NO WPF - CRASH SAFE for Revit
    /// 4-section layout with toolbar and status
    /// </summary>
    public partial class EmergencyMainDialog : WinForms.Form
    {
        private readonly ApplicationProfileService _appProfileService;
        
        // Main panels - 4-section layout
        private WinForms.Panel _leftPanel = null!;
        private WinForms.Panel _rightPanel = null!;
        private WinForms.Splitter _mainSplitter = null!;
        
        // 4 sections within left panel (2x2 grid)
        private WinForms.Panel _topLeftPanel = null!;      // Reference Elements (linked files)
        private WinForms.Panel _topRightPanel = null!;     // Reference Categories (MEP categories)
        private WinForms.Panel _bottomLeftPanel = null!;   // Host Elements (linked files)
        private WinForms.Panel _bottomRightPanel = null!;  // Host Categories (Revit categories)
        private WinForms.Splitter _horizontalSplitter = null!;
        private WinForms.Splitter _verticalSplitter = null!;
        
        // Header and toolbar
        private WinForms.Panel _headerPanel = null!;
        private WinForms.Label _titleLabel = null!;
        private WinForms.Label _profileLabel = null!;
        
        // Bottom status bar
        private WinForms.Panel _statusPanel = null!;
        private WinForms.Label _statusLabel = null!;
        private WinForms.ProgressBar _progressBar = null!;
        
        // Toolbar buttons
        private WinForms.Button _okButton = null!;
        private WinForms.Button _cancelButton = null!;
        private WinForms.Button _saveButton = null!;
        private WinForms.Button _closeButton = null!;
        
        // Dynamic UI controls (only what we actually use)
        private WinForms.ComboBox _mepTypeCombo = null!;
        private WinForms.Panel _clearancePanel = null!;
        private WinForms.Panel _cableTrayPanel = null!;
        private WinForms.Panel _damperPanel = null!;
        private LinkedFileService? _linkedFileService;

        public EmergencyMainDialog(ApplicationProfileService appProfileService, Document? document = null)
        {
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            _linkedFileService = new LinkedFileService();
            InitializeComponent();
            LoadProfileInfo();
            
            // Load real linked files if document is provided
            if (document != null)
            {
                LoadRealLinkedFiles(document);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = "JSE MEP Openings - Main Interface (Emergency WinForms Mode)";
            this.Size = new System.Drawing.Size(1200, 800);
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(1000, 600);
            this.TopMost = true;

            // Header Panel
            _headerPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Top,
                Height = 80,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_headerPanel);

            // Title Label
            _titleLabel = new WinForms.Label
            {
                Text = "MEP Openings Management",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(20, 15),
                Size = new System.Drawing.Size(400, 30),
                AutoSize = false
            };
            _headerPanel.Controls.Add(_titleLabel);

            // Profile Label
            _profileLabel = new WinForms.Label
            {
                Text = "Profile: Loading...",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(400, 20),
                AutoSize = false
            };
            _headerPanel.Controls.Add(_profileLabel);

            // Toolbar buttons in header - moved further left
            int buttonY = 20;
            int buttonWidth = 70;
            int buttonHeight = 35;
            int buttonSpacing = 80;
            int startX = 450;  // Change this line to move buttons closer

            _okButton = new WinForms.Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(startX, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _okButton.Click += OnOkClick;
            _headerPanel.Controls.Add(_okButton);

            _saveButton = new WinForms.Button
            {
                Text = "Save",
                Location = new System.Drawing.Point(startX + buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(40, 167, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _saveButton.Click += OnSaveClick;
            _headerPanel.Controls.Add(_saveButton);

            _cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(startX + 2 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _cancelButton.Click += OnCancelClick;
            _headerPanel.Controls.Add(_cancelButton);

            _closeButton = new WinForms.Button
            {
                Text = "Close",
                Location = new System.Drawing.Point(startX + 3 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(220, 53, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold)
            };
            _closeButton.Click += OnCloseClick;
            _headerPanel.Controls.Add(_closeButton);

            // Status Panel (bottom)
            _statusPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_statusPanel);

            _statusLabel = new WinForms.Label
            {
                Text = "Ready",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(10, 5),
                Size = new System.Drawing.Size(600, 20),
                AutoSize = false
            };
            _statusPanel.Controls.Add(_statusLabel);

            _progressBar = new WinForms.ProgressBar
            {
                Location = new System.Drawing.Point(this.Width - 220, 5),
                Size = new System.Drawing.Size(200, 20),
                Style = WinForms.ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            _statusPanel.Controls.Add(_progressBar);

            // Left Panel (expanded to fill most space) - start below header
            _leftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.None,  // Change from Fill to None
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Location = new System.Drawing.Point(0, 80),  // Start below header (80px)
                Size = new System.Drawing.Size(this.Width - 300, this.Height - 110)  // Full width minus right panel, minus header/status
            };
            this.Controls.Add(_leftPanel);

            // Main Splitter
            _mainSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Right,
                Width = 3,
                BackColor = System.Drawing.Color.Gray
            };
            this.Controls.Add(_mainSplitter);

            // Right Panel - reduced width
            _rightPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Right,
                Width = 300,
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            this.Controls.Add(_rightPanel);

            // Add content to panels
            InitializePanelContent();

            this.ResumeLayout(false);
        }

        private void InitializePanelContent()
        {
            // Create 4-section layout within left panel
            CreateFourSectionLayout();
            
            // Right Panel: Opening Configuration
            InitializeRightPanel();
        }

        private void CreateFourSectionLayout()
        {
            // Top-Left Panel: Reference Elements (linked files)
            _topLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Fill,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _leftPanel.Controls.Add(_topLeftPanel);

            // Horizontal Splitter (between top and bottom)
            _horizontalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _leftPanel.Controls.Add(_horizontalSplitter);

            // Bottom-Left Panel: Host Elements (linked files)
            _bottomLeftPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Bottom,
                Height = _leftPanel.Height / 2,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _leftPanel.Controls.Add(_bottomLeftPanel);

            // Now create the right side panels within the top and bottom panels
            CreateTopRightPanel();
            CreateBottomRightPanel();

            // Populate each section with content
            PopulateTopLeftSection();      // Reference Elements (linked files)
            PopulateTopRightSection();     // Reference Categories (MEP categories)
            PopulateBottomLeftSection();   // Host Elements (linked files)
            PopulateBottomRightSection();  // Host Categories (Revit categories)
        }

        private void CreateTopRightPanel()
        {
            // Vertical Splitter (between left and right in top section)
            _verticalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Right,
                Width = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _topLeftPanel.Controls.Add(_verticalSplitter);

            // Top-Right Panel: Reference Categories (MEP categories)
            _topRightPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Right,
                Width = _topLeftPanel.Width / 2,
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _topLeftPanel.Controls.Add(_topRightPanel);
        }

        private void CreateBottomRightPanel()
        {
            // Vertical Splitter (between left and right in bottom section)
            var bottomVerticalSplitter = new WinForms.Splitter
            {
                Dock = WinForms.DockStyle.Right,
                Width = 3,
                BackColor = System.Drawing.Color.Gray
            };
            _bottomLeftPanel.Controls.Add(bottomVerticalSplitter);

            // Bottom-Right Panel: Host Categories (Revit categories)
            _bottomRightPanel = new WinForms.Panel
            {
                Dock = WinForms.DockStyle.Right,
                Width = _bottomLeftPanel.Width / 2,
                BackColor = System.Drawing.Color.FromArgb(250, 250, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _bottomLeftPanel.Controls.Add(_bottomRightPanel);
        }

        private void PopulateTopLeftSection()
        {
            // Title: "Reference Elements (MEP Files)" - make it shorter
            var title = new WinForms.Label
            {
                Text = "Reference Elements",  // Remove "(MEP Files)" to make it shorter
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),  // Make it wider to fit
                AutoSize = false
            };
            _topLeftPanel.Controls.Add(title);

            // ListBox for linked files with checkboxes - starts right after title
            var referenceFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 40),
                Size = new System.Drawing.Size(_topLeftPanel.Width - 10, _topLeftPanel.Height - 25),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topLeftPanel.Controls.Add(referenceFilesListBox);

            // Add MEP files (EL, ME, PH, FF) - filtered from sample data
            referenceFilesListBox.Items.Add("Current Project - OCC-BEC-NA-M", false);
            referenceFilesListBox.Items.Add("OCC-BEC-NA-MS-EL-M3D-4021", true);
            referenceFilesListBox.Items.Add("OCC-BEC-NA-MS-EL-M3D-4022", false);
            referenceFilesListBox.Items.Add("OCC-BEC-NA-MS-ME-M3D-4021", false);
            referenceFilesListBox.Items.Add("OCC-BEC-NA-MS-PH-M3D-4021", false);
            referenceFilesListBox.Items.Add("OCC-BEC-NA-MS-FF-M3D-4021", false);
        }

        private void PopulateTopRightSection()
        {
            // Title: "MEP Categories" - smaller and positioned at top
            var title = new WinForms.Label
            {
                Text = "MEP Categories",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(100, 25),
                AutoSize = false
            };
            _topRightPanel.Controls.Add(title);

            // ListBox for MEP categories with checkboxes - move down more
            var referenceCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 35),  // Change from 30 to 35 (more space)
                Size = new System.Drawing.Size(_topRightPanel.Width - 10, _topRightPanel.Height - 40),  // Change from 35 to 40
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _topRightPanel.Controls.Add(referenceCategoriesListBox);

            // Add MEP categories - all visible now
            referenceCategoriesListBox.Items.Add("Cable Tray Fittings", false);
            referenceCategoriesListBox.Items.Add("Cable Trays", true);
            referenceCategoriesListBox.Items.Add("Communication Devices", false);
            referenceCategoriesListBox.Items.Add("Conduit Fittings", false);
            referenceCategoriesListBox.Items.Add("Conduits", true);
            referenceCategoriesListBox.Items.Add("Electrical Equipment", false);
            referenceCategoriesListBox.Items.Add("Electrical Fixtures", false);
            referenceCategoriesListBox.Items.Add("Duct Curves", false);
            referenceCategoriesListBox.Items.Add("Duct Fittings", false);
            referenceCategoriesListBox.Items.Add("Duct Accessories", false);
            referenceCategoriesListBox.Items.Add("Pipe Curves", false);
            referenceCategoriesListBox.Items.Add("Pipe Fittings", false);
        }

        private void PopulateBottomLeftSection()
        {
            // Title: "Host Elements (ARC/STR Files)" - make it shorter
            var title = new WinForms.Label
            {
                Text = "Host Elements",  // Remove "(ARC/STR Files)" to make it shorter
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(150, 30),  // Make it wider to fit
                AutoSize = false
            };
            _bottomLeftPanel.Controls.Add(title);

            // ListBox for host files with checkboxes - starts right after title
            var hostFilesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 30),
                Size = new System.Drawing.Size(_bottomLeftPanel.Width - 10, _bottomLeftPanel.Height - 25),
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right,
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomLeftPanel.Controls.Add(hostFilesListBox);

            // Add only ARC and STR files - filtered
            hostFilesListBox.Items.Add("Current Project - OCC-BEC-NA-MS-", false);
            hostFilesListBox.Items.Add("OCC-TA-NA-MS-ARC-M3D-3021", true);
            hostFilesListBox.Items.Add("OCC-TA-NA-MS-STR-M3D-3021", true);
            hostFilesListBox.Items.Add("OCC-TA-NA-MS-ARC-M3D-3022", false);
            hostFilesListBox.Items.Add("OCC-TA-NA-MS-STR-M3D-3022", false);
        }

        private void PopulateBottomRightSection()
        {
            // Title: "Host Categories" - smaller and positioned at top
            var title = new WinForms.Label
            {
                Text = "Host Categories",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(5, 2),
                Size = new System.Drawing.Size(100, 16),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(title);

            // Horizontal Openings Section - starts right after title
            var horizontalLabel = new WinForms.Label
            {
                Text = "Horizontal Openings:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 100, 200),
                Location = new System.Drawing.Point(5, 20),
                Size = new System.Drawing.Size(100, 14),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(horizontalLabel);

            // ListBox for horizontal host categories
            var horizontalCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 36),
                Size = new System.Drawing.Size(_bottomRightPanel.Width - 10, 60),
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomRightPanel.Controls.Add(horizontalCategoriesListBox);

            // Add horizontal categories - Walls and Structural Framing only
            horizontalCategoriesListBox.Items.Add("Walls", true);
            horizontalCategoriesListBox.Items.Add("Structural Framing", true);

            // Vertical Openings Section - positioned below horizontal
            var verticalLabel = new WinForms.Label
            {
                Text = "Vertical Openings:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(200, 100, 0),
                Location = new System.Drawing.Point(5, 100),
                Size = new System.Drawing.Size(100, 14),
                AutoSize = false
            };
            _bottomRightPanel.Controls.Add(verticalLabel);

            // ListBox for vertical host categories
            var verticalCategoriesListBox = new WinForms.CheckedListBox
            {
                Location = new System.Drawing.Point(5, 116),
                Size = new System.Drawing.Size(_bottomRightPanel.Width - 10, 60),
                CheckOnClick = true,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular)
            };
            _bottomRightPanel.Controls.Add(verticalCategoriesListBox);

            // Add vertical categories - Floors and Ceilings only
            verticalCategoriesListBox.Items.Add("Floors", true);
            verticalCategoriesListBox.Items.Add("Ceilings", true);
        }

        private void InitializeRightPanel()
        {
            // Title
            var rightTitle = new WinForms.Label
            {
                Text = "Opening Configuration & Clearances",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 25),
                AutoSize = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(rightTitle);

            // MEP Type Selection
            var mepTypeLabel = new WinForms.Label
            {
                Text = "MEP Type:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 45),
                Size = new System.Drawing.Size(100, 20),
                AutoSize = false
            };
            _rightPanel.Controls.Add(mepTypeLabel);

            _mepTypeCombo = new WinForms.ComboBox
            {
                Location = new System.Drawing.Point(120, 45),
                Size = new System.Drawing.Size(_rightPanel.Width - 130, 20),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _mepTypeCombo.Items.AddRange(new[] { "Pipe", "Duct", "Duct Accessories", "Duct Fittings", "Cable Tray", "Conduit" });
            _mepTypeCombo.SelectedIndex = 0;
            _rightPanel.Controls.Add(_mepTypeCombo);

            // Create clearance panels
            CreateClearancePanels();
        }

        private void CreateClearancePanels()
        {
            // Standard Clearance Panel
            _clearancePanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 80),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 100),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_clearancePanel);

            var clearanceLabel = new WinForms.Label
            {
                Text = "Clearance per Side:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(150, 20),
                AutoSize = false
            };
            _clearancePanel.Controls.Add(clearanceLabel);

            var clearanceTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(100, 20),
                Text = "25"
            };
            _clearancePanel.Controls.Add(clearanceTextBox);

            var mmLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new System.Drawing.Point(120, 37),
                Size = new System.Drawing.Size(20, 16),
                AutoSize = false
            };
            _clearancePanel.Controls.Add(mmLabel);

            // Cable Tray Panel (initially hidden)
            _cableTrayPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 80),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 120),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_cableTrayPanel);

            // Damper Panel (initially hidden)
            _damperPanel = new WinForms.Panel
            {
                Location = new System.Drawing.Point(10, 80),
                Size = new System.Drawing.Size(_rightPanel.Width - 20, 120),
                BackColor = System.Drawing.Color.FromArgb(248, 249, 250),
                BorderStyle = WinForms.BorderStyle.FixedSingle,
                Visible = false,
                Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right
            };
            _rightPanel.Controls.Add(_damperPanel);
        }

        private void LoadProfileInfo()
        {
            try
            {
                var currentProfile = _appProfileService.CurrentProfile;
                if (currentProfile != null)
                {
                    _profileLabel.Text = $"Profile: {currentProfile.Name}";
                }
                else
                {
                    _profileLabel.Text = "Profile: None";
                }
            }
            catch (Exception ex)
            {
                _profileLabel.Text = $"Profile: Error - {ex.Message}";
            }
        }

        private void LoadRealLinkedFiles(Document document)
        {
            // This will be implemented to load actual linked files from Revit
            // For now, we have sample data in the populate methods
        }

        // Event handlers
        private void OnOkClick(object? sender, EventArgs e)
        {
            _statusLabel.Text = "OK clicked";
        }

        private void OnSaveClick(object? sender, EventArgs e)
        {
            _statusLabel.Text = "Save clicked";
        }

        private void OnCancelClick(object? sender, EventArgs e)
        {
            _statusLabel.Text = "Cancel clicked";
        }

        private void OnCloseClick(object? sender, EventArgs e)
        {
            this.Close();
        }
    }
}
