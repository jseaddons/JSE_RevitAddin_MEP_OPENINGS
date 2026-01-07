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
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// ✅ NEW UI V2: Two-panel layout (Left: Prefixes, Right: Parameter Mapping)
    /// Skeleton version - functionality to be added later
    /// </summary>
    public partial class ParameterServiceDialogV2 : WinForms.Form
    {
        // ✅ Public property to expose Active View Only state
        public bool IsActiveViewOnly => true; // Always true (Session Context enforced)

        private readonly Document? _document;
        private readonly UIDocument? _uiDocument;

        // Two-panel containers
        private WinForms.Panel _leftPrefixPanel = null!;
        private WinForms.Panel _rightMappingPanel = null!;

        // Left Panel: Prefix Configuration
        private WinForms.TextBox _projectPrefixTextBox = null!;
        private WinForms.Button _projectPrefixLockButton = null!;
        private WinForms.TextBox _ductPrefixTextBox = null!;
        private WinForms.TextBox _pipePrefixTextBox = null!;
        private WinForms.TextBox _cableTrayPrefixTextBox = null!;
        private WinForms.TextBox _damperPrefixTextBox = null!;
        private WinForms.Panel _systemTypeOverridesPanel = null!;
        private List<WinForms.Panel> _systemTypeRows = new List<WinForms.Panel>();
        private List<string> _systemTypeOptions = new List<string>();

        // ✅ NEW: Track last focused category for system type dropdown population
        private string _lastFocusedCategory = null;

        // Number format and remark checkboxes
        private WinForms.ComboBox _numberFormatCombo = null!;
        private WinForms.Button _numberFormatLockButton = null!;
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
        // REMOVED: private WinForms.Button _applyMarksButton = null!;
        private WinForms.Button _addPrefixButton = null!;
        private WinForms.Button _addNumberButton = null!;
        private WinForms.CheckBox _enableNumberingCheckBox = null!; // ✅ New Safety Checkbox
        private WinForms.Button _remarkSelectedButton = null!;
        private WinForms.Button _transferParametersButton = null!;
        private WinForms.Button _closeButton = null!;
        // private WinForms.CheckBox _activeViewOnlyCheckBox = null!; // REMOVED
        private WinForms.TextBox _startNumberTextBox = null!;

        // Reset Buttons (Safety Locked)
        private WinForms.Button _resetLockButton = null!;
        private WinForms.Button _resetNumberingButton = null!;
        private WinForms.Button _resetSelectionButton = null!;
        private WinForms.Button _resetParametersButton = null!;
        private bool _isResetUnlocked = false;

        public ParameterServiceDialogV2(Document document, UIDocument uiDocument)
        {
            _document = document;
            _uiDocument = uiDocument;

            InitializeComponent();
            
            // ✅ PERSISTENCE: Load saved settings on startup
            LoadSavedSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "Parameter Service V2";
            this.Size = new Size(1100, 700);
            this.MinimumSize = new Size(1000, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Create UI sections
            CreateTopBar();
            CreateTwoPanelLayout();
            CreatePrefixConfigurationPanel();
            CreateParameterMappingPanel();

            // ✅ PERSISTENCE: Load saved user inputs
            LoadSavedSettings();
        }

        /// <summary>
        /// Top Bar: Action Buttons
        /// </summary>
        private void CreateTopBar()
        {
            var topBar = new WinForms.Panel
            {
                Dock = DockStyle.Top,
                Height = 95,
                BackColor = Color.FromArgb(240, 248, 255)
            };
            this.Controls.Add(topBar);

            // ✅ LEFT-ALIGNED LAYOUT (User Request)
            int buttonWidth = 140;
            int buttonSpacing = 10;
            int buttonsStartX = 20; // Fixed Left Margin

            // --- ROW 1: Main Action Buttons (Y=9) ---

            // Button 1: Add Prefix Only (Slot 0)
            _addPrefixButton = new WinForms.Button
            {
                Text = "Add Prefix",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX, 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left, // Changed to Left
                BackColor = Color.FromArgb(70, 130, 180), // SteelBlue
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _addPrefixButton.Click += OnAddPrefixClick;
            topBar.Controls.Add(_addPrefixButton);

            // Button 2: Remark Selected (Slot 1)
            _remarkSelectedButton = new WinForms.Button
            {
                Text = "Remark Selected",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(150, 100, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _remarkSelectedButton.Click += OnRemarkSelectedClick;
            topBar.Controls.Add(_remarkSelectedButton);

            // Button 3: Add Number (Slot 2)
            _addNumberButton = new WinForms.Button
            {
                Text = "Add Number",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(60, 179, 113), // MediumSeaGreen
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _addNumberButton.Click += OnAddNumberClick;
            _addNumberButton.Enabled = false; // Disabled by default
            _addNumberButton.BackColor = Color.Gray; 
            topBar.Controls.Add(_addNumberButton);

            // Button 4: Transfer Parameters (Slot 3)
            _transferParametersButton = new WinForms.Button
            {
                Text = "Transfer Parameters",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(100, 150, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _transferParametersButton.Click += OnTransferParametersClick;
            topBar.Controls.Add(_transferParametersButton);

            // Button 5: Close (Slot 4 - Now Visible!)
            _closeButton = new WinForms.Button
            {
                Text = "✕ Close",
                Size = new Size(buttonWidth, 32),
                // Fix: Slot 4
                Location = new Point(buttonsStartX + 4 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(200, 100, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _closeButton.Click += (s, e) => this.Close();
            topBar.Controls.Add(_closeButton);


            // --- ROW 2: Reset Buttons & Checkboxes (Y=50) ---
            int row2Y = 50;

            // Reset Lock (Left of buttons)
            _resetLockButton = new WinForms.Button
            {
                Text = "🔒",
                Size = new Size(25, 22),
                Location = new Point(buttonsStartX, row2Y + 5), // Aligned with start
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.LightGreen,
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Emoji", 12F)
            };
            _resetLockButton.Click += (s, e) => {
                _isResetUnlocked = !_isResetUnlocked;
                bool unlocked = _isResetUnlocked;

                _resetLockButton.Text = unlocked ? "🔓" : "🔒";
                _resetLockButton.BackColor = unlocked ? Color.IndianRed : Color.LightGreen;

                // Toggle reset buttons
                if (_resetSelectionButton != null) {
                   _resetSelectionButton.Enabled = unlocked;
                   _resetSelectionButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;
                }
                if (_resetNumberingButton != null) {
                   _resetNumberingButton.Enabled = unlocked;
                   _resetNumberingButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;
                }
                if (_resetParametersButton != null) {
                   _resetParametersButton.Enabled = unlocked;
                   _resetParametersButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;
                }
            };
            topBar.Controls.Add(_resetLockButton);

            // Active View Only Checkbox Removed - Session Context is enforced automatically
            // _activeViewOnlyCheckBox = ...


            // Reset Selection (Below Slot 1 - Remark Selected)
            _resetSelectionButton = new WinForms.Button
            {
                Text = "Reset Selection",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), row2Y),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.LightGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                Enabled = false
            };
            _resetSelectionButton.Click += OnResetSelectionClick;
            topBar.Controls.Add(_resetSelectionButton);


            // Reset Numbering (Below Slot 2 - Add Number)
            _resetNumberingButton = new WinForms.Button
            {
                Text = "Reset Numbering",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), row2Y),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.LightGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                Enabled = false
            };
            _resetNumberingButton.Click += OnResetNumberingClick;
            topBar.Controls.Add(_resetNumberingButton);

            // Activate Numbering Checkbox (Moved to Slot 4 - Below Close Button)
            _enableNumberingCheckBox = new WinForms.CheckBox
            {
                Text = "Activate Numbering",
                Location = new Point(buttonsStartX + 4 * (buttonWidth + buttonSpacing), row2Y + 7), // Slot 4 (Below Close)
                Size = new Size(140, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                Checked = false,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray
            };
             // Add Tooltip for explanation
            System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();
            toolTip.SetToolTip(_enableNumberingCheckBox, "Enable numbering only after prefixes have been applied.");
            
            _enableNumberingCheckBox.CheckedChanged += (s, e) => 
            {
                bool enabled = _enableNumberingCheckBox.Checked;

                // ✅ VALIDATION: Numbering activation allowed only in Session Context (Floor Plan)
                if (enabled && !(_document.ActiveView is ViewPlan))
                {
                     WinForms.MessageBox.Show(
                        "Activate Numbering is only allowed in Floor Plan Views.\n\nPlease switch to a valid view to define the numbering session context.", 
                        "Invalid View Context", 
                        WinForms.MessageBoxButtons.OK, 
                        WinForms.MessageBoxIcon.Warning);
                     _enableNumberingCheckBox.Checked = false;
                     return;
                }

                _addNumberButton.Enabled = enabled;
                _addNumberButton.BackColor = enabled ? Color.FromArgb(60, 179, 113) : Color.Gray;
            };
            topBar.Controls.Add(_enableNumberingCheckBox);


            // Reset Parameters (Below Slot 3 - Transfer)
            _resetParametersButton = new WinForms.Button
            {
                Text = "Reset Parameters",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), row2Y),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.LightGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                Enabled = false 
            };
            _resetParametersButton.Click += OnResetParametersClick;
            topBar.Controls.Add(_resetParametersButton);
        }
            /* DUPLICATE REMOVAL
            // --- ROW 1: Main Action Buttons ---

            // Button 1: Add Prefix Only (First Slot)
            _addPrefixButton = new WinForms.Button
            {
                Text = "Add Prefix",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX, 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(70, 130, 180), // SteelBlue
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _addPrefixButton.Click += OnAddPrefixClick;
            topBar.Controls.Add(_addPrefixButton);

            // Button 2: Remark Selected (Moved to Second Slot)
            _remarkSelectedButton = new WinForms.Button
            {
                Text = "Remark Selected",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), 9), // Slot 1
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(150, 100, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _remarkSelectedButton.Click += OnRemarkSelectedClick;
            topBar.Controls.Add(_remarkSelectedButton);

            // Button 3: Add Number Only (Moved to Third Slot)
            _addNumberButton = new WinForms.Button
            {
                Text = "Add Number",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), 9), // Slot 2
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(60, 179, 113), // MediumSeaGreen
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _addNumberButton.Click += OnAddNumberClick;
            _addNumberButton.Enabled = false; // ✅ Disabled by default
            _addNumberButton.BackColor = Color.Gray; 
            topBar.Controls.Add(_addNumberButton);

            // ✅ Activate Numbering Checkbox (Safety) - Moved to Slot 2
            _enableNumberingCheckBox = new WinForms.CheckBox
            {
                Text = "Activate Numbering",
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), 45), // Below Button 3
                Size = new Size(buttonWidth, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Checked = false,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Regular),
                ForeColor = Color.DarkSlateGray
            };
            // Add Tooltip for explanation
            System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();
            toolTip.SetToolTip(_enableNumberingCheckBox, "Enable numbering only after prefixes have been applied.");
            
            _enableNumberingCheckBox.CheckedChanged += (s, e) => 
            {
                bool enabled = _enableNumberingCheckBox.Checked;
                _addNumberButton.Enabled = enabled;
                _addNumberButton.BackColor = enabled ? Color.FromArgb(60, 179, 113) : Color.Gray;
            };
            // ✅ REMOVED: Moved Activate Numbering checkbox to replace Reset Remarks button below
            // topBar.Controls.Add(_enableNumberingCheckBox);



            // Button 4: Transfer Parameters (Fourth Slot)
            _transferParametersButton = new WinForms.Button
            {
                Text = "Transfer Parameters",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(100, 150, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _transferParametersButton.Click += OnTransferParametersClick;
            topBar.Controls.Add(_transferParametersButton);


            // --- ROW 2: Reset Buttons (Below corresponding actions) ---
            int row2Y = 50;

            // --- SAFETY LOCK ICON (Small, next to Reset row) ---
            _resetLockButton = new WinForms.Button
            {
                Text = "🔒",
                Size = new Size(25, 22),
                Location = new Point(buttonsStartX - 30, row2Y + 5), // Left of buttons
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.LightGreen,
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI Emoji", 12F)
            };
            _resetLockButton.Click += (s, e) => {
                _isResetUnlocked = !_isResetUnlocked;
                bool unlocked = _isResetUnlocked;

                _resetLockButton.Text = unlocked ? "🔓" : "🔒";
                _resetLockButton.BackColor = unlocked ? Color.IndianRed : Color.LightGreen;

                _resetNumberingButton.Enabled = unlocked;
                _resetNumberingButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;

                _resetSelectionButton.Enabled = unlocked;
                _resetSelectionButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;

                _resetParametersButton.Enabled = unlocked;
                _resetParametersButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;
            };
            topBar.Controls.Add(_resetLockButton);


            // Reset Numbering (Below Add Number? Or Add Prefix?)
            // Align Reset Numbering below Add Number (Slot 1) makes sense?
            // Or Add Prefix (Slot 0)?
            // Reset Numbering clears numbers. So maybe below Add Number.
            // But let's keep it below Slot 0/1 area.

            // ✅ MOVED: Activate Numbering checkbox (replaces Reset Remarks button)
            _enableNumberingCheckBox = new WinForms.CheckBox
            {
                Text = "Activate Numbering",
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), row2Y + 5), // Slot 1 (where Reset Remarks was)
                Size = new Size(buttonWidth, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Checked = false,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray
            };
            System.Windows.Forms.ToolTip toolTip2 = new System.Windows.Forms.ToolTip();
            toolTip2.SetToolTip(_enableNumberingCheckBox, "Enable numbering only after prefixes have been applied.");
            
            _enableNumberingCheckBox.CheckedChanged += (s, e) => 
            {
                bool enabled = _enableNumberingCheckBox.Checked;
                _addNumberButton.Enabled = enabled;
                _addNumberButton.BackColor = enabled ? Color.FromArgb(60, 179, 113) : Color.Gray;
            };
            topBar.Controls.Add(_enableNumberingCheckBox);
            
            // Reset Numbering (Below Add Number - Slot 2)
            _resetNumberingButton = new WinForms.Button
            {
                Text = "Reset Numbering",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), row2Y), // Slot 2
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.LightGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                Enabled = false // Default Locked
            };
            _resetNumberingButton.Click += OnResetNumberingClick;
            topBar.Controls.Add(_resetNumberingButton);

            // Reset Parameters (Below Transfer Parameters - Slot 3)
            _resetParametersButton = new WinForms.Button
            {
                Text = "Reset Parameters",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), row2Y), // Slot 3
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.LightGray,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold),
                Enabled = false // Default Locked
            };
            _resetParametersButton.Click += OnResetParametersClick;
            topBar.Controls.Add(_resetParametersButton);




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

            // ✅ RESTORED: Active View Only Checkbox
            _activeViewOnlyCheckBox = new WinForms.CheckBox
            {
                Text = "Active View Only",
                Location = new Point(buttonsStartX - 130, 15), // Position to the left of buttons
                Size = new Size(120, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular),
                Checked = true // Default to true for safety/performance
            };
            topBar.Controls.Add(_activeViewOnlyCheckBox);
       */

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
                Size = new Size(120, 22),
                Text = "", // ✅ FIX: Default project prefix is blank (no default needed)
                Enabled = false,
                BackColor = Color.LightGray
            };
            _leftPrefixPanel.Controls.Add(_projectPrefixTextBox);

            _projectPrefixLockButton = new WinForms.Button
            {
                Location = new Point(245, yPos - 2),
                Size = new Size(25, 22),
                Text = "🔒",
                Font = new Font("Segoe UI Emoji", 8F),
                BackColor = Color.LightGreen
            };
            _projectPrefixLockButton.Click += (s, e) => ToggleLock(_projectPrefixLockButton, _projectPrefixTextBox);
            _leftPrefixPanel.Controls.Add(_projectPrefixLockButton);

            // Remark checkbox for Project Prefix
            _remarkProjectCheckBox = new WinForms.CheckBox
            {
                Text = "Remark",
                Location = new Point(280, yPos - 2),
                AutoSize = true,
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
                Size = new Size(120, 22),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false,
                BackColor = Color.LightGray
            };
            _numberFormatCombo.Items.AddRange(new[] { "00 (01, 02...)", "000 (001, 002...)", "0000 (0001, 0002...)" });
            _numberFormatCombo.SelectedIndex = 1; // Default to "000"
            _leftPrefixPanel.Controls.Add(_numberFormatCombo);

            _numberFormatLockButton = new WinForms.Button
            {
                Location = new Point(245, yPos - 2),
                Size = new Size(25, 22),
                Text = "🔒",
                Font = new Font("Segoe UI Emoji", 8F),
                BackColor = Color.LightGreen
            };
            _numberFormatLockButton.Click += (s, e) => ToggleLock(_numberFormatLockButton, _numberFormatCombo);
            _leftPrefixPanel.Controls.Add(_numberFormatLockButton);
            yPos += 35;

            // Start Number
            var startNumberLabel = new WinForms.Label
            {
                Text = "Start Number:",
                Location = new Point(10, yPos),
                Size = new Size(100, 20)
            };
            _leftPrefixPanel.Controls.Add(startNumberLabel);

            _startNumberTextBox = new WinForms.TextBox
            {
                Location = new Point(120, yPos - 2),
                Size = new Size(120, 22),
                Text = "1" // Default to 1
            };
            _leftPrefixPanel.Controls.Add(_startNumberTextBox);
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
            _ductPrefixTextBox.GotFocus += (s, e) => { _lastFocusedCategory = "Ducts"; };
            _leftPrefixPanel.Controls.Add(_ductPrefixTextBox);
            _remarkDuctCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = true };
            _leftPrefixPanel.Controls.Add(_remarkDuctCheckBox);
            yPos += 30;

            // Row 2: Pipe
            var pipeLabel = new WinForms.Label { Text = "Pipe:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(pipeLabel);
            _pipePrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "P" };
            _leftPrefixPanel.Controls.Add(_pipePrefixTextBox);
            _remarkPipeCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = true };
            _leftPrefixPanel.Controls.Add(_remarkPipeCheckBox);
            yPos += 30;

            // Row 3: Cable Tray
            var cableTrayLabel = new WinForms.Label { Text = "Cable Tray:", Location = new Point(10, yPos), Size = new Size(70, 20) };
            _leftPrefixPanel.Controls.Add(cableTrayLabel);
            _cableTrayPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "E" };
            _leftPrefixPanel.Controls.Add(_cableTrayPrefixTextBox);
            _remarkCableTrayCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = true };
            _leftPrefixPanel.Controls.Add(_remarkCableTrayCheckBox);
            yPos += 30;

            // Row 4: Damper
            var damperLabel = new WinForms.Label { Text = "Damper:", Location = new Point(10, yPos), Size = new Size(80, 20) };
            _leftPrefixPanel.Controls.Add(damperLabel);
            _damperPrefixTextBox = new WinForms.TextBox { Location = new Point(90, yPos - 2), Size = new Size(50, 22), Text = "DMP" };
            _leftPrefixPanel.Controls.Add(_damperPrefixTextBox);
            _remarkDamperCheckBox = new WinForms.CheckBox { Text = "Remark", Location = new Point(150, yPos - 2), Size = new Size(70, 22), Checked = true };
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

            // ✅ CRITICAL FIX: Load system types from ALL placed sleeves initially (no category filter)
            // This ensures we get system types from all categories that have placed sleeves
            LoadSystemTypeOverrideOptions(null); // Explicitly pass null to load all

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

            // MEP Parameter Dropdown
            var mepCombo = new WinForms.ComboBox
            {
                Location = new Point(10, 2),
                Size = new Size(row.Width / 2 - 60, 22), // Adjusted for lookup button
                DropDownStyle = ComboBoxStyle.DropDownList,
                Tag = "mep"
            };
            var mepParams = GetMepParametersForCategory(category);
            mepCombo.Items.AddRange(mepParams);
            if (mepCombo.Items.Contains(mepParam))
                mepCombo.SelectedItem = mepParam;
            row.Controls.Add(mepCombo);

            // 🔍 MEP Search button (Greenish)
            var mepSearchBtn = new WinForms.Button
            {
                Text = "🔍",
                Location = new Point(mepCombo.Right + 2, 2),
                Size = new Size(24, 22),
                BackColor = Color.LightGreen,
                FlatStyle = FlatStyle.Flat
            };
            var mepSearchTooltip = new System.Windows.Forms.ToolTip();
            mepSearchTooltip.SetToolTip(mepSearchBtn, "Search MEP/Host parameters from linked files");
            mepSearchBtn.Click += (s, e) => ShowMepParameterSearchDialog(mepCombo, category);
            row.Controls.Add(mepSearchBtn);

            // Arrow
            var arrow = new WinForms.Label
            {
                Text = "→",
                Location = new Point(mepSearchBtn.Right + 5, 4),
                Size = new Size(20, 20),
                Font = new Font("Arial", 10, FontStyle.Bold)
            };
            row.Controls.Add(arrow);

            // Opening Parameter Dropdown 
            var openingCombo = new WinForms.ComboBox
            {
                Location = new Point(arrow.Right + 5, 2),
                Size = new Size(row.Width / 2 - 60, 22),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Tag = "opening"
            };
            var openingParams = GetOpeningParametersForCategory(category);
            openingCombo.Items.AddRange(openingParams);
            if (openingCombo.Items.Contains(openingParam))
                openingCombo.SelectedItem = openingParam;
            row.Controls.Add(openingCombo);

            // 🔍 Opening Search button (Yellowish)
            var openingSearchBtn = new WinForms.Button
            {
                Text = "🔍",
                Location = new Point(openingCombo.Right + 2, 2),
                Size = new Size(24, 22),
                BackColor = Color.Khaki,
                FlatStyle = FlatStyle.Flat
            };
            var openingSearchTooltip = new System.Windows.Forms.ToolTip();
            openingSearchTooltip.SetToolTip(openingSearchBtn, "Search Opening family instance parameters");
            openingSearchBtn.Click += (s, e) => ShowOpeningParameterSearchDialog(openingCombo, category);
            row.Controls.Add(openingSearchBtn);

            // Row Resize Logic to keep layout tight
            row.Resize += (s, e) => {
                int half = (row.Width - 140) / 2;
                mepCombo.Width = half;
                mepSearchBtn.Location = new Point(mepCombo.Right + 2, 2);
                arrow.Location = new Point(mepSearchBtn.Right + 5, 4);
                openingCombo.Location = new Point(arrow.Right + 5, 2);
                openingCombo.Width = half;
                openingSearchBtn.Location = new Point(openingCombo.Right + 2, 2);
            };

            // Remove button
            var removeBtn = new WinForms.Button
            {
                Text = "⨉",
                Location = new Point(row.Width - 30, 2),
                Size = new Size(24, 22),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.Red
            };
            removeBtn.Click += (s, e) => {
                panel.Controls.Remove(row);
                RepositionRows(panel);
            };
            row.Controls.Add(removeBtn);

            panel.Controls.Add(row);
            RepositionRows(panel);
        }

        private void ShowMepParameterSearchDialog(WinForms.ComboBox comboBox, string category)
        {
            // ✅ FIX: Get ALL parameters for search, scanning linked files/host elements
            var allParams = GetAllMepParametersForCategory(category);
            ShowSearchDialog(comboBox, allParams, "Search MEP/Host Parameters");
        }

        /// <summary>
        /// Get ALL MEP/Host parameters for search dialog
        /// Scans linked files and current document to find ALL available parameters for the category
        /// </summary>
        private string[] GetAllMepParametersForCategory(string category)
        {
            if (_document == null) return new string[0];

            var allParams = new HashSet<string>();

            try
            {
                // Map category to BuiltInCategory
                BuiltInCategory? targetCategory = category switch
                {
                    "Ducts" => BuiltInCategory.OST_DuctCurves,
                    "Pipes" => BuiltInCategory.OST_PipeCurves,
                    "Cable Trays" => BuiltInCategory.OST_CableTray,
                    "Duct Accessories" => BuiltInCategory.OST_DuctAccessory,
                    "Walls" => BuiltInCategory.OST_Walls,
                    "Floors" => BuiltInCategory.OST_Floors,
                    "Structural Framing" => BuiltInCategory.OST_StructuralFraming,
                    _ => null
                };

                if (targetCategory == null) return new string[0];

                // Determine if we should look in linked files or current document
                // MEP elements usually in linked files (Reference Elements tab) or current doc
                // Host elements usually in linked files (Host Elements tab)

                var linkedFileService = new Services.LinkedFileService();
                var allLinkedFiles = linkedFileService.GetLinkedFiles(_document);

                // For search, we look in BOTH current document AND all linked files to be comprehensive
                // 1. Scan Current Document
                ScanDocumentForParameters(_document, targetCategory.Value, allParams);

                // 2. Scan Linked Files
                foreach (var link in allLinkedFiles)
                {
                    if (link.LinkInstance != null)
                    {
                        var linkDoc = link.LinkInstance.GetLinkDocument();
                        if (linkDoc != null)
                        {
                            ScanDocumentForParameters(linkDoc, targetCategory.Value, allParams);
                        }
                    }
                }

                return allParams.OrderBy(p => p).ToArray();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error getting all MEP parameters: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string paramLogPath = SafeFileLogger.GetLogFilePath("parameter_service_debug.log");
                    System.IO.File.AppendAllText(paramLogPath, $"[{DateTime.Now}] ERROR getting all MEP params: {ex.Message}\n");
                }
                // Fallback to startup essentials if search fails
                return GetMepParametersForCategory(category);
            }
        }

        private void ScanDocumentForParameters(Document doc, BuiltInCategory category, HashSet<string> paramsSet)
        {
            try
            {
                // Get first few elements to sample parameters
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .Take(5); // Sample 5 elements to get a good mix of instance parameters

                foreach (Element elem in collector)
                {
                    foreach (Parameter param in elem.Parameters)
                    {
                        if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                        {
                            paramsSet.Add(param.Definition.Name);
                        }
                    }

                    // Also check Type parameters
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        Element typeElem = doc.GetElement(typeId);
                        if (typeElem != null)
                        {
                            foreach (Parameter param in typeElem.Parameters)
                            {
                                if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                                {
                                    paramsSet.Add(param.Definition.Name);
                                }
                            }
                        }
                    }
                }
            }
            catch { /* Ignore read errors in individual docs */ }
        }

        private void ShowOpeningParameterSearchDialog(WinForms.ComboBox comboBox, string category)
        {
            // ✅ FIX: Get ALL parameters for search, not just the essential ones
            var allParams = GetAllOpeningParametersForCategory(category);
            ShowSearchDialog(comboBox, allParams, "Search Opening Parameters");
        }

        /// <summary>
        /// Get ALL parameters from opening families for the search dialog
        /// Scans all parameters in all opening families found in the document
        /// </summary>
        private string[] GetAllOpeningParametersForCategory(string category)
        {
            if (_document == null) return new string[0];

            try
            {
                var openingFamilies = GetOpeningFamilies(_document);
                var allParams = new HashSet<string>();

                foreach (var familySymbol in openingFamilies)
                {
                    // 1. Get Type Parameters
                    foreach (Parameter param in familySymbol.Parameters)
                    {
                        if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                        {
                            allParams.Add(param.Definition.Name);
                        }
                    }

                    // 2. Get Instance Parameters (try to find one instance)
                    // We need to check if there are any instances of this symbol placed in the model
                    try
                    {
                        var filter = new FamilyInstanceFilter(_document, familySymbol.Id);
                        var instance = new FilteredElementCollector(_document)
                            .WherePasses(filter)
                            .FirstOrDefault();

                        if (instance != null)
                        {
                            foreach (Parameter param in instance.Parameters)
                            {
                                if (param.Definition != null && !string.IsNullOrEmpty(param.Definition.Name))
                                {
                                    allParams.Add(param.Definition.Name);
                                }
                            }
                        }
                    }
                    catch (Exception instEx)
                    {
                        DebugLogger.Warning($"[ParameterServiceDialogV2] Error getting instance params for {familySymbol.Name}: {instEx.Message}");
                    }
                }

                return allParams.OrderBy(p => p).ToArray();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error getting all opening parameters: {ex.Message}");
                return new string[0];
            }
        }

        private void ShowSearchDialog(WinForms.ComboBox comboBox, string[] allParams, string title)
        {
            var searchForm = new WinForms.Form
            {
                Text = title,
                Size = new Size(400, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var searchBox = new WinForms.TextBox { Location = new Point(10, 10), Width = 365 };
            var listBox = new WinForms.ListBox { Location = new Point(10, 40), Width = 365, Height = 380 };

            listBox.Items.AddRange(allParams);

            searchBox.TextChanged += (s, e) => {
                listBox.Items.Clear();
                var filtered = allParams.Where(p => p.IndexOf(searchBox.Text, StringComparison.OrdinalIgnoreCase) >= 0).ToArray();
                listBox.Items.AddRange(filtered);
            };

            listBox.DoubleClick += (s, e) => {
                if (listBox.SelectedItem != null)
                {
                    var val = listBox.SelectedItem.ToString();
                    if (!comboBox.Items.Contains(val)) comboBox.Items.Add(val);
                    comboBox.SelectedItem = val;
                    searchForm.Close();
                }
            };

            var selectBtn = new WinForms.Button { Text = "Select", Location = new Point(300, 430), DialogResult = DialogResult.OK };
            selectBtn.Click += (s, e) => {
                if (listBox.SelectedItem != null)
                {
                    var val = listBox.SelectedItem.ToString();
                    if (!comboBox.Items.Contains(val)) comboBox.Items.Add(val);
                    comboBox.SelectedItem = val;
                }
                searchForm.Close();
            };

            searchForm.Controls.Add(searchBox);
            searchForm.Controls.Add(listBox);
            searchForm.Controls.Add(selectBtn);
            searchForm.ShowDialog();
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
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };

            // ✅ CRITICAL FIX: Populate dropdown ONCE when row is created - don't reload on DropDown event
            // This prevents memory access violations from database/Revit API calls during dropdown opening
            // The dropdown will use the cached _systemTypeOptions that were loaded when dialog opened
            PopulateSystemTypeDropdown(systemTypeCombo, systemType);

            // ✅ REMOVED: DropDown event handler that was causing memory access violations
            // Instead, system types are loaded once when dialog opens and cached in _systemTypeOptions
            // If user needs fresh data, they can close and reopen the dialog
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
                Checked = true // Default to Checked as per user request
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
        /// Load available System Type / Service Type values from persisted XML (and future data sources).
        /// Populates <see cref="_systemTypeOptions"/> for use by override rows.
        /// ✅ TESTING MODE: Always loads ALL system types and service types from ALL categories (ignores category filter)
        /// </summary>
        private void LoadSystemTypeOverrideOptions(string category = null)
        {
            try
            {
                var collected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (_document != null)
                {
                    var catalogService = new SystemTypeCatalogService(msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info(msg);
                        }
                        SafeFileLogger.SafeAppendText("parameter_service_debug.log", $"[{DateTime.Now}] {msg}\n");
                    });

                    // ✅ TESTING MODE: Always load ALL types from ALL categories (ignore category parameter)
                    var catalog = catalogService.Load(_document, category: null);

                    // ✅ LOAD ALL: Include both SystemTypes and ServiceTypes from all categories
                    foreach (var value in catalog.SystemTypes ?? Enumerable.Empty<string>())
                    {
                        if (!string.IsNullOrWhiteSpace(value))
                            collected.Add(value.Trim());
                    }
                    foreach (var value in catalog.ServiceTypes ?? Enumerable.Empty<string>())
                    {
                        if (!string.IsNullOrWhiteSpace(value))
                            collected.Add(value.Trim());
                    }
                }

                _systemTypeOptions = collected
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Select(v => v.Trim())
                    .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // ✅ ENHANCED LOGGING: Log what will be populated in the dropdown
                var categoryInfo = string.IsNullOrWhiteSpace(category) ? "all categories" : $"category '{category}'";
                if (_systemTypeOptions.Count > 0)
                {
                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                        $"[{DateTime.Now}] [ParameterServiceDialogV2] ✅ System Type dropdown will be populated with {_systemTypeOptions.Count} options for {categoryInfo}: {string.Join(", ", _systemTypeOptions)}\n");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                        $"[{DateTime.Now}] [ParameterServiceDialogV2] ⚠️ System Type dropdown will be EMPTY for {categoryInfo} - no system/service types found in placed sleeves\n");
                }

                if (_systemTypeOptions.Count == 0)
                {
                    _systemTypeOptions = new List<string>();
                }
            }
            catch (Exception ex)
            {
                _systemTypeOptions = new List<string>();
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[ParameterServiceDialogV2] Failed to load system type catalog: {ex.Message}");
                }
                SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                    $"[{DateTime.Now}] [ParameterServiceDialogV2] Failed to load system type catalog: {ex.Message}\n");
            }
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
        /// ✅ NEW: Populate system type dropdown with available options
        /// </summary>
        private void PopulateSystemTypeDropdown(WinForms.ComboBox comboBox, string selectedValue = null)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add("<Select>");

            // ✅ CRITICAL FIX: Only show system types from placed sleeves - NO hardcoded fallback
            // If no system types found, dropdown will only show "<Select>"
            var options = _systemTypeOptions ?? new List<string>();

            foreach (var option in options)
            {
                if (!string.IsNullOrWhiteSpace(option))
                {
                    comboBox.Items.Add(option);
                }
            }

            if (!string.IsNullOrWhiteSpace(selectedValue) && comboBox.Items.Contains(selectedValue))
            {
                comboBox.SelectedItem = selectedValue;
            }
            else
            {
                comboBox.Text = string.IsNullOrWhiteSpace(selectedValue) ? "<Select>" : selectedValue;
            }
        }

        /// <summary>
        /// ✅ NEW: Detect current category based on which discipline prefix textbox was last focused/edited
        /// Returns "Ducts", "Pipes", "Cable Trays", or null if cannot determine
        /// </summary>
        private string DetectCurrentCategory()
        {
            // ✅ CRITICAL FIX: First check which textbox currently has focus
            if (_ductPrefixTextBox != null && _ductPrefixTextBox.Focused)
                return "Ducts";
            if (_pipePrefixTextBox != null && _pipePrefixTextBox.Focused)
                return "Pipes";
            if (_cableTrayPrefixTextBox != null && _cableTrayPrefixTextBox.Focused)
                return "Cable Trays";
            if (_damperPrefixTextBox != null && _damperPrefixTextBox.Focused)
                return "Duct Accessories";

            // ✅ ENHANCEMENT: If no textbox has focus, use the last focused category
            // This handles the case when user clicks dropdown after focusing a textbox
            if (!string.IsNullOrWhiteSpace(_lastFocusedCategory))
            {
                SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                    $"[{DateTime.Now}] [ParameterServiceDialogV2] Using last focused category: '{_lastFocusedCategory}'\n");
                return _lastFocusedCategory;
            }

            // If we can't determine, return null to load all types
            SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                $"[{DateTime.Now}] [ParameterServiceDialogV2] ⚠️ Cannot determine category - loading all system/service types\n");
            return null;
        }

        private void ToggleLock(WinForms.Button lockBtn, WinForms.TextBox textBox)
        {
            if (lockBtn.Text == "🔒")
            {
                lockBtn.Text = "🔓";
                lockBtn.BackColor = Color.LightCoral;
                textBox.Enabled = true;
                textBox.BackColor = Color.White;
            }
            else
            {
                lockBtn.Text = "🔒";
                lockBtn.BackColor = Color.LightGreen;
                textBox.Enabled = false;
                textBox.BackColor = Color.LightGray;
            }
        }

        private void ToggleLock(WinForms.Button lockBtn, WinForms.ComboBox comboBox)
        {
            if (lockBtn.Text == "🔒")
            {
                lockBtn.Text = "🔓";
                lockBtn.BackColor = Color.LightCoral;
                comboBox.Enabled = true;
                comboBox.BackColor = Color.White;
            }
            else
            {
                lockBtn.Text = "🔒";
                lockBtn.BackColor = Color.LightGreen;
                comboBox.Enabled = false;
                comboBox.BackColor = Color.LightGray;
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

                    // ✅ PERFORMANCE MONITORING: Track parameter transfer performance
                    using (var perfMonitor = new ParameterOperationPerformanceMonitor("Parameter Transfer"))
                    {
                        // Create parameter transfer service
                        var transferService = new ParameterTransferService();

                        // ✅ OPTIMIZATION: Filter by Active View OR Section Box
                        // Filter sleeves at collection time to reduce processing set
                        FilteredElementCollector collector;
                        ElementFilter sectionBoxFilter = null;

                        // 2. GLOBAL COLLECTION (with optional Section Box filter)
                        // Active View Only mode removed from Transfer logic (defaults to Global/Section Box)
                        collector = new FilteredElementCollector(_document)
                                .OfClass(typeof(FamilyInstance));

                            // Section Box Filter (only applied if not using Active View filter)
                            if (Services.OptimizationFlags.UseSectionBoxFilterForParameterTransfer && _uiDocument != null)
                            {
                                try
                                {
                                    if (_uiDocument.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                                    {
                                        var sectionBoxBounds = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                                        if (sectionBoxBounds != null)
                                        {
                                            var sectionBoxOutline = new Outline(sectionBoxBounds.Min, sectionBoxBounds.Max);
                                            sectionBoxFilter = new BoundingBoxIntersectsFilter(sectionBoxOutline);
                                            collector = (FilteredElementCollector)collector.WherePasses(sectionBoxFilter);
                                            DebugLogger.Info($"[ParameterServiceDialogV2] ✅ Section box filtering applied during collection");
                                        }
                                    }
                                }
                                catch (Exception sectionBoxEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[ParameterServiceDialogV2] ⚠️ Section box filter setup failed: {sectionBoxEx.Message}");
                                }
                            }


                        // Now get all openings (individual + cluster) in the document, pre-filtered
                        // Only match the 4 specific opening families: RectangularOpeningOnWall, RectangularOpeningOnSlab, CircularOpeningOnWall, CircularOpeningOnSlab
                        var sleevesToProcess = collector
                            .Cast<FamilyInstance>()
                            .Where(fi => {
                                var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                                // Match only the 4 specific opening families
                                return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                                    || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                            })
                            .ToList();

                        // Log family names for debugging
                        var familyNames = sleevesToProcess
                            .Select(fi => fi.Symbol?.Family?.Name ?? "Unknown")
                            .Distinct()
                            .ToList();
                        DebugLogger.Info($"[ParameterServiceDialogV2] Found opening families: {string.Join(", ", familyNames)}");

                        if (sectionBoxFilter != null && !DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ParameterServiceDialogV2] ✅ Section box filtering applied during collection: {sleevesToProcess.Count} sleeves within section box");
                        }

                        var openings = sleevesToProcess.Select(fi => fi.Id).ToList();

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

                        // ✅ CRITICAL FIX: Add source parameter names to learned keys whitelist
                        // This ensures these parameters are captured during the next refresh
                        foreach (var mapping in allMappings)
                        {
                            if (!string.IsNullOrWhiteSpace(mapping.SourceParameter))
                            {
                                Services.ParameterSnapshotService.AddLearnedKey(mapping.SourceParameter);
                                DebugLogger.Info($"[ParameterServiceDialogV2] Added '{mapping.SourceParameter}' to learned parameter keys whitelist");
                            }
                        }

                        // Create configuration with all mappings
                        var config = new Models.ParameterTransferConfiguration
                        {
                            SourceCategoryName = "All", // Transfer from all categories
                            Mappings = allMappings
                        };

                        // ✅ FIX: Use ExecuteTransferConfigurationInTransaction directly to ensure snapshots are loaded
                        // This method loads snapshots from database (saved during placement) and uses them for transfer
                        // The 3-arg ExecuteTransferConfiguration also calls this internally, but calling directly ensures
                        // we're using the snapshot-based transfer path
                        ParameterTransferResult result;
                        using (var transaction = new Transaction(_document, "Transfer Parameters to Sleeves"))
                        {
                            transaction.Start();
                            result = transferService.ExecuteTransferConfigurationInTransaction(_document, openings, config, _uiDocument);
                            transaction.Commit();
                        }

                        progressForm.Close();

                        // Set item count for performance monitoring
                        perfMonitor.SetItemCount(result.TransferredCount);

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
                    } // End of performance monitor using block
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

                // ✅ VALIDATION: Numbering/Apply Marks ONLY allowed in Floor Plan with Crop Box or Scope Box (Session Context)
                if (!(_document.ActiveView is ViewPlan plan) || !(plan.CropBoxActive || (plan.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)?.AsElementId() ?? ElementId.InvalidElementId) != ElementId.InvalidElementId))
                {
                    WinForms.MessageBox.Show(
                        "Numbering is only allowed in Floor Plan Views with an active Crop Box or Scope Box applied.\n\nPlease switch to a valid view to define the numbering session context.", 
                        "Invalid View Context", 
                        WinForms.MessageBoxButtons.OK, 
                        WinForms.MessageBoxIcon.Warning);
                    return;
                }

                // ✅ DATABASE-BASED: Check if categories have clash zone data in the database
                // Apply Marks needs clash zone data to determine which category each sleeve belongs to
                var missingCategories = GetCategoriesWithoutData();
                if (missingCategories.Count > 0)
                {
                    var categoryList = string.Join("\n• ", missingCategories);
                    var result = WinForms.MessageBox.Show(
                        $"⚠️ CLASH DATA NOT FOUND\n\n" +
                        $"The following categories cannot be processed because no clash zone data is found in the database:\n\n" +
                        $"• {categoryList}\n\n" +
                        $"Apply Marks requires clash zone data in the database to determine sleeve categories.\n\n" +
                        $"Please run Refresh to detect clash zones before applying marks.\n\n" +
                        $"Would you like to continue with available categories only?",
                        "Missing Clash Data",
                        WinForms.MessageBoxButtons.YesNo,
                        WinForms.MessageBoxIcon.Warning);

                    if (result != WinForms.DialogResult.Yes)
                    {
                        return; // User cancelled
                    }

                     // Γ£à WARNING SUPPRESSED: Log warning instead of showing popup
                     RemarkDebugLogger.LogStep($"[OnRemarkSelectedClick] Warning: Some categories missing clash data ({string.Join(", ", missingCategories)}). Continuing with available data as per user preference (Warning Box Suppressed).");
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

                    int startNum = 1;
                    int.TryParse(_startNumberTextBox.Text, out startNum);

                    var markPrefixes = new Models.MarkPrefixSettings
                    {
                        ProjectPrefix = projectPrefix,
                        DuctPrefix = ductPrefix,
                        PipePrefix = pipePrefix,
                        CableTrayPrefix = cableTrayPrefix,
                        DamperPrefix = damperPrefix,
                        NumberFormat = numberFormat,
                        // ✅ OPTIMIZATION: Always use Active View context (validated above)
                        ActiveViewOnly = true, 
                        StartNumber = startNum
                    };

                    // Sync remark checkboxes so MarkParameterCommand honours the user's selection
                    markPrefixes.RemarkAll = remarkProject;
                    markPrefixes.RemarkProjectPrefix = remarkProject;
                    markPrefixes.RemarkDuctPrefix = remarkDuct;
                    markPrefixes.RemarkPipePrefix = remarkPipe;
                    markPrefixes.RemarkCableTrayPrefix = remarkCableTray;
                    markPrefixes.RemarkDamperPrefix = remarkDamper;

                    // Collect checked System Type Overrides
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

                    // Persist any per-system overrides that were checked
                    foreach (var (systemType, prefix) in systemTypeOverrides)
                    {
                        if (!string.IsNullOrWhiteSpace(systemType))
                        {
                            markPrefixes.DuctSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                            markPrefixes.PipeSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                            markPrefixes.DuctAccessoriesSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                        }
                    }

                    // Debug: Log the prefix values being used
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Apply Marks - Project: '{projectPrefix}', Duct: '{ductPrefix}', Pipe: '{pipePrefix}', CableTray: '{cableTrayPrefix}', Damper: '{damperPrefix}', Format: '{numberFormat}', Start: {startNum}\n");

                    // ✅ PERFORMANCE MONITORING: Track Apply Marks performance
                    int totalProcessed = 0;
                    using (var perfMonitor = new ParameterOperationPerformanceMonitor("Apply Marks"))
                    {
                        if (markPrefixes.ActiveViewOnly && _document != null)
                        {
                            // ✅ BIM 360 OPTIMIZATION: Per-sheet numbering using database-driven logic
                            var markService = new Services.MarkParameterService(null, msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); });

                            // ✅ PARALLEL OPTIMIZATION: Process combined sleeves in parallel with individual categories
                            System.Threading.Tasks.Task<(int processed, int errors)>? combinedTask = null;
                            combinedTask = System.Threading.Tasks.Task.Run(() =>
                            {
                                using (var context = new Data.SleeveDbContext(_document))
                                {
                                    var repo = new Data.Repositories.ClashZoneRepository(context, msg => { });
                                    var levelName = (_document.ActiveView as ViewPlan)?.GenLevel?.Name ?? "";

                                    // Check if combined sleeves exist in current session (IsCurrentClash=1)
                                    var combinedCount = repo.GetSleevesForLevel(levelName, "Combined")
                                        .Count(z => z.IsCurrentClash && z.CombinedClusterSleeveInstanceId > 0);

                                    if (combinedCount > 0)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[ParameterServiceDialogV2] Found {combinedCount} combined sleeves in session - processing with MEP prefix");

                                        // Process combined sleeves in parallel
                                        return markService.ApplyMarksFromDatabase(_document, markPrefixes, "Combined");
                                    }

                                    return (0, 0); // No combined sleeves
                                }
                            });

                            // Pass 1: Individual Disciplines (runs in parallel with combined task)
                            var disciplineCategories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                            foreach (var cat in disciplineCategories)
                            {
                                // Only process categories that have their remark checkbox enabled
                                if (!markPrefixes.GetRemarkFlag(cat)) continue;

                                var (processed, errors) = markService.ApplyMarksFromDatabase(_document, markPrefixes, cat);
                                totalProcessed += processed;
                            }

                            // Pass 2: Wait for combined sleeves to complete
                            var (combinedProcessed, combinedErrors) = combinedTask.Result;
                            totalProcessed += combinedProcessed;
                        }
                        else
                        {
                            // Legacy logic: Mark ALL sleeves regardless of checkbox state (using old command)
                            var cmd = new MarkParameterCommand("ALL", projectPrefix, "", false, markPrefixes);
                            cmd.Execute(_uiDocument.Application);

                            // Get total sleeves count for performance monitoring
                            totalProcessed = new FilteredElementCollector(_document)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .Where(fi => {
                                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                                })
                                .Count();
                        }

                        perfMonitor.SetItemCount(totalProcessed);
                    }

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
            // ✅ PERSISTENCE: Save settings before processing
            SaveCurrentSettings();

            RemarkDebugLogger.LogStep("--- Remark Selected Clicked ---");
            try
            {
                if (_document == null || _uiDocument == null)
                {
                    RemarkDebugLogger.LogError("Document or UIDocument is NULL");
                    WinForms.MessageBox.Show("Document not available.", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                    return;
                }

                // Collect remark checkbox states
                var remarkProject = _remarkProjectCheckBox.Checked;
                var remarkDuct = _remarkDuctCheckBox.Checked;
                var remarkPipe = _remarkPipeCheckBox.Checked;
                var remarkCableTray = _remarkCableTrayCheckBox.Checked;
                var remarkDamper = _remarkDamperCheckBox.Checked;

                RemarkDebugLogger.LogSelection("RemarkProject", remarkProject);
                RemarkDebugLogger.LogSelection("RemarkDuct", remarkDuct);
                RemarkDebugLogger.LogSelection("RemarkPipe", remarkPipe);
                RemarkDebugLogger.LogSelection("RemarkCableTray", remarkCableTray);
                RemarkDebugLogger.LogSelection("RemarkDamper", remarkDamper);

                // Collect prefix values
                var projectPrefix = _projectPrefixTextBox.Text.Trim();
                var ductPrefix = _ductPrefixTextBox.Text.Trim();
                var pipePrefix = _pipePrefixTextBox.Text.Trim();
                var cableTrayPrefix = _cableTrayPrefixTextBox.Text.Trim();
                var damperPrefix = _damperPrefixTextBox.Text.Trim();

                RemarkDebugLogger.LogSelection("ProjectPrefix", projectPrefix);
                RemarkDebugLogger.LogSelection("DuctPrefix", ductPrefix);
                RemarkDebugLogger.LogSelection("PipePrefix", pipePrefix);
                RemarkDebugLogger.LogSelection("CableTrayPrefix", cableTrayPrefix);
                RemarkDebugLogger.LogSelection("DamperPrefix", damperPrefix);

                var numberFormat = _numberFormatCombo.SelectedIndex switch
                {
                    0 => "00",
                    1 => "000",
                    2 => "0000",
                    _ => "000"
                };
                RemarkDebugLogger.LogSelection("NumberFormat", numberFormat);

                // ✅ CRITICAL FIX: Collect System/Service Type Overrides and detect which category they belong to
                var ductSystemTypeOverrides = new List<(string systemType, string prefix)>();
                var pipeSystemTypeOverrides = new List<(string systemType, string prefix)>();
                var ductAccessoriesSystemTypeOverrides = new List<(string systemType, string prefix)>();
                var cableTrayServiceTypeOverrides = new List<(string serviceType, string prefix)>();
                bool hasDuctSystemOverrideRemark = false;
                bool hasPipeSystemOverrideRemark = false;
                bool hasDuctAccessoriesSystemOverrideRemark = false;
                bool hasCableTrayServiceOverrideRemark = false;

                // ✅ STEP 1: Get available system/service types for each category to detect which category they belong to
                var cableTrayServiceTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var ductSystemTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var pipeSystemTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var ductAccessoriesSystemTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                try
                {
                    if (_document != null)
                    {
                        var catalogService = new SystemTypeCatalogService(msg => { });

                        // Get Cable Tray Service Types
                        var cableTrayCatalog = catalogService.Load(_document, "Cable Trays");
                        foreach (var st in cableTrayCatalog.ServiceTypes ?? Enumerable.Empty<string>())
                        {
                            if (!string.IsNullOrWhiteSpace(st))
                                cableTrayServiceTypes.Add(st.Trim());
                        }

                        // Get Duct System Types
                        var ductCatalog = catalogService.Load(_document, "Ducts");
                        foreach (var st in ductCatalog.SystemTypes ?? Enumerable.Empty<string>())
                        {
                            if (!string.IsNullOrWhiteSpace(st))
                                ductSystemTypes.Add(st.Trim());
                        }

                        // Get Pipe System Types
                        var pipeCatalog = catalogService.Load(_document, "Pipes");
                        foreach (var st in pipeCatalog.SystemTypes ?? Enumerable.Empty<string>())
                        {
                            if (!string.IsNullOrWhiteSpace(st))
                                pipeSystemTypes.Add(st.Trim());
                        }

                        // Get Duct Accessories System Types
                        var ductAccessoriesCatalog = catalogService.Load(_document, "Duct Accessories");
                        foreach (var st in ductAccessoriesCatalog.SystemTypes ?? Enumerable.Empty<string>())
                        {
                            if (!string.IsNullOrWhiteSpace(st))
                                ductAccessoriesSystemTypes.Add(st.Trim());
                        }
                    }
                }
                catch { /* Ignore errors */ }

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
                            var systemOrServiceType = comboBox.Text;
                            var prefix = textBox.Text;

                            // ✅ CRITICAL: Detect which category this system/service type belongs to
                            // Priority: Cable Tray Service Types > Duct System Types > Pipe System Types > Duct Accessories System Types
                            if (cableTrayServiceTypes.Contains(systemOrServiceType))
                            {
                                // This is a Cable Tray Service Type
                                cableTrayServiceTypeOverrides.Add((systemOrServiceType, prefix));
                                hasCableTrayServiceOverrideRemark = true;
                            }
                            else if (ductSystemTypes.Contains(systemOrServiceType))
                            {
                                // This is a Duct System Type
                                ductSystemTypeOverrides.Add((systemOrServiceType, prefix));
                                hasDuctSystemOverrideRemark = true;
                            }
                            else if (pipeSystemTypes.Contains(systemOrServiceType))
                            {
                                // This is a Pipe System Type
                                pipeSystemTypeOverrides.Add((systemOrServiceType, prefix));
                                hasPipeSystemOverrideRemark = true;
                            }
                            else if (ductAccessoriesSystemTypes.Contains(systemOrServiceType))
                            {
                                // This is a Duct Accessories System Type
                                ductAccessoriesSystemTypeOverrides.Add((systemOrServiceType, prefix));
                                hasDuctAccessoriesSystemOverrideRemark = true;
                            }
                            else
                            {
                                // Unknown - default to Duct System Type (backward compatibility)
                                ductSystemTypeOverrides.Add((systemOrServiceType, prefix));
                                hasDuctSystemOverrideRemark = true;
                            }
                        }
                    }
                }

                // ✅ NEW: Check if filter XML files exist for any category we plan to remark
                var categoriesNeedingFilters = new List<string>();
                if (remarkProject || remarkDuct || hasDuctSystemOverrideRemark) categoriesNeedingFilters.Add("Ducts");
                if (remarkProject || remarkPipe || hasPipeSystemOverrideRemark) categoriesNeedingFilters.Add("Pipes");
                if (remarkProject || remarkCableTray || hasCableTrayServiceOverrideRemark) categoriesNeedingFilters.Add("Cable Trays");
                if (remarkProject || remarkDamper || hasDuctAccessoriesSystemOverrideRemark) categoriesNeedingFilters.Add("Duct Accessories");

                if (categoriesNeedingFilters.Count > 0)
                {
                    RemarkDebugLogger.LogInfo($"Categories needing filters: {string.Join(", ", categoriesNeedingFilters)}");
                    var missingCategories = GetCategoriesWithoutData()
                        .Where(cat => categoriesNeedingFilters.Contains(cat))
                        .ToList();

                    if (missingCategories.Count > 0)
                    {
                        RemarkDebugLogger.LogStep($"WARNING: Missing clash data for: {string.Join(", ", missingCategories)}");
                        // ✅ SUPPRESSED: User requested to continue without warning dialog
                        // Continue to process what we can
                        RemarkDebugLogger.LogStep("Continuing with available categories (Warning suppressed via user request)");
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

                        // Sync remark checkboxes so MarkParameterCommand honours the user's selection
                        markPrefixes.RemarkAll = remarkProject;
                        markPrefixes.RemarkProjectPrefix = remarkProject;
                        markPrefixes.RemarkDuctPrefix = remarkDuct || hasDuctSystemOverrideRemark;
                        markPrefixes.RemarkPipePrefix = remarkPipe || hasPipeSystemOverrideRemark;
                        markPrefixes.RemarkCableTrayPrefix = remarkCableTray || hasCableTrayServiceOverrideRemark;
                        markPrefixes.RemarkDamperPrefix = remarkDamper || hasDuctAccessoriesSystemOverrideRemark;

                        // ✅ CRITICAL FIX: Add system/service type overrides to correct dictionaries based on category
                        // Duct System Type Overrides
                        foreach (var (systemType, prefix) in ductSystemTypeOverrides)
                        {
                            if (!string.IsNullOrWhiteSpace(systemType))
                            {
                                markPrefixes.DuctSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                            }
                        }

                        // Pipe System Type Overrides
                        foreach (var (systemType, prefix) in pipeSystemTypeOverrides)
                        {
                            if (!string.IsNullOrWhiteSpace(systemType))
                            {
                                markPrefixes.PipeSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                            }
                        }

                        // Duct Accessories System Type Overrides
                        foreach (var (systemType, prefix) in ductAccessoriesSystemTypeOverrides)
                        {
                            if (!string.IsNullOrWhiteSpace(systemType))
                            {
                                markPrefixes.DuctAccessoriesSystemTypeOverrides[systemType] = prefix ?? string.Empty;
                            }
                        }

                        // Cable Tray Service Type Overrides
                        foreach (var (serviceType, prefix) in cableTrayServiceTypeOverrides)
                        {
                            if (!string.IsNullOrWhiteSpace(serviceType))
                            {
                                markPrefixes.CableTrayServiceTypeOverrides[serviceType] = prefix ?? string.Empty;
                            }
                        }

                        // ✅ PERFORMANCE MONITORING: Track Remark Selected performance
                        int totalProcessed = 0;
                        var categoriesProcessed = new List<string>();

                        using (var perfMonitor = new ParameterOperationPerformanceMonitor("Remark Selected"))
                        {
                            // ✅ CRITICAL FIX: Use markPrefixes.GetRemarkFlag() instead of hardcoded true
                            // This ensures that each category's checkbox state is respected

                            // If Project Prefix remark is checked, re-mark ALL categories with new project prefix
                            // ✅ CRITICAL FIX: Use PrefixOnly mode for Remark Selected
                            // This aligns with the new separated workflow where prefixes and numbers are handled independently.
                            RemarkDebugLogger.LogStep("Executing MarkParameterCommand in SELECTED/PrefixOnly mode");
                            var cmd = new MarkParameterCommand(
                                "SELECTED",
                                projectPrefix,
                                "",
                                false,
                                markPrefixes,
                                MarkParameterCommand.MarkingMode.PrefixOnly
                            );
                            cmd.Execute(_uiDocument.Application);

                            // Populate summary for completion message
                            if (remarkProject)
                            {
                                categoriesProcessed.Add("All Categories (Project Prefix)");
                            }
                            else
                            {
                                if (markPrefixes.RemarkDuctPrefix)
                                    categoriesProcessed.Add(hasDuctSystemOverrideRemark && !remarkDuct ? "Ducts (System Type Overrides)" : "Ducts");
                                if (markPrefixes.RemarkPipePrefix)
                                    categoriesProcessed.Add(hasPipeSystemOverrideRemark && !remarkPipe ? "Pipes (System Type Overrides)" : "Pipes");
                                if (markPrefixes.RemarkCableTrayPrefix)
                                    categoriesProcessed.Add(hasCableTrayServiceOverrideRemark && !remarkCableTray ? "Cable Trays (Service Type Overrides)" : "Cable Trays");
                                if (markPrefixes.RemarkDamperPrefix)
                                    categoriesProcessed.Add(hasDuctAccessoriesSystemOverrideRemark && !remarkDamper ? "Duct Accessories (System Type Overrides)" : "Duct Accessories");
                            }
                            RemarkDebugLogger.LogInfo($"Categories processed by UI logic: {string.Join(", ", categoriesProcessed)}");

                            totalProcessed = categoriesProcessed.Count > 0 ? 1 : 0;

                            // Get total sleeves count for performance monitoring
                            var allSleeves = new FilteredElementCollector(_document)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .Where(fi =>
                                {
                                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                                })
                                .Count();
                            perfMonitor.SetItemCount(allSleeves);
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
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error during remark: {ex.Message}",
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ✅ DATABASE-BASED: Check if categories have clash zone data in the database
        /// Returns list of categories that don't have clash data in the database
        /// Replaces the old XML file-based check
        /// </summary>
        private List<string> GetCategoriesWithoutData()
        {
            var missingCategories = new List<string>();

            try
            {
                if (_document == null) return missingCategories;

                // ✅ DATABASE MIGRATION: Check database instead of XML files
                var dbContext = new Data.SleeveDbContext(_document);
                var clashZoneRepo = new Data.Repositories.ClashZoneRepository(dbContext, null);

                // Check each MEP category in the database
                var categoriesToCheck = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };

                foreach (var category in categoriesToCheck)
                {
                    // Query database to check if category has clash zones
                    var clashZones = clashZoneRepo.GetClashZonesByCategory(category);

                    if (clashZones == null || clashZones.Count == 0)
                    {
                        // Also check SleeveSnapshots table as fallback
                        // (some categories might have data in snapshots but not in ClashZones table)
                        try
                        {
                            using (var cmd = dbContext.Connection.CreateCommand())
                            {
                                // Check if there are any snapshots for this category
                                // We can check by MEP parameters JSON or by checking if any zones reference this category
                                cmd.CommandText = @"
                                    SELECT COUNT(*) FROM SleeveSnapshots 
                                    WHERE MepParametersJson LIKE @categoryPattern
                                    LIMIT 1";
                                cmd.Parameters.AddWithValue("@categoryPattern", $"%{category}%");

                                var snapshotCount = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);

                                // Also check ClashZones table one more time with direct SQL
                                cmd.Parameters.Clear();
                                cmd.CommandText = @"
                                    SELECT COUNT(*) FROM ClashZones 
                                    WHERE MepCategory = @category
                                    LIMIT 1";
                                cmd.Parameters.AddWithValue("@category", category);

                                var clashZoneCount = Convert.ToInt32(cmd.ExecuteScalar() ?? 0);

                                if (snapshotCount == 0 && clashZoneCount == 0)
                                {
                                    missingCategories.Add(category);
                                }
                            }
                        }
                        catch (Exception dbEx)
                        {
                            // If database query fails, assume category is missing
                            DebugLogger.Warning($"[ParameterServiceDialogV2] Error checking database for category {category}: {dbEx.Message}");
                            missingCategories.Add(category);
                        }
                    }
                }

                DebugLogger.Info($"[ParameterServiceDialogV2] Database check: {missingCategories.Count} categories missing clash data");
                if (missingCategories.Count > 0)
                {
                    DebugLogger.Warning($"[ParameterServiceDialogV2] Missing clash data in database for: {string.Join(", ", missingCategories)}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceDialogV2] Error checking database for categories: {ex.Message}");
                // On error, don't block processing - return empty list (allow processing to continue)
            }

            return missingCategories;
        }

        /// <summary>
        /// ✅ DEPRECATED: Old XML-based check (kept for fallback only)
        /// Use GetCategoriesWithoutData() instead
        /// </summary>
        [Obsolete("Use GetCategoriesWithoutData() instead - checks database instead of XML files")]
        private List<string> CheckFilterFilesForCategories()
        {
            // ✅ FALLBACK: Try database first, then XML if database check fails
            try
            {
                var dbMissing = GetCategoriesWithoutData();
                if (dbMissing.Count < 4) // If we found at least some categories in DB, use DB result
                {
                    return dbMissing;
                }
            }
            catch
            {
                // If database check fails, fall back to XML check below
            }

            // Fallback to XML file check (for backward compatibility)
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
                                        if (filter?.ClashZoneStorage?.AllZones != null && filter.ClashZoneStorage.AllZones.Count > 0)
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

                DebugLogger.Info($"[ParameterServiceDialogV2] Filter file check (fallback): {missingCategories.Count} categories missing filter data");
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
        /// <summary>
        /// ✅ RESET NUMBERING: Clears Marks and Resets Counters
        /// </summary>
        private void OnResetNumberingClick(object sender, EventArgs e)
        {
            if (_document == null) return;
            bool activeViewOnly = true; // Hardcoded (Safety: Default to View Context)
            // bool activeViewOnly = _activeViewOnlyCheckBox?.Checked ?? false; (Removed)

            string scopeMsg = activeViewOnly ? "ACTIVE VIEW ONLY" : "ENTIRE PROJECT";
            if (WinForms.MessageBox.Show(
                $"Are you sure you want to RESET NUMBERING?\n\nScope: {scopeMsg}\n\n" +
                "This will clear the 'MEP Mark' from sleeves and reset the internal counters to 0.\n" +
                "The next time you apply marks, they will start from '1' (or your Start Number).",
                "CONFIRM RESET NUMBERING",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Warning) != WinForms.DialogResult.Yes)
            {
                return;
            }

            try
            {
                var markService = new MarkParameterService(_document);

                // Get View Extent if Active View Only
                BoundingBoxXYZ viewExtent = null;
                string levelName = "ALL";
                if (activeViewOnly && _document.ActiveView != null)
                {
                    viewExtent = _document.ActiveView.CropBox;
                    levelName = _document.ActiveView.GenLevel?.Name ?? "ActiveView";
                }

                int clearedCount = markService.ResetMarksForLevel(_document, levelName, viewExtent);

                // Reset Counters (Global)
                if (!activeViewOnly)
                {
                    // Only reset global counters if we are not in "Active View" mode
                    // OR should we reset them anyway? 
                    // User said "Reset Numbering", which implies resetting the sequence key.
                    // But if I only clear 5 sleeves in a view, should the counter reset?
                    // Safe approach: Reset counters if they asked for it. 
                    // Actually, if they reset "Active View", they might want to re-number those 5.
                    // But duplicates might occur if other views have 6-100.
                    // Re-read user request: "reset the numbering or delete apply marks"
                    // I will reset counters.
                    markService.ResetCategoryCounters(_document); // Reset all category counters
                }
                else
                {
                    // If active view only, we simply clear marks. Resetting global counter creates collision risk.
                    // I will Warn about counters not being reset in partial clear?
                    // No, let's keep it simple. Clears marks.
                    // User can manually reset counters by unchecking active view? 
                    // Let's reset counters anyway, assuming they want to "start over".
                    markService.ResetCategoryCounters(_document); // Reset all category counters
                }

                WinForms.MessageBox.Show($"Successfully cleared Marks from {clearedCount} sleeves.\nCounters have been reset.", "Reset Complete");
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error resetting numbering: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// ✅ RESET SELECTION: Clears Marks from Selected Elements ONLY
        /// </summary>
        private void OnResetSelectionClick(object sender, EventArgs e)
        {
            if (_document == null || _uiDocument == null) return;

            var selectedIds = _uiDocument.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                WinForms.MessageBox.Show("Please select at least one sleeve element.", "No Selection");
                return;
            }

            if (WinForms.MessageBox.Show(
                $"Are you sure you want to clear 'MEP Mark' from the {selectedIds.Count} selected elements?\n\n" +
                "Note: This does NOT reset the global numbering counters.",
                "CONFIRM RESET SELECTION",
                WinForms.MessageBoxButtons.YesNo,
                WinForms.MessageBoxIcon.Question) != WinForms.DialogResult.Yes)
            {
                return;
            }

            try
            {
                var markService = new MarkParameterService(_document);
                int clearedCount = markService.ResetMarksForSelection(_document, selectedIds);

                WinForms.MessageBox.Show($"Cleared {clearedCount} marks.", "Success");
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error clearing selection: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// ✅ RESET PARAMETERS: Clears Transferred Parameters
        /// Strictly Session Sensitive (Active View Only) if checked
        /// </summary>
        private void OnResetParametersClick(object sender, EventArgs e)
        {
            if (_document == null) return;
            bool activeViewOnly = true; // Hardcoded (Safety: Default to View Context)
            // bool activeViewOnly = _activeViewOnlyCheckBox?.Checked ?? false; (Removed)

            if (!activeViewOnly)
            {
                if (WinForms.MessageBox.Show(
                    "WARNING: 'Active View Only' is invalid (unchecked).\n" +
                    "Resetting parameters for the ENTIRE PROJECT is dangerous.\n\n" +
                    "Are you absolutely sure you want to clear parameters for ALL sleeves?",
                    "DANGER: RESET ALL PARAMETERS",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Stop) != WinForms.DialogResult.Yes)
                {
                    return;
                }
            }
            else
            {
                if (WinForms.MessageBox.Show(
                    "Are you sure you want to reset parameters for sleeves in the ACTIVE VIEW?",
                    "Confirm Reset Parameters",
                    WinForms.MessageBoxButtons.YesNo,
                    WinForms.MessageBoxIcon.Warning) != WinForms.DialogResult.Yes)
                {
                    return;
                }
            }

            try
            {
                // 1. Collect IDs based on Scope
                var repo = new ClashZoneRepository(new SleeveDbContext(_document));
                var allSleeves = repo.GetAllSleeves(); // Baseline
                var idsToReset = new List<ElementId>();

                if (activeViewOnly && _document.ActiveView != null)
                {
                    // Use same logic as Apply Marks to filter by view
                    BoundingBoxXYZ viewExtent = _document.ActiveView.CropBox;
                    Transform viewTransform = viewExtent.Transform;

                    // Get coords
                    XYZ bMin = viewExtent.Min;
                    XYZ bMax = viewExtent.Max;
                    var corners = new List<XYZ>
                    {
                        viewTransform.OfPoint(new XYZ(bMin.X, bMin.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMax.X, bMin.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMax.X, bMax.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMin.X, bMax.Y, bMin.Z))
                    };
                    double worldMinX = corners.Min(c => c.X) - 0.01;
                    double worldMinY = corners.Min(c => c.Y) - 0.01;
                    double worldMaxX = corners.Max(c => c.X) + 0.01;
                    double worldMaxY = corners.Max(c => c.Y) + 0.01;

                    var filtered = allSleeves.Where(z =>
                    {
                        double pX = z.SleevePlacementPointActiveDocumentX;
                        double pY = z.SleevePlacementPointActiveDocumentY;
                        // Fallback logic
                        if (Math.Abs(pX) < 0.001) pX = z.SleevePlacementPointX;
                        if (Math.Abs(pY) < 0.001) pY = z.SleevePlacementPointY;

                        return pX >= worldMinX && pX <= worldMaxX &&
                               pY >= worldMinY && pY <= worldMaxY;
                    }).ToList();

                    foreach (var z in filtered) idsToReset.Add(new ElementId(z.SleeveInstanceId));
                }
                else
                {
                    // All Sleeves
                    foreach (var z in allSleeves) idsToReset.Add(new ElementId(z.SleeveInstanceId));
                }

                // 2. Call Service
                var svc = new ParameterTransferService();
                int count = svc.ResetParameters(_document, idsToReset);

                WinForms.MessageBox.Show($"Reset parameters for {count} sleeves.", "Reset Complete");
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error resetting parameters: {ex.Message}", "Error");
            }
        }
        /// <summary>
        /// Handle Add Prefix button click
        /// Applies ONLY prefixes (DB-based) to marked items
        /// </summary>
        private void OnAddPrefixClick(object sender, EventArgs e)
        {
            // ✅ PERSISTENCE: Save user inputs before applying
            SaveCurrentSettings();

            var (settings, isValid) = GetMarkPrefixSettingsFromUI();
            if (!isValid) return;

            try
            {
                // Show mini progress
                using (var progressForm = new WinForms.Form { Text = "Applying Prefixes...", Size = new Size(300, 50), StartPosition = FormStartPosition.CenterScreen })
                {
                    progressForm.Show();
                    progressForm.Refresh();

                    // Execute Command in PrefixOnly mode
                    var cmd = new MarkParameterCommand(
                        "ALL",
                        settings.ProjectPrefix,
                        "", // Discipline prefix handled by settings
                        settings.RemarkAll,
                        settings,
                        MarkParameterCommand.MarkingMode.PrefixOnly
                    );

                    cmd.Execute(_uiDocument.Application);

                    progressForm.Close();
                    WinForms.MessageBox.Show("Prefixes applied successfully.", "Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error applying prefixes: {ex.Message}", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Handle Add Number button click
        /// Applies ONLY numbers (Revit scan) to marked items
        /// </summary>
        private void OnAddNumberClick(object sender, EventArgs e)
        {
            // ✅ PERSISTENCE: Save user inputs before applying
            SaveCurrentSettings();

            var (settings, isValid) = GetMarkPrefixSettingsFromUI(requireDbData: false);
            if (!isValid) return;

            try
            {
                using (var progressForm = new WinForms.Form { Text = "Applying Numbers...", Size = new Size(300, 50), StartPosition = FormStartPosition.CenterScreen })
                {
                    progressForm.Show();
                    progressForm.Refresh();

                    // Execute Command in NumberOnly mode
                    var cmd = new MarkParameterCommand(
                        "ALL",
                        settings.ProjectPrefix,
                        "",
                        settings.RemarkAll,
                        settings,
                        MarkParameterCommand.MarkingMode.NumberOnly
                    );

                    cmd.Execute(_uiDocument.Application);

                    progressForm.Close();
                    WinForms.MessageBox.Show("Numbers applied successfully.", "Complete", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error applying numbers: {ex.Message}", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Helper to gather UI settings and validate state
        /// </summary>
        private (Models.MarkPrefixSettings Settings, bool IsValid) GetMarkPrefixSettingsFromUI(bool requireDbData = true)
        {
            if (_document == null || _uiDocument == null)
            {
                WinForms.MessageBox.Show("Document not available.", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                return (null, false);
            }

            // ✅ REMOVED: Obsolete warning check - no longer needed with optimized database columns
            /*
            if (requireDbData)
            {
                var missingCategories = GetCategoriesWithoutData();
                if (missingCategories.Count > 0)
                {
                    var result = WinForms.MessageBox.Show(
                        $"⚠️ CLASH DATA MISSING\nSome categories missing data might delay processing.\nContinue?",
                        "Warning", WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Warning);

                    if (result != WinForms.DialogResult.Yes) return (null, false);
                }
            }
            */

            var projectPrefix = _projectPrefixTextBox.Text.Trim();
            var numberFormat = _numberFormatCombo.SelectedIndex switch
            {
                0 => "00",
                1 => "000",
                2 => "0000",
                _ => "000"
            };

            int startNum = 1;
            int.TryParse(_startNumberTextBox.Text, out startNum);

            var settings = new Models.MarkPrefixSettings
            {
                ProjectPrefix = projectPrefix,
                DuctPrefix = _ductPrefixTextBox.Text.Trim(),
                PipePrefix = _pipePrefixTextBox.Text.Trim(),
                CableTrayPrefix = _cableTrayPrefixTextBox.Text.Trim(),
                DamperPrefix = _damperPrefixTextBox.Text.Trim(),
                NumberFormat = numberFormat,
                ActiveViewOnly = true, // Hardcoded enforce session context
                // ActiveViewOnly = _activeViewOnlyCheckBox.Checked, (Removed)
                StartNumber = startNum,
                RemarkAll = _remarkProjectCheckBox.Checked, // Use project checkbox as global 'All' default
                RemarkProjectPrefix = _remarkProjectCheckBox.Checked,
                RemarkDuctPrefix = _remarkDuctCheckBox.Checked,
                RemarkPipePrefix = _remarkPipeCheckBox.Checked,
                RemarkCableTrayPrefix = _remarkCableTrayCheckBox.Checked,
                RemarkDamperPrefix = _remarkDamperCheckBox.Checked
            };

            // System Type Overrides
            foreach (var row in _systemTypeRows)
            {
                // ✅ LOGIC UPDATE: Checkbox is redundant. If data exists, use it.
                // var checkbox = row.Controls.OfType<WinForms.CheckBox>().FirstOrDefault();
                
                var comboBox = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault();
                var textBox = row.Controls.OfType<WinForms.TextBox>().FirstOrDefault();

                // If System Type is selected and Prefix is entered (or even if Prefix is empty? No, prefix usually needed. But maybe empty = empty prefix?)
                // Assuming user wants override if they selected a System Type.
                if (comboBox != null && textBox != null && !string.IsNullOrWhiteSpace(comboBox.Text))
                {
                    string sysType = comboBox.Text.Trim();
                    string prefix = textBox.Text.Trim();

                    settings.DuctSystemTypeOverrides[sysType] = prefix;
                    settings.PipeSystemTypeOverrides[sysType] = prefix;
                    settings.CableTrayServiceTypeOverrides[sysType] = prefix;
                    settings.DuctAccessoriesSystemTypeOverrides[sysType] = prefix;
                }
            }

            return (settings, true);
        }

        /// <summary>
        /// ✅ PERSISTENCE: Load saved user inputs from JSON file
        /// </summary>
        private void LoadSavedSettings()
        {
            RemarkDebugLogger.LogInfo("=== LoadSavedSettings CALLED (UI) ===");
            try
            {
                var savedSettings = SystemTypeOverridePersistenceService.LoadSettings();
                RemarkDebugLogger.LogInfo("Loaded settings from persistence service");

                // Restore project prefix
                if (_projectPrefixTextBox != null)
                    _projectPrefixTextBox.Text = savedSettings.ProjectPrefix ?? "";

                // Restore discipline prefixes
                if (_ductPrefixTextBox != null)
                    _ductPrefixTextBox.Text = savedSettings.DuctPrefix ?? "M";
                if (_pipePrefixTextBox != null)
                    _pipePrefixTextBox.Text = savedSettings.PipePrefix ?? "P";
                if (_cableTrayPrefixTextBox != null)
                    _cableTrayPrefixTextBox.Text = savedSettings.CableTrayPrefix ?? "E";
                if (_damperPrefixTextBox != null)
                    _damperPrefixTextBox.Text = savedSettings.DamperPrefix ?? "D";

                // Restore number format
                if (_numberFormatCombo != null)
                {
                    string format = savedSettings.NumberFormat ?? "000";
                    int index = format == "00" ? 0 : (format == "000" ? 1 : 2);
                    _numberFormatCombo.SelectedIndex = index;
                }

                // Restore start number
                if (_startNumberTextBox != null)
                    _startNumberTextBox.Text = savedSettings.StartNumber.ToString();

                // Restore remark checkboxes - DEFAULT TO TRUE (checked) if no saved value
                if (_remarkProjectCheckBox != null)
                    _remarkProjectCheckBox.Checked = savedSettings.RemarkProjectPrefix;
                if (_remarkDuctCheckBox != null)
                    _remarkDuctCheckBox.Checked = savedSettings.RemarkDuctPrefix;
                if (_remarkPipeCheckBox != null)
                    _remarkPipeCheckBox.Checked = savedSettings.RemarkPipePrefix;
                if (_remarkCableTrayCheckBox != null)
                    _remarkCableTrayCheckBox.Checked = savedSettings.RemarkCableTrayPrefix;
                if (_remarkDamperCheckBox != null)
                    _remarkDamperCheckBox.Checked = savedSettings.RemarkDamperPrefix;

                // ✅ PERSISTENCE: Restore System Type Overrides
                var allOverrides = new Dictionary<string, string>();
                
                if (savedSettings.DuctSystemTypeOverrides != null)
                    foreach(var kvp in savedSettings.DuctSystemTypeOverrides) allOverrides[kvp.Key] = kvp.Value;
                    
                if (savedSettings.PipeSystemTypeOverrides != null)
                    foreach(var kvp in savedSettings.PipeSystemTypeOverrides) allOverrides[kvp.Key] = kvp.Value;
                    
                if (savedSettings.CableTrayServiceTypeOverrides != null)
                    foreach(var kvp in savedSettings.CableTrayServiceTypeOverrides) allOverrides[kvp.Key] = kvp.Value;
                    
                if (savedSettings.DuctAccessoriesSystemTypeOverrides != null)
                    foreach(var kvp in savedSettings.DuctAccessoriesSystemTypeOverrides) allOverrides[kvp.Key] = kvp.Value;

                if (allOverrides.Count > 0)
                {
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST] Restoring {allOverrides.Count} system type overrides");
                    
                    // Clear existing default rows
                    foreach(var row in _systemTypeRows.ToList()) // ToList to avoid modification exception
                    {
                        if (_systemTypeOverridesPanel.Controls.Contains(row))
                        {
                            _systemTypeOverridesPanel.Controls.Remove(row);
                        }
                        row.Dispose();
                    }
                    _systemTypeRows.Clear();
                    
                    // Add saved rows
                    foreach(var kvp in allOverrides)
                    {
                        if (!string.IsNullOrWhiteSpace(kvp.Key))
                        {
                            AddSystemTypeRow(kvp.Key, kvp.Value);
                        }
                    }
                    
                    // Fill up to 4 rows if needed
                    while (_systemTypeRows.Count < 4)
                    {
                         AddSystemTypeRow("<Select>", "");
                    }
                }

                RemarkDebugLogger.LogInfo("[ParameterServiceDialogV2] Loaded saved settings successfully");
            }
            catch (Exception ex)
            {
                RemarkDebugLogger.LogError($"[ParameterServiceDialogV2] Error loading saved settings: {ex.Message}");
                RemarkDebugLogger.LogError($"[ParameterServiceDialogV2] Error loading saved settings: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ PERSISTENCE: Save current user inputs to JSON file
        /// </summary>
        private void SaveCurrentSettings()
        {
            RemarkDebugLogger.LogInfo("=== SaveCurrentSettings CALLED (UI) ===");
            try
            {
                var settings = new MarkPrefixSettings
                {
                    ProjectPrefix = _projectPrefixTextBox?.Text ?? "",
                    DuctPrefix = _ductPrefixTextBox?.Text ?? "M",
                    PipePrefix = _pipePrefixTextBox?.Text ?? "P",
                    CableTrayPrefix = _cableTrayPrefixTextBox?.Text ?? "E",
                    DamperPrefix = _damperPrefixTextBox?.Text ?? "D",
                    NumberFormat = GetNumberFormatFromCombo(),
                    StartNumber = int.TryParse(_startNumberTextBox?.Text, out int startNum) ? startNum : 1,
                    RemarkProjectPrefix = _remarkProjectCheckBox?.Checked ?? false,
                    RemarkDuctPrefix = _remarkDuctCheckBox?.Checked ?? true,
                    RemarkPipePrefix = _remarkPipeCheckBox?.Checked ?? true,
                    RemarkCableTrayPrefix = _remarkCableTrayCheckBox?.Checked ?? true,
                    RemarkDamperPrefix = _remarkDamperCheckBox?.Checked ?? true
                };

                // Collect system type overrides from UI rows
                if (_systemTypeRows != null)
                {
                    foreach (var row in _systemTypeRows)
                    {
                        var comboBox = row.Controls.OfType<WinForms.ComboBox>().FirstOrDefault();
                        var textBox = row.Controls.OfType<WinForms.TextBox>().FirstOrDefault();
                        if (comboBox != null && textBox != null && !string.IsNullOrWhiteSpace(comboBox.Text))
                        {
                            settings.DuctSystemTypeOverrides[comboBox.Text] = textBox.Text;
                            settings.PipeSystemTypeOverrides[comboBox.Text] = textBox.Text;
                            settings.CableTrayServiceTypeOverrides[comboBox.Text] = textBox.Text;
                            settings.DuctAccessoriesSystemTypeOverrides[comboBox.Text] = textBox.Text;
                        }
                    }
                }

                SystemTypeOverridePersistenceService.SaveSettings(settings);
                RemarkDebugLogger.LogInfo("[ParameterServiceDialogV2] Saved current settings successfully");
            }
            catch (Exception ex)
            {
                RemarkDebugLogger.LogError($"[ParameterServiceDialogV2] Error saving settings: {ex.Message}");
            }
        }

        private string GetNumberFormatFromCombo()
        {
            if (_numberFormatCombo == null) return "000";
            int index = _numberFormatCombo.SelectedIndex;
            return index == 0 ? "00" : (index == 1 ? "000" : "0000");
        }
    } }


