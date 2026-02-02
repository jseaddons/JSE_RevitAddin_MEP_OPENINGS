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
// using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;
// using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Γ£à NEW UI V2: Two-panel layout (Left: Prefixes, Right: Parameter Mapping)
    /// Skeleton version - functionality to be added later
    /// </summary>
    public partial class ParameterServiceDialogV2 : WinForms.Form
    {
        // Γ£à Public property to expose Active View Only state
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

        // Γ£à NEW: Track last focused category for system type dropdown population
        private string _lastFocusedCategory = null;

        // Number format
        private WinForms.ComboBox _numberFormatCombo = null!;
        private WinForms.Button _numberFormatLockButton = null!;

        // Right Panel: Parameter Mapping
        private WinForms.TabControl _masterParameterTabs = null!;
        private WinForms.TabControl _referenceParameterTabs = null!;
        private WinForms.TabControl _hostParameterTabs = null!;

        // Top bar
        // REMOVED: private WinForms.Button _applyMarksButton = null!;
        private WinForms.Button _addPrefixButton = null!;
        private WinForms.Button _addNumberButton = null!;
        private WinForms.CheckBox _enableNumberingCheckBox = null!; // Γ£à New Safety Checkbox
        private WinForms.Button _resetPrefixButton = null!;
        private WinForms.Button _transferParametersButton = null!;
        private WinForms.Button _closeButton = null!;
        // private WinForms.CheckBox _activeViewOnlyCheckBox = null!; // REMOVED
        private WinForms.TextBox _startNumberTextBox = null!;
        private WinForms.Panel _numberingSequencePanel = null!;
        private WinForms.RadioButton _modeNewSheetRadio = null!;
        private WinForms.RadioButton _modeContinueRadio = null!;
        private WinForms.ComboBox _sourceViewCombo = null!;

        // Reset Buttons (Safety Locked)
        private WinForms.Button _resetLockButton = null!;
        private WinForms.Button _resetNumberingButton = null!;
        private WinForms.Button _resetParametersButton = null!;
        private bool _isResetUnlocked = false;

        public ParameterServiceDialogV2(Document document, UIDocument uiDocument)
        {
            _document = document;
            _uiDocument = uiDocument;

            InitializeComponent();
            
            // Γ£à PERSISTENCE: Load saved settings on startup
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

            // Γ£à PERSISTENCE: Load saved user inputs
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

            // Γ£à LEFT-ALIGNED LAYOUT (User Request)
            int buttonWidth = 140;
            int buttonSpacing = 10;
            int buttonsStartX = 20; // Fixed Left Margin

            // --- LAYOUT CONSTANTS ---
            int row1Y = 9;
            int row2Y = 50;

            // --- ROW 1: Main Action Buttons (Y=9) ---

            // Button 1: Add Prefix Only (Slot 1 - Vertically above Reset Prefix)
            _addPrefixButton = new WinForms.Button
            {
                Text = "Add Prefix",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), row1Y), // Moved to Slot 1
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(70, 130, 180), // SteelBlue
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _addPrefixButton.Click += OnAddPrefixClick;
            topBar.Controls.Add(_addPrefixButton);

            // Button 2: Reset Prefix (Slot 1 Row 2 - See Row 2 Setup)
            _resetPrefixButton = new WinForms.Button
            {
                Text = "Reset Prefix",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 1 * (buttonWidth + buttonSpacing), row2Y), // Moved to Row 2, Slot 1
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.LightGray, // Initially Grey (Locked)
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Bold), // Smaller font like other reset buttons
                Enabled = false // Initially Disabled
            };
            _resetPrefixButton.Click += OnResetPrefixClick;
            topBar.Controls.Add(_resetPrefixButton);

            // Button 3: Add Number (Slot 2)
            _addNumberButton = new WinForms.Button
            {
                Text = "Add Number",
                Size = new Size(buttonWidth, 32),
                Location = new Point(buttonsStartX + 2 * (buttonWidth + buttonSpacing), row1Y),
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
                Location = new Point(buttonsStartX + 3 * (buttonWidth + buttonSpacing), row1Y),
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
                Text = "Γ£ò Close",
                Size = new Size(buttonWidth, 32),
                // Fix: Slot 4
                Location = new Point(buttonsStartX + 4 * (buttonWidth + buttonSpacing), row1Y),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.FromArgb(200, 100, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold)
            };
            _closeButton.Click += (s, e) => this.Close();
            topBar.Controls.Add(_closeButton);


            // --- ROW 2: Reset Buttons & Checkboxes (Y=50) ---

            // Reset Lock (Left of buttons)
            _resetLockButton = new WinForms.Button
            {
                Text = "≡ƒöÆ",
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

                _resetLockButton.Text = unlocked ? "≡ƒöô" : "≡ƒöÆ";
                _resetLockButton.BackColor = unlocked ? Color.IndianRed : Color.LightGreen;

                // Toggle reset buttons
                if (_resetPrefixButton != null) {
                   _resetPrefixButton.Enabled = unlocked;
                   _resetPrefixButton.BackColor = unlocked ? Color.IndianRed : Color.LightGray;
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

                // Γ£à VALIDATION: Numbering activation allowed only in Session Context (Floor Plan)
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
            _addNumberButton.Enabled = false; // Γ£à Disabled by default
            _addNumberButton.BackColor = Color.Gray; 
            topBar.Controls.Add(_addNumberButton);

            // Γ£à Activate Numbering Checkbox (Safety) - Moved to Slot 2
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
            // Γ£à REMOVED: Moved Activate Numbering checkbox to replace Reset Remarks button below
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
                Text = "≡ƒöÆ",
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

                _resetLockButton.Text = unlocked ? "≡ƒöô" : "≡ƒöÆ";
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

            // Γ£à MOVED: Activate Numbering checkbox (replaces Reset Remarks button)
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
                Text = "Γ£ò Close",
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

            // Γ£à RESTORED: Active View Only Checkbox
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
                Text = "", // Γ£à FIX: Default project prefix is blank (no default needed)
                Enabled = false,
                BackColor = Color.LightGray
            };
            _leftPrefixPanel.Controls.Add(_projectPrefixTextBox);

            _projectPrefixLockButton = new WinForms.Button
            {
                Location = new Point(245, yPos - 2),
                Size = new Size(25, 22),
                Text = "≡ƒöÆ",
                Font = new Font("Segoe UI Emoji", 8F),
                BackColor = Color.LightGreen
            };
            _projectPrefixLockButton.Click += (s, e) => ToggleLock(_projectPrefixLockButton, _projectPrefixTextBox);
            _leftPrefixPanel.Controls.Add(_projectPrefixLockButton);
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
                Text = "≡ƒöÆ",
                Font = new Font("Segoe UI Emoji", 8F),
                BackColor = Color.LightGreen
            };
            _numberFormatLockButton.Click += (s, e) => ToggleLock(_numberFormatLockButton, _numberFormatCombo);
            _leftPrefixPanel.Controls.Add(_numberFormatLockButton);
            yPos += 35;

            // Start Number - MOVED TO Numbering Sequence Section below
            // ...

            // Section Title: Discipline Prefixes
            var disciplineLabel = new WinForms.Label
            {
                Text = "Discipline Prefixes (2x2 Grid):",
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(180, 20),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            _leftPrefixPanel.Controls.Add(disciplineLabel);
            yPos += 25;

            // Row 1: Duct & Pipe
            var ductLabel = new WinForms.Label { Text = "Duct:", Location = new Point(10, yPos), Size = new Size(40, 20) };
            _leftPrefixPanel.Controls.Add(ductLabel);
            _ductPrefixTextBox = new WinForms.TextBox { Location = new Point(55, yPos - 2), Size = new Size(50, 22), Text = "M" };
            _ductPrefixTextBox.GotFocus += (s, e) => { _lastFocusedCategory = "Ducts"; };
            _leftPrefixPanel.Controls.Add(_ductPrefixTextBox);

            var pipeLabel = new WinForms.Label { Text = "Pipe:", Location = new Point(135, yPos), Size = new Size(40, 20) };
            _leftPrefixPanel.Controls.Add(pipeLabel);
            _pipePrefixTextBox = new WinForms.TextBox { Location = new Point(180, yPos - 2), Size = new Size(50, 22), Text = "P" };
            _pipePrefixTextBox.GotFocus += (s, e) => { _lastFocusedCategory = "Pipes"; };
            _leftPrefixPanel.Controls.Add(_pipePrefixTextBox);
            yPos += 30;

            // Row 2: Cable Tray & Damper
            var cableTrayLabel = new WinForms.Label { Text = "Tray:", Location = new Point(10, yPos), Size = new Size(40, 20) };
            _leftPrefixPanel.Controls.Add(cableTrayLabel);
            _cableTrayPrefixTextBox = new WinForms.TextBox { Location = new Point(55, yPos - 2), Size = new Size(50, 22), Text = "E" };
            _cableTrayPrefixTextBox.GotFocus += (s, e) => { _lastFocusedCategory = "Cable Trays"; };
            _leftPrefixPanel.Controls.Add(_cableTrayPrefixTextBox);

            var damperLabel = new WinForms.Label { Text = "Damp:", Location = new Point(135, yPos), Size = new Size(40, 20) };
            _leftPrefixPanel.Controls.Add(damperLabel);
            _damperPrefixTextBox = new WinForms.TextBox { Location = new Point(180, yPos - 2), Size = new Size(50, 22), Text = "DMP" };
            _damperPrefixTextBox.GotFocus += (s, e) => { _lastFocusedCategory = "Duct Accessories"; };
            _leftPrefixPanel.Controls.Add(_damperPrefixTextBox);
            yPos += 40;

            // Γ£à NEW: Numbering Sequence Management
            var numberingTitle = new WinForms.Label
            {
                Text = "Numbering Sequence Logic:",
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold),
                Location = new Point(10, yPos),
                Size = new Size(200, 20),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            _leftPrefixPanel.Controls.Add(numberingTitle);
            yPos += 25;

            // Container for Numbering Modes
            _numberingSequencePanel = new WinForms.Panel
            {
                Location = new Point(10, yPos),
                Size = new Size(_leftPrefixPanel.Width - 25, 80),
                BackColor = Color.FromArgb(245, 245, 245),
                BorderStyle = BorderStyle.FixedSingle
            };
            _leftPrefixPanel.Controls.Add(_numberingSequencePanel);
            
            // Mode 1: New Sheet (Start Number)
            _modeNewSheetRadio = new WinForms.RadioButton
            {
                Text = "Start Number New Sheet:",
                Location = new Point(5, 5),
                Size = new Size(160, 20),
                Checked = true,
                Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular)
            };
            _numberingSequencePanel.Controls.Add(_modeNewSheetRadio);

            _startNumberTextBox = new WinForms.TextBox
            {
                Location = new Point(170, 4),
                Size = new Size(60, 20),
                Text = "" 
            };
            _numberingSequencePanel.Controls.Add(_startNumberTextBox);

            // Mode 2: Continue from Floor Plan
            _modeContinueRadio = new WinForms.RadioButton
            {
                Text = "Continue from Floor Plan:",
                Location = new Point(5, 30),
                Size = new Size(160, 20),
                Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular)
            };
            _numberingSequencePanel.Controls.Add(_modeContinueRadio);

            _sourceViewCombo = new WinForms.ComboBox
            {
                Location = new Point(5, 52),
                Size = new Size(_numberingSequencePanel.Width - 10, 21),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false // Only enabled if _modeContinueRadio is checked
            };
            _numberingSequencePanel.Controls.Add(_sourceViewCombo);

            // Wire up toggle logic
            _modeNewSheetRadio.CheckedChanged += (s, e) => {
                _startNumberTextBox.Enabled = _modeNewSheetRadio.Checked;
                _sourceViewCombo.Enabled = _modeContinueRadio.Checked;
            };
            _modeContinueRadio.CheckedChanged += (s, e) => {
                _startNumberTextBox.Enabled = _modeNewSheetRadio.Checked;
                _sourceViewCombo.Enabled = _modeContinueRadio.Checked;
            };

            // Populate Floor Plans
            PopulateFloorPlans();

            yPos += 90;

            // Γ£à NEW: System Type Overrides section (always visible, not collapsible)
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

            // Γ£à CRITICAL FIX: Load system types from ALL placed sleeves initially (no category filter)
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

            // Γ£à REMOVED: Big "+ Add Parameter" button
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

            // Γ£à Add header with "+" button aligned with close buttons
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
            // Γ£à PERFORMANCE FIX: Only 4 essential parameters from linked MEP files on startup
            // Rest load when user clicks "+" button
            var systemTypeParam = category == "Cable Trays" ? "Service Type" : "System Type";

            var mappings = new List<(string Mep, string Opening)>
            {
                ("Size", "MEP Size"),  // Γ£à FIX: Use "Size" as default (more commonly used parameter name)
                (systemTypeParam, "MEP System Type"),
                ("System Abbreviation", "MEP System Abbreviation"),
                ("Reference Level", "Level")
            };

            // Γ£à NOTE: Dropdown will include both "Size" and "System Size" if both exist in linked files
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
        /// Γ£à PERFORMANCE FIX: Returns only 4 essential parameters for MEP categories on startup
        /// Rest load when user clicks "+" button
        /// </summary>
        private string[] GetMepParametersForCategory(string category)
        {
            // Handle host elements (Floors, Walls, Structural Framing) - load from architectural/structural linked files
            if (category == "Floors" || category == "Walls" || category == "Structural Framing")
            {
                return GetHostParametersForCategory(category);
            }

            // Γ£à PERFORMANCE FIX: Only return 4 essential parameters for MEP categories on startup
            // 1. Size (or "System Size" - include both as some files use different names)
            // 2. System Type (or "Service Type" for Cable Trays)
            // 3. System Abbreviation
            // 4. Reference Level
            var systemTypeParam = category == "Cable Trays" ? "Service Type" : "System Type";
            // Γ£à FIX: Include both "Size" and "System Size" as different linked files might use different parameter names
            var essentialParams = new List<string> { "Size", "System Size", systemTypeParam, "System Abbreviation", "Reference Level" };
            // Remove duplicates while preserving order
            essentialParams = essentialParams.Distinct().ToList();

            if (_document == null)
            {
                // Fallback to essentials if no document available
                return essentialParams.ToArray();
            }

            // Γ£à PERFORMANCE FIX: Just return the 4 essential parameters on startup
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
        /// Γ£à PERFORMANCE FIX: Only loads "Fire Rating" on startup, rest load when user clicks "+"
        /// </summary>
        private string[] GetHostParametersForCategory(string category)
        {
            // Γ£à PERFORMANCE FIX: Only return "Fire Rating" parameter on startup
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

                // Γ£à PERFORMANCE FIX: Only load "Fire Rating" parameter on startup (rest load when user clicks "+")
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
        /// Γ£à PERFORMANCE FIX: Only loads 4 essential parameters on startup: MEP System Type, MEP Size, MEP System Abbreviation, Level
        /// Rest load when user clicks "+"
        /// </summary>
        private string[] GetOpeningParametersForCategory(string category)
        {
            // Γ£à PERFORMANCE FIX: Only return 4 essential parameters on startup
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

                // Γ£à PERFORMANCE FIX: Only check if the 4 essential parameters exist, then return them
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

            // ≡ƒöì MEP Search button (Greenish)
            var mepSearchBtn = new WinForms.Button
            {
                Text = "≡ƒöì",
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
                Text = "ΓåÆ",
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

            // ≡ƒöì Opening Search button (Yellowish)
            var openingSearchBtn = new WinForms.Button
            {
                Text = "≡ƒöì",
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
                Text = "Γ¿ë",
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
            // Γ£à FIX: Get ALL parameters for search, scanning linked files/host elements
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
            // Γ£à FIX: Get ALL parameters for search, not just the essential ones
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

            // Γ£à CRITICAL FIX: Populate dropdown ONCE when row is created - don't reload on DropDown event
            // This prevents memory access violations from database/Revit API calls during dropdown opening
            // The dropdown will use the cached _systemTypeOptions that were loaded when dialog opened
            PopulateSystemTypeDropdown(systemTypeCombo, systemType);

            // Γ£à REMOVED: DropDown event handler that was causing memory access violations
            // Instead, system types are loaded once when dialog opens and cached in _systemTypeOptions
            // If user needs fresh data, they can close and reopen the dialog
            row.Controls.Add(systemTypeCombo);

            // Arrow
            var arrow = new WinForms.Label { Text = "ΓåÆ", Location = new Point(138, 6), Size = new Size(15, 16) };
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
                Text = "├ù",
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
        /// Γ£à TESTING MODE: Always loads ALL system types and service types from ALL categories (ignores category filter)
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

                    // Γ£à TESTING MODE: Always load ALL types from ALL categories (ignore category parameter)
                    var catalog = catalogService.Load(_document, category: null);

                    // Γ£à LOAD ALL: Include both SystemTypes and ServiceTypes from all categories
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

                // Γ£à ENHANCED LOGGING: Log what will be populated in the dropdown
                var categoryInfo = string.IsNullOrWhiteSpace(category) ? "all categories" : $"category '{category}'";
                if (_systemTypeOptions.Count > 0)
                {
                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                        $"[{DateTime.Now}] [ParameterServiceDialogV2] Γ£à System Type dropdown will be populated with {_systemTypeOptions.Count} options for {categoryInfo}: {string.Join(", ", _systemTypeOptions)}\n");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                        $"[{DateTime.Now}] [ParameterServiceDialogV2] ΓÜá∩╕Å System Type dropdown will be EMPTY for {categoryInfo} - no system/service types found in placed sleeves\n");
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
        /// Γ£à NEW: Populate system type dropdown with available options
        /// </summary>
        private void PopulateSystemTypeDropdown(WinForms.ComboBox comboBox, string selectedValue = null)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add("<Select>");

            // Γ£à CRITICAL FIX: Only show system types from placed sleeves - NO hardcoded fallback
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
        /// Γ£à NEW: Populate Floor Plan dropdown for numbering continuation
        /// </summary>
        private void PopulateFloorPlans()
        {
            if (_document == null || _sourceViewCombo == null) return;

            try 
            {
                var floorPlans = new FilteredElementCollector(_document)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .Where(vp => !vp.IsTemplate)
                    .OrderBy(vp => vp.Name)
                    .ToList();

                _sourceViewCombo.Items.Clear();
                _sourceViewCombo.Items.Add("<Select Floor Plan>");

                foreach (var vp in floorPlans)
                {
                    // Γ£à FIX: Use View Name as requested by user ("view name not levels")
                    // This aligns with MarkParameterService logic which scopes by View Name
                    string displayName = vp.Name;
                    
                    // Optional: Append Level Name for clarity if needed, but sticking to requested "View Name"
                    // string displayName = $"{vp.Name} ({vp.GenLevel?.Name ?? "No Level"})";
                    
                    _sourceViewCombo.Items.Add(displayName);
                }

                if (_sourceViewCombo.Items.Count > 0)
                    _sourceViewCombo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                RemarkDebugLogger.LogError($"Error populating floor plans: {ex.Message}");
            }
        }

        /// <summary>
        /// Γ£à NEW: Detect current category based on which discipline prefix textbox was last focused/edited
        /// Returns "Ducts", "Pipes", "Cable Trays", or null if cannot determine
        /// </summary>
        private string DetectCurrentCategory()
        {
            // Γ£à CRITICAL FIX: First check which textbox currently has focus
            if (_ductPrefixTextBox != null && _ductPrefixTextBox.Focused)
                return "Ducts";
            if (_pipePrefixTextBox != null && _pipePrefixTextBox.Focused)
                return "Pipes";
            if (_cableTrayPrefixTextBox != null && _cableTrayPrefixTextBox.Focused)
                return "Cable Trays";
            if (_damperPrefixTextBox != null && _damperPrefixTextBox.Focused)
                return "Duct Accessories";

            // Γ£à ENHANCEMENT: If no textbox has focus, use the last focused category
            // This handles the case when user clicks dropdown after focusing a textbox
            if (!string.IsNullOrWhiteSpace(_lastFocusedCategory))
            {
                SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                    $"[{DateTime.Now}] [ParameterServiceDialogV2] Using last focused category: '{_lastFocusedCategory}'\n");
                return _lastFocusedCategory;
            }

            // If we can't determine, return null to load all types
            SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                $"[{DateTime.Now}] [ParameterServiceDialogV2] ΓÜá∩╕Å Cannot determine category - loading all system/service types\n");
            return null;
        }

        private void ToggleLock(WinForms.Button lockBtn, WinForms.TextBox textBox)
        {
            if (lockBtn.Text == "≡ƒöÆ")
            {
                lockBtn.Text = "≡ƒöô";
                lockBtn.BackColor = Color.LightCoral;
                textBox.Enabled = true;
                textBox.BackColor = Color.White;
            }
            else
            {
                lockBtn.Text = "≡ƒöÆ";
                lockBtn.BackColor = Color.LightGreen;
                textBox.Enabled = false;
                textBox.BackColor = Color.LightGray;
            }
        }

        private void ToggleLock(WinForms.Button lockBtn, WinForms.ComboBox comboBox)
        {
            if (lockBtn.Text == "≡ƒöÆ")
            {
                lockBtn.Text = "≡ƒöô";
                lockBtn.BackColor = Color.LightCoral;
                comboBox.Enabled = true;
                comboBox.BackColor = Color.White;
            }
            else
            {
                lockBtn.Text = "≡ƒöÆ";
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

                    // Γ£à PERFORMANCE MONITORING: Track parameter transfer performance
                    using (var perfMonitor = new ParameterOperationPerformanceMonitor("Parameter Transfer"))
                    {
                        // Create parameter transfer service
                        var transferService = new ParameterTransferService();

                        // Γ£à OPTIMIZATION: Filter by Active View OR Section Box
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
                                            DebugLogger.Info($"[ParameterServiceDialogV2] Γ£à Section box filtering applied during collection");
                                        }
                                    }
                                }
                                catch (Exception sectionBoxEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[ParameterServiceDialogV2] ΓÜá∩╕Å Section box filter setup failed: {sectionBoxEx.Message}");
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
                            DebugLogger.Info($"[ParameterServiceDialogV2] Γ£à Section box filtering applied during collection: {sleevesToProcess.Count} sleeves within section box");
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

                        // Γ£à CRITICAL FIX: Add source parameter names to learned keys whitelist
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

                        // Γ£à FIX: Use ExecuteTransferConfigurationInTransaction directly to ensure snapshots are loaded
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
        /// Handle Reset Prefix button click
        /// Re-marks only the categories where prefixes have changed
        /// </summary>
        private void OnResetPrefixClick(object sender, EventArgs e)
        {
            WinForms.MessageBox.Show("This functionality is deprecated.", "Deprecated", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// Handle Add Prefix button click
        /// Applies ONLY prefixes (DB-based) to items where prefix has changed
        /// </summary>
        private void OnAddPrefixClick(object sender, EventArgs e)
        {
            WinForms.MessageBox.Show("This functionality is deprecated.", "Deprecated", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// Handle Add Number button click
        /// Applies ONLY numbers (Revit scan) to marked items
        /// </summary>
        private void OnAddNumberClick(object sender, EventArgs e)
        {
            WinForms.MessageBox.Show("This functionality is deprecated.", "Deprecated", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }


        /// <summary>
        /// Γ£à DATABASE-BASED: Check if categories have clash zone data in the database
        /// Returns list of categories that don't have clash data in the database
        /// Replaces the old XML file-based check
        /// </summary>
        private List<string> GetCategoriesWithoutData()
        {
            var missingCategories = new List<string>();

            try
            {
                if (_document == null) return missingCategories;

                // Γ£à DATABASE MIGRATION: Check database instead of XML files
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
        /// Γ£à DEPRECATED: Old XML-based check (kept for fallback only)
        /// Use GetCategoriesWithoutData() instead
        /// </summary>
        [Obsolete("Use GetCategoriesWithoutData() instead - checks database instead of XML files")]
        private List<string> CheckFilterFilesForCategories()
        {
            // Γ£à FALLBACK: Try database first, then XML if database check fails
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
        /// Γ£à RESET NUMBERING: Clears Marks and Resets Counters
        /// </summary>
        private void OnResetNumberingClick(object sender, EventArgs e)
        {
            WinForms.MessageBox.Show("This functionality is deprecated.", "Deprecated", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }


        /// <summary>
        /// Γ£à RESET PARAMETERS: Clears Transferred Parameters
        /// Strictly Session Sensitive (Active View Only) if checked
        /// </summary>
        private void OnResetParametersClick(object sender, EventArgs e)
        {
            WinForms.MessageBox.Show("This functionality is deprecated.", "Deprecated", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// Helper to gather UI settings and validate state
        /// </summary>
        private (object Settings, bool IsValid) GetMarkPrefixSettingsFromUI(bool requireDbData = true)
        {
            return (new object(), true);
        }

        /// <summary>
        /// Γ£à PERSISTENCE: Load saved user inputs from JSON file
        /// </summary>
        private void LoadSavedSettings()
        {
            RemarkDebugLogger.LogInfo("=== LoadSavedSettings CALLED (UI) - DEPRECATED ===");
            // DEPRECATED: Logic stubbed out.
        }

        /// <summary>
        /// Γ£à PERSISTENCE: Save current user inputs to JSON file
        /// </summary>
        private void SaveCurrentSettings()
        {
             // DEPRECATED: Logic stubbed out.
             RemarkDebugLogger.LogInfo("=== SaveCurrentSettings CALLED (UI) - DEPRECATED ===");
        }

        private string GetNumberFormatFromCombo()
        {
            if (_numberFormatCombo == null) return "000";
            int index = _numberFormatCombo.SelectedIndex;
            return index == 0 ? "00" : (index == 1 ? "000" : "0000");
        }
    } }


