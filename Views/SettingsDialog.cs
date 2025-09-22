using System;
using System.Drawing;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

using JSE_RevitAddin_MEP_OPENINGS.Services;
namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Settings/Configuration dialog for MEP Openings
    /// </summary>
    public partial class SettingsDialog : WinForms.Form
    {
        private SettingsModel _settings;
        
        // Manage Section Controls
        private WinForms.CheckBox _resetApprovalStatusCheckBox;
        private WinForms.TextBox _dimensionChangeThresholdTextBox;
        private WinForms.TextBox _locationChangeThresholdTextBox;
        
        // Elements Section Controls
        private WinForms.CheckBox _cutOpeningWithHostsCheckBox;
        private WinForms.CheckBox _createConstraintCheckBox;
        private WinForms.CheckBox _createVerticalOpeningsCheckBox;
        private WinForms.CheckBox _createHorizontalOpeningsCheckBox;
        private WinForms.CheckBox _adoptProvisionForVoidsCheckBox;
        
        // Element Filter Section Controls
        private WinForms.ComboBox _elementFilterComboBox;
        private WinForms.CheckBox _includeHostElementsNotVisibleCheckBox;
        private WinForms.CheckBox _includeReferenceElementsNotVisibleCheckBox;
        private WinForms.CheckBox _includeHostElementsDemolishedCheckBox;
        
        // Limits Section Controls
        private WinForms.TextBox _ignoreOpeningsSmallerThanTextBox;
        private WinForms.TextBox _roundOpeningsRectangularTextBox;
        private WinForms.TextBox _joinOpeningsDistanceTextBox;
        private WinForms.CheckBox _createOpeningsWithSlopeCheckBox;
        private WinForms.ComboBox _roundUpDimensionsComboBox;
        
        // Action Buttons
        private WinForms.Button _resetButton;
        private WinForms.Button _okButton;
        private WinForms.Button _cancelButton;

        public SettingsDialog(SettingsModel settings)
        {
            _settings = settings ?? new SettingsModel();
            InitializeComponent();
            LoadSettings();
        }
        
        public SettingsDialog()
        {
            // Load settings from file
            var settingsService = new SettingsService();
            _settings = settingsService.LoadSettings();
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Settings";
            this.Size = new Drawing.Size(650, 600); // Reduced size since we consolidated sections
            this.StartPosition = WinForms.FormStartPosition.CenterParent; // Center on parent window
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Drawing.Color.FromArgb(240, 240, 240);
            this.TopMost = true; // Keep on top of parent
            this.ShowInTaskbar = false; // Don't show in taskbar for modal dialog
            this.ControlBox = true; // Ensure close button is visible
            this.KeyPreview = true; // Enable keyboard events
            
            // Create sections
            CreateManageSection();
            CreateElementsSection();
            CreateElementFilterSection();
            CreateLimitsSection();
            CreateActionButtons();
            
            this.ResumeLayout(false);
        }

        private void CreateManageSection()
        {
            // Manage Section Group - BIGGER to show 2-line text properly
            var manageGroupBox = new WinForms.GroupBox
            {
                Text = "Manage",
                Location = new Drawing.Point(20, 20),
                Size = new Drawing.Size(600, 120), // Increased height for 2-line text
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(manageGroupBox);

            int yPos = 25;

            // Reset approval status label - LEFT
            var resetApprovalLabel = new WinForms.Label
            {
                Text = "Reset approval status of openings when changes occur:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20)
            };
            manageGroupBox.Controls.Add(resetApprovalLabel);

            // Reset approval status checkbox - RIGHT
            _resetApprovalStatusCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new Drawing.Point(480, yPos),
                Size = new Drawing.Size(20, 20),
                Checked = true
            };
            manageGroupBox.Controls.Add(_resetApprovalStatusCheckBox);
            yPos += 30; // More space for 2-line text

            // Dimension change threshold - LABEL ON LEFT, INPUT ON RIGHT
            var dimensionLabel = new WinForms.Label
            {
                Text = "Openings won't be marked as changed if change in dimensions is less than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20) // Label width
            };
            manageGroupBox.Controls.Add(dimensionLabel);

            _dimensionChangeThresholdTextBox = new WinForms.TextBox
            {
                Text = "1",
                Location = new Drawing.Point(480, yPos - 2), // Right side
                Size = new Drawing.Size(50, 20)
            };
            manageGroupBox.Controls.Add(_dimensionChangeThresholdTextBox);

            var dimensionUnitLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new Drawing.Point(535, yPos - 2),
                Size = new Drawing.Size(30, 20)
            };
            manageGroupBox.Controls.Add(dimensionUnitLabel);
            yPos += 30; // More space for 2-line text

            // Location change threshold - LABEL ON LEFT, INPUT ON RIGHT
            var locationLabel = new WinForms.Label
            {
                Text = "Openings won't be marked as changed if change in location is less than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20) // Label width
            };
            manageGroupBox.Controls.Add(locationLabel);

            _locationChangeThresholdTextBox = new WinForms.TextBox
            {
                Text = "1",
                Location = new Drawing.Point(480, yPos - 2), // Right side
                Size = new Drawing.Size(50, 20)
            };
            manageGroupBox.Controls.Add(_locationChangeThresholdTextBox);

            var locationUnitLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new Drawing.Point(535, yPos - 2),
                Size = new Drawing.Size(30, 20)
            };
            manageGroupBox.Controls.Add(locationUnitLabel);
        }

        private void CreateElementsSection()
        {
            // Elements Section Group - ONLY 2 ITEMS (as per image)
            var elementsGroupBox = new WinForms.GroupBox
            {
                Text = "Elements",
                Location = new Drawing.Point(20, 150), // Positioned after bigger Manage section
                Size = new Drawing.Size(600, 80), // Height for 2 items
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(elementsGroupBox);

            int yPos = 25;

            // Cut opening with hosts label - LEFT
            var cutOpeningLabel = new WinForms.Label
            {
                Text = "Cut opening with Hosts:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20)
            };
            elementsGroupBox.Controls.Add(cutOpeningLabel);

            // Cut opening with hosts checkbox - RIGHT
            _cutOpeningWithHostsCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new Drawing.Point(480, yPos),
                Size = new Drawing.Size(20, 20),
                Checked = false
            };
            elementsGroupBox.Controls.Add(_cutOpeningWithHostsCheckBox);
            yPos += 25;

            // Create constraint label - LEFT
            var createConstraintLabel = new WinForms.Label
            {
                Text = "Create a constraint between openings and Hosts:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20)
            };
            elementsGroupBox.Controls.Add(createConstraintLabel);

            // Create constraint checkbox - RIGHT
            _createConstraintCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new Drawing.Point(480, yPos),
                Size = new Drawing.Size(20, 20),
                Checked = true
            };
            elementsGroupBox.Controls.Add(_createConstraintCheckBox);
        }

        private void CreateElementFilterSection()
        {
            // Element Filter Section - REMOVED
            // All items moved to Manage section above
        }

        private void CreateLimitsSection()
        {
            // Limits Section Group - REMAINING ITEMS (as per image)
            var limitsGroupBox = new WinForms.GroupBox
            {
                Text = "Limits",
                Location = new Drawing.Point(20, 240), // Positioned after bigger Elements section
                Size = new Drawing.Size(600, 120), // Height for 4 items
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(limitsGroupBox);

            int yPos = 25;

            // Ignore openings smaller than
            var ignoreSmallLabel = new WinForms.Label
            {
                Text = "Ignore openings smaller than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20) // Label width
            };
            limitsGroupBox.Controls.Add(ignoreSmallLabel);

            _ignoreOpeningsSmallerThanTextBox = new WinForms.TextBox
            {
                Text = "0.1", // Updated to match image
                Location = new Drawing.Point(480, yPos - 2), // Right side
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_ignoreOpeningsSmallerThanTextBox);

            var ignoreSmallUnitLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new Drawing.Point(535, yPos - 2),
                Size = new Drawing.Size(30, 20)
            };
            limitsGroupBox.Controls.Add(ignoreSmallUnitLabel);
            yPos += 25;

            // Round openings rectangular
            var roundRectLabel = new WinForms.Label
            {
                Text = "Round openings become rectangular if diameter is greater than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20) // Label width
            };
            limitsGroupBox.Controls.Add(roundRectLabel);

            _roundOpeningsRectangularTextBox = new WinForms.TextBox
            {
                Text = "200",
                Location = new Drawing.Point(480, yPos - 2), // Right side
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_roundOpeningsRectangularTextBox);

            var roundRectUnitLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new Drawing.Point(535, yPos - 2),
                Size = new Drawing.Size(30, 20)
            };
            limitsGroupBox.Controls.Add(roundRectUnitLabel);
            yPos += 25;

            // Join openings distance
            var joinDistanceLabel = new WinForms.Label
            {
                Text = "Join openings if their distance is less than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(460, 20) // Label width
            };
            limitsGroupBox.Controls.Add(joinDistanceLabel);

            _joinOpeningsDistanceTextBox = new WinForms.TextBox
            {
                Text = "200", // Updated to match image
                Location = new Drawing.Point(480, yPos - 2), // Right side
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_joinOpeningsDistanceTextBox);

            var joinDistanceUnitLabel = new WinForms.Label
            {
                Text = "mm",
                Location = new Drawing.Point(535, yPos - 2),
                Size = new Drawing.Size(30, 20)
            };
            limitsGroupBox.Controls.Add(joinDistanceUnitLabel);
        }

        private void CreateActionButtons()
        {
            // Reset Button
            _resetButton = new WinForms.Button
            {
                Text = "Reset",
                Location = new Drawing.Point(20, 380), // Positioned after bigger Limits section
                Size = new Drawing.Size(75, 30),
                BackColor = Drawing.Color.FromArgb(200, 200, 200),
                FlatStyle = WinForms.FlatStyle.Flat
            };
            _resetButton.Click += OnResetClick;
            this.Controls.Add(_resetButton);

            // OK Button
            _okButton = new WinForms.Button
            {
                Text = "OK",
                Location = new Drawing.Point(450, 380), // Positioned after bigger Limits section
                Size = new Drawing.Size(75, 30),
                BackColor = Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat
            };
            _okButton.Click += OnOkClick;
            this.Controls.Add(_okButton);

            // Cancel Button
            _cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                Location = new Drawing.Point(535, 380), // Positioned after bigger Limits section
                Size = new Drawing.Size(75, 30),
                BackColor = Drawing.Color.FromArgb(200, 200, 200),
                FlatStyle = WinForms.FlatStyle.Flat
            };
            _cancelButton.Click += OnCancelClick;
            this.Controls.Add(_cancelButton);
            
            // Add keyboard event handling
            this.KeyDown += OnKeyDown;
        }
        
        private void OnKeyDown(object? sender, WinForms.KeyEventArgs e)
        {
            if (e.KeyCode == WinForms.Keys.Escape)
            {
                this.DialogResult = WinForms.DialogResult.Cancel;
                this.Close();
            }
            else if (e.KeyCode == WinForms.Keys.Enter && e.Control)
            {
                // Ctrl+Enter to save
                OnOkClick(sender, e);
            }
        }

        private void LoadSettings()
        {
            // Load settings from the model - with null checks
            // Convert from feet (Revit internal) to mm for display
            const double feetToMm = 304.8;

            if (_resetApprovalStatusCheckBox != null)
                _resetApprovalStatusCheckBox.Checked = _settings.ResetApprovalStatus;
            if (_dimensionChangeThresholdTextBox != null)
                _dimensionChangeThresholdTextBox.Text = (_settings.DimensionChangeThreshold * feetToMm).ToString("F1");
            if (_locationChangeThresholdTextBox != null)
                _locationChangeThresholdTextBox.Text = (_settings.LocationChangeThreshold * feetToMm).ToString("F1");

            if (_cutOpeningWithHostsCheckBox != null)
                _cutOpeningWithHostsCheckBox.Checked = _settings.CutOpeningWithHosts;
            if (_createConstraintCheckBox != null)
                _createConstraintCheckBox.Checked = _settings.CreateConstraint;
            if (_createVerticalOpeningsCheckBox != null)
                _createVerticalOpeningsCheckBox.Checked = _settings.CreateVerticalOpenings;
            if (_createHorizontalOpeningsCheckBox != null)
                _createHorizontalOpeningsCheckBox.Checked = _settings.CreateHorizontalOpenings;
            if (_adoptProvisionForVoidsCheckBox != null)
                _adoptProvisionForVoidsCheckBox.Checked = _settings.AdoptProvisionForVoids;

            if (_elementFilterComboBox != null)
                _elementFilterComboBox.SelectedItem = _settings.ElementFilter;
            if (_includeHostElementsNotVisibleCheckBox != null)
                _includeHostElementsNotVisibleCheckBox.Checked = _settings.IncludeHostElementsNotVisible;
            if (_includeReferenceElementsNotVisibleCheckBox != null)
                _includeReferenceElementsNotVisibleCheckBox.Checked = _settings.IncludeReferenceElementsNotVisible;
            if (_includeHostElementsDemolishedCheckBox != null)
                _includeHostElementsDemolishedCheckBox.Checked = _settings.IncludeHostElementsDemolished;

            if (_ignoreOpeningsSmallerThanTextBox != null)
                _ignoreOpeningsSmallerThanTextBox.Text = (_settings.IgnoreOpeningsSmallerThan * feetToMm).ToString("F1");
            if (_roundOpeningsRectangularTextBox != null)
                _roundOpeningsRectangularTextBox.Text = (_settings.RoundOpeningsRectangular * feetToMm).ToString("F0");
            if (_joinOpeningsDistanceTextBox != null)
                _joinOpeningsDistanceTextBox.Text = (_settings.JoinOpeningsDistance * feetToMm).ToString("F0");
            if (_createOpeningsWithSlopeCheckBox != null)
                _createOpeningsWithSlopeCheckBox.Checked = _settings.CreateOpeningsWithSlope;
            if (_roundUpDimensionsComboBox != null)
                _roundUpDimensionsComboBox.SelectedItem = _settings.RoundUpDimensions;
        }

        private void SaveSettings()
        {
            try
            {
                // Save settings to the model - only for fields that actually exist
                // Convert from mm (UI) to feet (Revit internal)
                const double mmToFeet = 1.0 / 304.8;

                if (_resetApprovalStatusCheckBox != null)
                    _settings.ResetApprovalStatus = _resetApprovalStatusCheckBox.Checked;

                if (_dimensionChangeThresholdTextBox != null && double.TryParse(_dimensionChangeThresholdTextBox.Text, out double dimThreshold))
                    _settings.DimensionChangeThreshold = dimThreshold * mmToFeet;

                if (_locationChangeThresholdTextBox != null && double.TryParse(_locationChangeThresholdTextBox.Text, out double locThreshold))
                    _settings.LocationChangeThreshold = locThreshold * mmToFeet;

                if (_cutOpeningWithHostsCheckBox != null)
                    _settings.CutOpeningWithHosts = _cutOpeningWithHostsCheckBox.Checked;

                if (_createConstraintCheckBox != null)
                    _settings.CreateConstraint = _createConstraintCheckBox.Checked;

                if (_ignoreOpeningsSmallerThanTextBox != null && double.TryParse(_ignoreOpeningsSmallerThanTextBox.Text, out double ignoreSmall))
                    _settings.IgnoreOpeningsSmallerThan = ignoreSmall * mmToFeet;

                if (_roundOpeningsRectangularTextBox != null && double.TryParse(_roundOpeningsRectangularTextBox.Text, out double roundRect))
                    _settings.RoundOpeningsRectangular = roundRect * mmToFeet;

                if (_joinOpeningsDistanceTextBox != null && double.TryParse(_joinOpeningsDistanceTextBox.Text, out double joinDist))
                    _settings.JoinOpeningsDistance = joinDist * mmToFeet;

                // Log successful save
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Settings saved successfully\n");
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", $"[{DateTime.Now}] Error in SaveSettings: {ex.Message}\n");
                throw; // Re-throw to show error dialog
            }
        }

        private void OnResetClick(object? sender, EventArgs e)
        {
            // Reset to default values
            _settings = new SettingsModel();
            LoadSettings();
        }

        private void OnOkClick(object? sender, EventArgs e)
        {
            try
            {
                SaveSettings();
                
                // Save to file using SettingsService
                var settingsService = new SettingsService();
                settingsService.SaveSettings(_settings);
                
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                // Show error message instead of unhandled exception
                WinForms.MessageBox.Show($"Error saving settings: {ex.Message}", "Settings Error", 
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnCancelClick(object? sender, EventArgs e)
        {
            this.DialogResult = WinForms.DialogResult.Cancel;
            this.Close();
        }

        public SettingsModel GetSettings()
        {
            return _settings;
        }
    }
}
