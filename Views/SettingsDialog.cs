using System;
using System.Drawing;
using System.IO;
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
        private WinForms.CheckBox _pipeOpeningTypeRectangularCheckBox;
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
        private WinForms.TextBox _roundingValueTextBox;
        private WinForms.CheckBox _roundAlwaysUpCheckBox;
        private WinForms.TextBox _minWallThicknessTextBox;
        private WinForms.CheckBox _ignoreArchitecturalFloorsCheckBox;
        
        // Clash Detection Control (now in Manage section)
        private WinForms.CheckBox _enableThreePointValidationCheckBox;
        private WinForms.CheckBox _forceDetectionModeCheckBox;
        
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
            
            // Debug logging
            System.Diagnostics.Debug.WriteLine($"[SettingsDialog] Loaded settings: ResetApprovalStatus={_settings.ResetApprovalStatus}, CutOpeningWithHosts={_settings.CutOpeningWithHosts}");
            
            // Also log to file
            try
            {
                DebugLogger.Info($"[{DateTime.Now}] [SettingsDialog] Loaded settings: ResetApprovalStatus={_settings.ResetApprovalStatus}, CutOpeningWithHosts={_settings.CutOpeningWithHosts}\n");
            }
            catch { }
            
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Settings";
            // Form height calculation:
            // Manage section: Y=20, Height=85, ends at Y=105
            // Limits section: Y=115, Height=220, ends at Y=335
            // Buttons: Y=345, Height=30, ends at Y=375
            // Need extra margin at bottom to show buttons fully, so total height = 415px
            this.Size = new Drawing.Size(800, 415); // Increased height for 2 items in Manage
            this.AutoScroll = false; // Fixed size - no scroll needed
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
            // Elements section removed - nothing to show (pipe opening type controlled by UI)
            CreateElementFilterSection();
            CreateLimitsSection();
            CreateActionButtons();
            
            this.ResumeLayout(false);
        }

        private void CreateManageSection()
        {
            // Manage Section Group - "Adopt to modified document" and "Force Full Detection Mode" visible
            var manageGroupBox = new WinForms.GroupBox
            {
                Text = "Manage",
                Location = new Drawing.Point(20, 20),
                Size = new Drawing.Size(600, 85), // Height for 2 visible items
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(manageGroupBox);

            int yPos = 25;

            // ✅ HIDDEN: Reset approval status checkbox - Still created for loading/saving but not visible
            _resetApprovalStatusCheckBox = new WinForms.CheckBox
            {
                Text = "",
                Location = new Drawing.Point(-1000, -1000), // Off-screen
                Size = new Drawing.Size(20, 20),
                Checked = true,
                Visible = false // Hidden from UI
            };
            manageGroupBox.Controls.Add(_resetApprovalStatusCheckBox);

            // ✅ HIDDEN: Dimension change threshold - Still created for loading/saving but not visible
            _dimensionChangeThresholdTextBox = new WinForms.TextBox
            {
                Text = "1mm",
                Location = new Drawing.Point(-1000, -1000), // Off-screen
                Size = new Drawing.Size(50, 20),
                Visible = false // Hidden from UI
            };
            manageGroupBox.Controls.Add(_dimensionChangeThresholdTextBox);

            // ✅ HIDDEN: Location change threshold - Still created for loading/saving but not visible
            _locationChangeThresholdTextBox = new WinForms.TextBox
            {
                Text = "1mm",
                Location = new Drawing.Point(-1000, -1000), // Off-screen
                Size = new Drawing.Size(50, 20),
                Visible = false // Hidden from UI
            };
            manageGroupBox.Controls.Add(_locationChangeThresholdTextBox);

            // ✅ VISIBLE: Adopt to modified document checkbox - CHECKBOX ON RIGHT, TEXT ON LEFT
            var adoptModificationLabel = new WinForms.Label
            {
                Text = "Adopt to modified document to adjust clash detections automatically:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20)
            };
            manageGroupBox.Controls.Add(adoptModificationLabel);
            
            _enableThreePointValidationCheckBox = new WinForms.CheckBox
            {
                Text = "", // No text, just checkbox
                Location = new Drawing.Point(520, yPos),
                Size = new Drawing.Size(20, 20),
                Checked = true // Default enabled for safety
            };
            manageGroupBox.Controls.Add(_enableThreePointValidationCheckBox);
            yPos += 25;

            // ✅ NEW: Force Full Detection Mode checkbox - CHECKBOX ON RIGHT, TEXT ON LEFT
            var forceDetectionLabel = new WinForms.Label
            {
                Text = "Force Full Detection Mode (always run PATH 3 - full detection with validation):",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20)
            };
            manageGroupBox.Controls.Add(forceDetectionLabel);
            
            _forceDetectionModeCheckBox = new WinForms.CheckBox
            {
                Text = "", // No text, just checkbox
                Location = new Drawing.Point(520, yPos),
                Size = new Drawing.Size(20, 20),
                Checked = false // Default disabled
            };
            manageGroupBox.Controls.Add(_forceDetectionModeCheckBox);
        }

        // ✅ REMOVED: Elements section - Nothing to show (pipe opening type controlled by UI)
        // Cut opening with hosts and pipe opening type are still loaded/saved but not shown in UI
        // Hidden controls are created in LoadSettings() to ensure they exist for loading/saving

        private void CreateElementFilterSection()
        {
            // Element Filter Section - REMOVED
            // All items moved to Manage section above
        }

        private void CreateLimitsSection()
        {
            // Limits Section Group - ALL 8 ITEMS VISIBLE
            // Positioned after Manage section (20 + 85 height + 10 gap = 115)
            var limitsGroupBox = new WinForms.GroupBox
            {
                Text = "Limits",
                Location = new Drawing.Point(20, 115), // Positioned after Manage section (was 90, now 20 + 85 + 10)
                Size = new Drawing.Size(600, 220), // Height for 8 items
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(limitsGroupBox);

            int yPos = 25;

            // Ignore openings smaller than
            var ignoreSmallLabel = new WinForms.Label
            {
                Text = "Ignore openings smaller than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Increased width for more text space
            };
            limitsGroupBox.Controls.Add(ignoreSmallLabel);

            _ignoreOpeningsSmallerThanTextBox = new WinForms.TextBox
            {
                Text = "0.1mm", // Updated to match image
                Location = new Drawing.Point(520, yPos), // Fixed: removed the -2 offset
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_ignoreOpeningsSmallerThanTextBox);
            yPos += 25;

            // Round openings rectangular
            var roundRectLabel = new WinForms.Label
            {
                Text = "Round openings become rectangular if diameter is greater than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Increased width for more text space
            };
            limitsGroupBox.Controls.Add(roundRectLabel);

            _roundOpeningsRectangularTextBox = new WinForms.TextBox
            {
                Text = "200mm",
                Location = new Drawing.Point(520, yPos), // Fixed: removed the -2 offset
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_roundOpeningsRectangularTextBox);
            yPos += 25;

            // Join openings distance
            var joinDistanceLabel = new WinForms.Label
            {
                Text = "Join openings if their distance is less than:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Increased width for more text space
            };
            limitsGroupBox.Controls.Add(joinDistanceLabel);

            _joinOpeningsDistanceTextBox = new WinForms.TextBox
            {
                Text = "200mm", // Updated to match image
                Location = new Drawing.Point(520, yPos), // Fixed: removed the -2 offset
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_joinOpeningsDistanceTextBox);
            yPos += 25;

            // Round opening sizes to custom value - INPUT BOX FOR ROUNDING VALUE (X-ALIGNED)
            var roundSizesLabel = new WinForms.Label
            {
                Text = "Round opening sizes to nearest value (mm):",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Match other sections width
            };
            limitsGroupBox.Controls.Add(roundSizesLabel);
            
            _roundingValueTextBox = new WinForms.TextBox
            {
                Text = "5", // Default value
                Location = new Drawing.Point(520, yPos), // X-aligned with other textboxes (same as dimension/location thresholds)
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_roundingValueTextBox);
            yPos += 25;

            // Round always up checkbox - CHECKBOX ON RIGHT, TEXT ON LEFT (X-ALIGNED)
            var roundUpLabel = new WinForms.Label
            {
                Text = "Always round up:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Match other sections width
            };
            limitsGroupBox.Controls.Add(roundUpLabel);
            
            _roundAlwaysUpCheckBox = new WinForms.CheckBox
            {
                Text = "", // No text, just checkbox
                Location = new Drawing.Point(520, yPos), // X-aligned with other checkboxes (same as reset approval, cut opening, etc.)
                Size = new Drawing.Size(20, 20),
                Checked = false // Default to round to nearest
            };
            limitsGroupBox.Controls.Add(_roundAlwaysUpCheckBox);
            yPos += 25;

            // Ignore walls with thickness below minimum - INPUT BOX FOR MIN WALL THICKNESS (X-ALIGNED)
            var minWallThicknessLabel = new WinForms.Label
            {
                Text = "Ignore walls if thickness is below (mm):",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Match other sections width
            };
            limitsGroupBox.Controls.Add(minWallThicknessLabel);
            
            _minWallThicknessTextBox = new WinForms.TextBox
            {
                Text = "0", // Default value (0 means disabled)
                Location = new Drawing.Point(520, yPos), // X-aligned with other textboxes
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_minWallThicknessTextBox);
            yPos += 25;

            // Ignore architectural floors checkbox - CHECKBOX ON RIGHT, TEXT ON LEFT (X-ALIGNED)
            var ignoreArchFloorsLabel = new WinForms.Label
            {
                Text = "Ignore architectural floors:",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20) // Match other sections width
            };
            limitsGroupBox.Controls.Add(ignoreArchFloorsLabel);
            
            _ignoreArchitecturalFloorsCheckBox = new WinForms.CheckBox
            {
                Text = "", // No text, just checkbox
                Location = new Drawing.Point(520, yPos), // X-aligned with other checkboxes
                Size = new Drawing.Size(20, 20),
                Checked = false // Default to process all floors
            };
            limitsGroupBox.Controls.Add(_ignoreArchitecturalFloorsCheckBox);
            yPos += 25;

            // Footnote for all sections
            var footnoteLabel = new WinForms.Label
            {
                Text = "Note: All units are in mm",
                Location = new Drawing.Point(15, yPos),
                Size = new Drawing.Size(500, 20),
                Font = new Drawing.Font("Microsoft Sans Serif", 8F, Drawing.FontStyle.Italic),
                ForeColor = Drawing.Color.FromArgb(100, 100, 100) // Gray color for footnote
            };
            limitsGroupBox.Controls.Add(footnoteLabel);
        }


        private void CreateActionButtons()
        {
            // Reset Button
            // Positioned after Limits section (115 + 220 height + 10 gap = 345)
            _resetButton = new WinForms.Button
            {
                Text = "Reset",
                Location = new Drawing.Point(20, 345), // Positioned after Limits section (was 320, now 115 + 220 + 10)
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
                Location = new Drawing.Point(450, 345), // Positioned after Limits section (was 320, now 115 + 220 + 10)
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
                Location = new Drawing.Point(535, 345), // Positioned after Limits section (was 320, now 115 + 220 + 10)
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
            if (_resetApprovalStatusCheckBox != null)
                _resetApprovalStatusCheckBox.Checked = _settings.ResetApprovalStatus;
            if (_dimensionChangeThresholdTextBox != null)
                _dimensionChangeThresholdTextBox.Text = _settings.DimensionChangeThreshold.ToString();
            if (_locationChangeThresholdTextBox != null)
                _locationChangeThresholdTextBox.Text = _settings.LocationChangeThreshold.ToString();
            
            // ✅ HIDDEN: Elements section controls - Create hidden controls if they don't exist
            if (_cutOpeningWithHostsCheckBox == null)
            {
                _cutOpeningWithHostsCheckBox = new WinForms.CheckBox
                {
                    Visible = false,
                    Location = new Drawing.Point(-1000, -1000) // Off-screen
                };
            }
            _cutOpeningWithHostsCheckBox.Checked = _settings.CutOpeningWithHosts;
            
            if (_pipeOpeningTypeRectangularCheckBox == null)
            {
                _pipeOpeningTypeRectangularCheckBox = new WinForms.CheckBox
                {
                    Visible = false,
                    Location = new Drawing.Point(-1000, -1000) // Off-screen
                };
            }
            _pipeOpeningTypeRectangularCheckBox.Checked = _settings.PipeOpeningTypeRectangular;
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
                _ignoreOpeningsSmallerThanTextBox.Text = _settings.IgnoreOpeningsSmallerThan.ToString();
            if (_roundOpeningsRectangularTextBox != null)
                _roundOpeningsRectangularTextBox.Text = _settings.RoundOpeningsRectangular.ToString();
            if (_joinOpeningsDistanceTextBox != null)
                _joinOpeningsDistanceTextBox.Text = _settings.JoinOpeningsDistance.ToString();
            if (_createOpeningsWithSlopeCheckBox != null)
                _createOpeningsWithSlopeCheckBox.Checked = _settings.CreateOpeningsWithSlope;
            if (_roundingValueTextBox != null)
                _roundingValueTextBox.Text = _settings.RoundingValue.ToString();
            if (_roundAlwaysUpCheckBox != null)
                _roundAlwaysUpCheckBox.Checked = _settings.RoundAlwaysUp;
            if (_minWallThicknessTextBox != null)
                _minWallThicknessTextBox.Text = _settings.MinWallThickness.ToString();
            if (_ignoreArchitecturalFloorsCheckBox != null)
                _ignoreArchitecturalFloorsCheckBox.Checked = _settings.IgnoreArchitecturalFloors;
            
            if (_enableThreePointValidationCheckBox != null)
                _enableThreePointValidationCheckBox.Checked = _settings.EnableThreePointValidation;
            
            if (_forceDetectionModeCheckBox != null)
                _forceDetectionModeCheckBox.Checked = _settings.ForceDetectionMode;
        }

        private void SaveSettings()
        {
            try
            {
                // Save settings to the model - only for fields that actually exist
                if (_resetApprovalStatusCheckBox != null)
                    _settings.ResetApprovalStatus = _resetApprovalStatusCheckBox.Checked;
                
                if (_dimensionChangeThresholdTextBox != null && double.TryParse(_dimensionChangeThresholdTextBox.Text, out double dimThreshold))
                    _settings.DimensionChangeThreshold = dimThreshold;
                
                if (_locationChangeThresholdTextBox != null && double.TryParse(_locationChangeThresholdTextBox.Text, out double locThreshold))
                    _settings.LocationChangeThreshold = locThreshold;
                
                if (_cutOpeningWithHostsCheckBox != null)
                    _settings.CutOpeningWithHosts = _cutOpeningWithHostsCheckBox.Checked;
                
                if (_pipeOpeningTypeRectangularCheckBox != null)
                    _settings.PipeOpeningTypeRectangular = _pipeOpeningTypeRectangularCheckBox.Checked;
                
                if (_ignoreOpeningsSmallerThanTextBox != null && double.TryParse(_ignoreOpeningsSmallerThanTextBox.Text, out double ignoreSmall))
                    _settings.IgnoreOpeningsSmallerThan = ignoreSmall;
                
                if (_roundOpeningsRectangularTextBox != null && double.TryParse(_roundOpeningsRectangularTextBox.Text, out double roundRect))
                    _settings.RoundOpeningsRectangular = roundRect;
                
                if (_joinOpeningsDistanceTextBox != null && double.TryParse(_joinOpeningsDistanceTextBox.Text, out double joinDist))
                    _settings.JoinOpeningsDistance = joinDist;

                if (_roundingValueTextBox != null && double.TryParse(_roundingValueTextBox.Text, out double roundValue))
                    _settings.RoundingValue = roundValue;

                if (_roundAlwaysUpCheckBox != null)
                    _settings.RoundAlwaysUp = _roundAlwaysUpCheckBox.Checked;

                if (_minWallThicknessTextBox != null && double.TryParse(_minWallThicknessTextBox.Text, out double minWallThickness))
                    _settings.MinWallThickness = minWallThickness;

                if (_ignoreArchitecturalFloorsCheckBox != null)
                    _settings.IgnoreArchitecturalFloors = _ignoreArchitecturalFloorsCheckBox.Checked;
                
                if (_enableThreePointValidationCheckBox != null)
                    _settings.EnableThreePointValidation = _enableThreePointValidationCheckBox.Checked;
                
                if (_forceDetectionModeCheckBox != null)
                    _settings.ForceDetectionMode = _forceDetectionModeCheckBox.Checked;

                // Log successful save
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_debug.log"), $"[{DateTime.Now}] Settings saved successfully\n");
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(SafeFileLogger.GetLogFilePath("refresh_debug.log"), $"[{DateTime.Now}] Error in SaveSettings: {ex.Message}\n");
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
                
                // Debug logging
                System.Diagnostics.Debug.WriteLine($"[SettingsDialog] Saving settings: ResetApprovalStatus={_settings.ResetApprovalStatus}, CutOpeningWithHosts={_settings.CutOpeningWithHosts}");
                
                // Also log to file
                try
                {
                    DebugLogger.Info($"[{DateTime.Now}] [SettingsDialog] Saving settings: ResetApprovalStatus={_settings.ResetApprovalStatus}, CutOpeningWithHosts={_settings.CutOpeningWithHosts}\n");
                }
                catch { }
                
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