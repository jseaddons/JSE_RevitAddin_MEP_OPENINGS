using System;
using System.Drawing;
using System.Windows.Forms;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

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
        private WinForms.TextBox _ignoreOpeningsAngleTextBox;
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

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Form properties
            this.Text = "Settings";
            this.Size = new Drawing.Size(500, 600);
            this.StartPosition = WinForms.FormStartPosition.CenterParent;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Drawing.Color.FromArgb(240, 240, 240);
            
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
            // Manage Section Group
            var manageGroupBox = new WinForms.GroupBox
            {
                Text = "Manage",
                Location = new Drawing.Point(20, 20),
                Size = new Drawing.Size(460, 100),
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(manageGroupBox);

            // Reset approval status checkbox
            _resetApprovalStatusCheckBox = new WinForms.CheckBox
            {
                Text = "Reset approval status of openings when changes occur:",
                Location = new Drawing.Point(15, 25),
                Size = new Drawing.Size(400, 20),
                Checked = true
            };
            manageGroupBox.Controls.Add(_resetApprovalStatusCheckBox);

            // Dimension change threshold
            var dimensionLabel = new WinForms.Label
            {
                Text = "Openings won't be marked as changed if change in dimensions is less than:",
                Location = new Drawing.Point(15, 50),
                Size = new Drawing.Size(300, 20)
            };
            manageGroupBox.Controls.Add(dimensionLabel);

            _dimensionChangeThresholdTextBox = new WinForms.TextBox
            {
                Text = "1",
                Location = new Drawing.Point(320, 48),
                Size = new Drawing.Size(50, 20)
            };
            manageGroupBox.Controls.Add(_dimensionChangeThresholdTextBox);
        }

        private void CreateElementsSection()
        {
            // Elements Section Group
            var elementsGroupBox = new WinForms.GroupBox
            {
                Text = "Elements",
                Location = new Drawing.Point(20, 130),
                Size = new Drawing.Size(460, 150),
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(elementsGroupBox);

            // Cut opening with hosts checkbox
            _cutOpeningWithHostsCheckBox = new WinForms.CheckBox
            {
                Text = "Cut opening with Hosts:",
                Location = new Drawing.Point(15, 25),
                Size = new Drawing.Size(200, 20),
                Checked = false
            };
            elementsGroupBox.Controls.Add(_cutOpeningWithHostsCheckBox);

            // Create constraint checkbox
            _createConstraintCheckBox = new WinForms.CheckBox
            {
                Text = "Create a constraint between openings and Hosts:",
                Location = new Drawing.Point(15, 50),
                Size = new Drawing.Size(300, 20),
                Checked = true
            };
            elementsGroupBox.Controls.Add(_createConstraintCheckBox);

            // Create vertical openings checkbox
            _createVerticalOpeningsCheckBox = new WinForms.CheckBox
            {
                Text = "Create vertical openings if Reference and Host-Elements are parallel:",
                Location = new Drawing.Point(15, 75),
                Size = new Drawing.Size(400, 20),
                Checked = false
            };
            elementsGroupBox.Controls.Add(_createVerticalOpeningsCheckBox);

            // Create horizontal openings checkbox
            _createHorizontalOpeningsCheckBox = new WinForms.CheckBox
            {
                Text = "Create horizontal openings if Reference and Host-Elements are parallel:",
                Location = new Drawing.Point(15, 100),
                Size = new Drawing.Size(400, 20),
                Checked = false
            };
            elementsGroupBox.Controls.Add(_createHorizontalOpeningsCheckBox);

            // Adopt provision for voids checkbox
            _adoptProvisionForVoidsCheckBox = new WinForms.CheckBox
            {
                Text = "Adopt Provision for Voids (Openings) from linked Model:",
                Location = new Drawing.Point(15, 125),
                Size = new Drawing.Size(400, 20),
                Checked = false
            };
            elementsGroupBox.Controls.Add(_adoptProvisionForVoidsCheckBox);
        }

        private void CreateElementFilterSection()
        {
            // Element Filter Section Group
            var filterGroupBox = new WinForms.GroupBox
            {
                Text = "Element Filter",
                Location = new Drawing.Point(20, 290),
                Size = new Drawing.Size(460, 120),
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(filterGroupBox);

            // Element filter dropdown
            var filterLabel = new WinForms.Label
            {
                Text = "View Filter:",
                Location = new Drawing.Point(15, 25),
                Size = new Drawing.Size(80, 20)
            };
            filterGroupBox.Controls.Add(filterLabel);

            _elementFilterComboBox = new WinForms.ComboBox
            {
                Location = new Drawing.Point(100, 23),
                Size = new Drawing.Size(100, 20),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            _elementFilterComboBox.Items.AddRange(new[] { "{3D}", "{Plan}", "{Section}" });
            _elementFilterComboBox.SelectedIndex = 0;
            filterGroupBox.Controls.Add(_elementFilterComboBox);

            // Include host elements not visible checkbox
            _includeHostElementsNotVisibleCheckBox = new WinForms.CheckBox
            {
                Text = "Include Host Elements not visible in the selected 3D view:",
                Location = new Drawing.Point(15, 50),
                Size = new Drawing.Size(400, 20),
                Checked = true
            };
            filterGroupBox.Controls.Add(_includeHostElementsNotVisibleCheckBox);

            // Include reference elements not visible checkbox
            _includeReferenceElementsNotVisibleCheckBox = new WinForms.CheckBox
            {
                Text = "Include Reference Elements not visible in the selected 3D view:",
                Location = new Drawing.Point(15, 75),
                Size = new Drawing.Size(400, 20),
                Checked = true
            };
            filterGroupBox.Controls.Add(_includeReferenceElementsNotVisibleCheckBox);

            // Include host elements demolished checkbox
            _includeHostElementsDemolishedCheckBox = new WinForms.CheckBox
            {
                Text = "Include Host Elements in demolished Phase:",
                Location = new Drawing.Point(15, 100),
                Size = new Drawing.Size(300, 20),
                Checked = false
            };
            filterGroupBox.Controls.Add(_includeHostElementsDemolishedCheckBox);
        }

        private void CreateLimitsSection()
        {
            // Limits Section Group
            var limitsGroupBox = new WinForms.GroupBox
            {
                Text = "Limits",
                Location = new Drawing.Point(20, 420),
                Size = new Drawing.Size(460, 120),
                Font = new Drawing.Font("Microsoft Sans Serif", 9F, Drawing.FontStyle.Bold)
            };
            this.Controls.Add(limitsGroupBox);

            // Ignore openings smaller than
            var ignoreSmallLabel = new WinForms.Label
            {
                Text = "Ignore openings smaller than:",
                Location = new Drawing.Point(15, 25),
                Size = new Drawing.Size(200, 20)
            };
            limitsGroupBox.Controls.Add(ignoreSmallLabel);

            _ignoreOpeningsSmallerThanTextBox = new WinForms.TextBox
            {
                Text = "10",
                Location = new Drawing.Point(220, 23),
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_ignoreOpeningsSmallerThanTextBox);

            // Round openings rectangular
            var roundRectLabel = new WinForms.Label
            {
                Text = "Round openings become rectangular if diameter is greater than:",
                Location = new Drawing.Point(15, 50),
                Size = new Drawing.Size(300, 20)
            };
            limitsGroupBox.Controls.Add(roundRectLabel);

            _roundOpeningsRectangularTextBox = new WinForms.TextBox
            {
                Text = "200",
                Location = new Drawing.Point(320, 48),
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_roundOpeningsRectangularTextBox);

            // Join openings distance
            var joinDistanceLabel = new WinForms.Label
            {
                Text = "Join openings if their distance is less than:",
                Location = new Drawing.Point(15, 75),
                Size = new Drawing.Size(250, 20)
            };
            limitsGroupBox.Controls.Add(joinDistanceLabel);

            _joinOpeningsDistanceTextBox = new WinForms.TextBox
            {
                Text = "200",
                Location = new Drawing.Point(270, 73),
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_joinOpeningsDistanceTextBox);

            // Ignore openings angle
            var ignoreAngleLabel = new WinForms.Label
            {
                Text = "Ignore openings with an angle greater than:",
                Location = new Drawing.Point(15, 100),
                Size = new Drawing.Size(250, 20)
            };
            limitsGroupBox.Controls.Add(ignoreAngleLabel);

            _ignoreOpeningsAngleTextBox = new WinForms.TextBox
            {
                Text = "45.00°",
                Location = new Drawing.Point(270, 98),
                Size = new Drawing.Size(50, 20)
            };
            limitsGroupBox.Controls.Add(_ignoreOpeningsAngleTextBox);

            // Create openings with slope checkbox
            _createOpeningsWithSlopeCheckBox = new WinForms.CheckBox
            {
                Text = "Create openings with a slope:",
                Location = new Drawing.Point(15, 125),
                Size = new Drawing.Size(200, 20),
                Checked = true
            };
            limitsGroupBox.Controls.Add(_createOpeningsWithSlopeCheckBox);

            // Round up dimensions dropdown
            var roundUpLabel = new WinForms.Label
            {
                Text = "Round up opening dimensions:",
                Location = new Drawing.Point(15, 150),
                Size = new Drawing.Size(200, 20)
            };
            limitsGroupBox.Controls.Add(roundUpLabel);

            _roundUpDimensionsComboBox = new WinForms.ComboBox
            {
                Location = new Drawing.Point(220, 148),
                Size = new Drawing.Size(150, 20),
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            _roundUpDimensionsComboBox.Items.AddRange(new[] { "Do not round up", "Round up", "Round down" });
            _roundUpDimensionsComboBox.SelectedIndex = 0;
            limitsGroupBox.Controls.Add(_roundUpDimensionsComboBox);
        }

        private void CreateActionButtons()
        {
            // Reset Button
            _resetButton = new WinForms.Button
            {
                Text = "Reset",
                Location = new Drawing.Point(20, 550),
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
                Location = new Drawing.Point(350, 550),
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
                Location = new Drawing.Point(430, 550),
                Size = new Drawing.Size(75, 30),
                BackColor = Drawing.Color.FromArgb(200, 200, 200),
                FlatStyle = WinForms.FlatStyle.Flat
            };
            _cancelButton.Click += OnCancelClick;
            this.Controls.Add(_cancelButton);
        }

        private void LoadSettings()
        {
            // Load settings from the model
            _resetApprovalStatusCheckBox.Checked = _settings.ResetApprovalStatus;
            _dimensionChangeThresholdTextBox.Text = _settings.DimensionChangeThreshold.ToString();
            _locationChangeThresholdTextBox.Text = _settings.LocationChangeThreshold.ToString();
            
            _cutOpeningWithHostsCheckBox.Checked = _settings.CutOpeningWithHosts;
            _createConstraintCheckBox.Checked = _settings.CreateConstraint;
            _createVerticalOpeningsCheckBox.Checked = _settings.CreateVerticalOpenings;
            _createHorizontalOpeningsCheckBox.Checked = _settings.CreateHorizontalOpenings;
            _adoptProvisionForVoidsCheckBox.Checked = _settings.AdoptProvisionForVoids;
            
            _elementFilterComboBox.SelectedItem = _settings.ElementFilter;
            _includeHostElementsNotVisibleCheckBox.Checked = _settings.IncludeHostElementsNotVisible;
            _includeReferenceElementsNotVisibleCheckBox.Checked = _settings.IncludeReferenceElementsNotVisible;
            _includeHostElementsDemolishedCheckBox.Checked = _settings.IncludeHostElementsDemolished;
            
            _ignoreOpeningsSmallerThanTextBox.Text = _settings.IgnoreOpeningsSmallerThan.ToString();
            _roundOpeningsRectangularTextBox.Text = _settings.RoundOpeningsRectangular.ToString();
            _joinOpeningsDistanceTextBox.Text = _settings.JoinOpeningsDistance.ToString();
            _ignoreOpeningsAngleTextBox.Text = _settings.IgnoreOpeningsAngle.ToString();
            _createOpeningsWithSlopeCheckBox.Checked = _settings.CreateOpeningsWithSlope;
            _roundUpDimensionsComboBox.SelectedItem = _settings.RoundUpDimensions;
        }

        private void SaveSettings()
        {
            // Save settings to the model
            _settings.ResetApprovalStatus = _resetApprovalStatusCheckBox.Checked;
            if (double.TryParse(_dimensionChangeThresholdTextBox.Text, out double dimThreshold))
                _settings.DimensionChangeThreshold = dimThreshold;
            if (double.TryParse(_locationChangeThresholdTextBox.Text, out double locThreshold))
                _settings.LocationChangeThreshold = locThreshold;
            
            _settings.CutOpeningWithHosts = _cutOpeningWithHostsCheckBox.Checked;
            _settings.CreateConstraint = _createConstraintCheckBox.Checked;
            _settings.CreateVerticalOpenings = _createVerticalOpeningsCheckBox.Checked;
            _settings.CreateHorizontalOpenings = _createHorizontalOpeningsCheckBox.Checked;
            _settings.AdoptProvisionForVoids = _adoptProvisionForVoidsCheckBox.Checked;
            
            _settings.ElementFilter = _elementFilterComboBox.SelectedItem?.ToString() ?? "{3D}";
            _settings.IncludeHostElementsNotVisible = _includeHostElementsNotVisibleCheckBox.Checked;
            _settings.IncludeReferenceElementsNotVisible = _includeReferenceElementsNotVisibleCheckBox.Checked;
            _settings.IncludeHostElementsDemolished = _includeHostElementsDemolishedCheckBox.Checked;
            
            if (double.TryParse(_ignoreOpeningsSmallerThanTextBox.Text, out double ignoreSmall))
                _settings.IgnoreOpeningsSmallerThan = ignoreSmall;
            if (double.TryParse(_roundOpeningsRectangularTextBox.Text, out double roundRect))
                _settings.RoundOpeningsRectangular = roundRect;
            if (double.TryParse(_joinOpeningsDistanceTextBox.Text, out double joinDist))
                _settings.JoinOpeningsDistance = joinDist;
            if (double.TryParse(_ignoreOpeningsAngleTextBox.Text.Replace("°", ""), out double ignoreAngle))
                _settings.IgnoreOpeningsAngle = ignoreAngle;
            _settings.CreateOpeningsWithSlope = _createOpeningsWithSlopeCheckBox.Checked;
            _settings.RoundUpDimensions = _roundUpDimensionsComboBox.SelectedItem?.ToString() ?? "Do not round up";
        }

        private void OnResetClick(object? sender, EventArgs e)
        {
            // Reset to default values
            _settings = new SettingsModel();
            LoadSettings();
        }

        private void OnOkClick(object? sender, EventArgs e)
        {
            SaveSettings();
            this.DialogResult = WinForms.DialogResult.OK;
            this.Close();
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