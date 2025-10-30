using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Emergency profile setup using WinForms - NO WPF BINDINGS - CRASH SAFE
    /// </summary>
    public partial class EmergencyProfileSetup : WinForms.Form
    {
        private readonly ProfileManagementService _profileService;
        private readonly StatusManager _statusManager;

        // Controls
        private WinForms.TextBox _profileNameTextBox = null!;
        private WinForms.RadioButton _selectNoneRadio = null!, _archRadio = null!, _structRadio = null!, _techRadio = null!, _coordRadio = null!;
        private WinForms.CheckBox _mechCheck = null!, _elecCheck = null!, _plumbCheck = null!, _fireCheck = null!;
        private WinForms.Label _validationLabel = null!;
        private WinForms.Button _okButton = null!, _cancelButton = null!;
        private WinForms.GroupBox _primaryGroupBox = null!;

        public UserProfile? CreatedProfile { get; private set; }
        public bool WasSuccessful { get; private set; }

        public EmergencyProfileSetup(ProfileManagementService profileService, StatusManager statusManager)
        {
            _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
            _statusManager = statusManager ?? throw new ArgumentNullException(nameof(statusManager));

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = "JSE MEP Openings - Profile Setup (Emergency Mode)";
            this.Size = new System.Drawing.Size(500, 700);
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int yPos = 20;

            // Title
            var titleLabel = new WinForms.Label
            {
                Text = "Profile Setup",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(200, 30)
            };
            this.Controls.Add(titleLabel);
            yPos += 50;

            // Profile Name
            var nameLabel = new WinForms.Label
            {
                Text = "Profile Name:",
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(100, 20)
            };
            this.Controls.Add(nameLabel);

            _profileNameTextBox = new WinForms.TextBox
            {
                Location = new System.Drawing.Point(130, yPos),
                Size = new System.Drawing.Size(300, 20)
            };
            _profileNameTextBox.TextChanged += OnInputChanged;
            this.Controls.Add(_profileNameTextBox);
            yPos += 40;

            // Primary Disciplines with GroupBox for proper radio button grouping
            var primaryLabel = new WinForms.Label
            {
                Text = "Primary Discipline (Select ONE or None):",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(400, 20)
            };
            this.Controls.Add(primaryLabel);
            yPos += 25;

            _primaryGroupBox = new WinForms.GroupBox
            {
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(450, 140),
                Text = ""
            };
            this.Controls.Add(_primaryGroupBox);

            int radioYPos = 20;

            _selectNoneRadio = new WinForms.RadioButton
            {
                Text = "Select None",
                Location = new System.Drawing.Point(20, radioYPos),
                Size = new System.Drawing.Size(150, 20),
                Checked = true // Default selection
            };
            _selectNoneRadio.CheckedChanged += OnPrimaryRadioChanged;
            _primaryGroupBox.Controls.Add(_selectNoneRadio);
            radioYPos += 25;

            _archRadio = new WinForms.RadioButton
            {
                Text = "Architectural",
                Location = new System.Drawing.Point(20, radioYPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _archRadio.CheckedChanged += OnPrimaryRadioChanged;
            _primaryGroupBox.Controls.Add(_archRadio);
            radioYPos += 25;

            _structRadio = new WinForms.RadioButton
            {
                Text = "Structural",
                Location = new System.Drawing.Point(20, radioYPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _structRadio.CheckedChanged += OnPrimaryRadioChanged;
            _primaryGroupBox.Controls.Add(_structRadio);
            radioYPos += 25;

            _techRadio = new WinForms.RadioButton
            {
                Text = "Technical Planner",
                Location = new System.Drawing.Point(20, radioYPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _techRadio.CheckedChanged += OnPrimaryRadioChanged;
            _primaryGroupBox.Controls.Add(_techRadio);
            radioYPos += 25;

            _coordRadio = new WinForms.RadioButton
            {
                Text = "Coordination",
                Location = new System.Drawing.Point(20, radioYPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _coordRadio.CheckedChanged += OnPrimaryRadioChanged;
            _primaryGroupBox.Controls.Add(_coordRadio);
            
            yPos += 150;

            // MEP Disciplines
            var mepLabel = new WinForms.Label
            {
                Text = "MEP Disciplines (Select multiple OR none if primary selected):",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(400, 20)
            };
            this.Controls.Add(mepLabel);
            yPos += 25;

            _mechCheck = new WinForms.CheckBox
            {
                Text = "Mechanical",
                Location = new System.Drawing.Point(40, yPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _mechCheck.CheckedChanged += OnMepCheckChanged;
            this.Controls.Add(_mechCheck);
            yPos += 25;

            _elecCheck = new WinForms.CheckBox
            {
                Text = "Electrical",
                Location = new System.Drawing.Point(40, yPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _elecCheck.CheckedChanged += OnMepCheckChanged;
            this.Controls.Add(_elecCheck);
            yPos += 25;

            _plumbCheck = new WinForms.CheckBox
            {
                Text = "Plumbing",
                Location = new System.Drawing.Point(40, yPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _plumbCheck.CheckedChanged += OnMepCheckChanged;
            this.Controls.Add(_plumbCheck);
            yPos += 25;

            _fireCheck = new WinForms.CheckBox
            {
                Text = "Fire Fighting",
                Location = new System.Drawing.Point(40, yPos),
                Size = new System.Drawing.Size(150, 20)
            };
            _fireCheck.CheckedChanged += OnMepCheckChanged;
            this.Controls.Add(_fireCheck);
            yPos += 40;

            // Validation message
            _validationLabel = new WinForms.Label
            {
                Text = "",
                ForeColor = System.Drawing.Color.Red,
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(400, 40)
            };
            this.Controls.Add(_validationLabel);
            yPos += 50;

            // Buttons
            _okButton = new WinForms.Button
            {
                Text = "Create Profile",
                Location = new System.Drawing.Point(250, yPos),
                Size = new System.Drawing.Size(100, 30),
                Enabled = false
            };
            _okButton.Click += OnOkClick;
            this.Controls.Add(_okButton);

            _cancelButton = new WinForms.Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(360, yPos),
                Size = new System.Drawing.Size(80, 30)
            };
            _cancelButton.Click += OnCancelClick;
            this.Controls.Add(_cancelButton);

            this.ResumeLayout(false);

            // Initialize MEP checkboxes as enabled since "Select None" is default
            _mechCheck.Enabled = true;
            _elecCheck.Enabled = true;
            _plumbCheck.Enabled = true;
            _fireCheck.Enabled = true;

            // Initial validation
            ValidateInput();
        }

        private void OnPrimaryRadioChanged(object sender, EventArgs e)
        {
            try
            {
                var radio = sender as WinForms.RadioButton;
                if (radio != null && radio.Checked)
                {
                    if (radio == _selectNoneRadio)
                    {
                        // "Select None" chosen - enable MEP options
                        _mechCheck.Enabled = true;
                        _elecCheck.Enabled = true;
                        _plumbCheck.Enabled = true;
                        _fireCheck.Enabled = true;
                    }
                    else
                    {
                        // A primary discipline chosen - clear and disable MEP options
                        _mechCheck.Checked = false;
                        _elecCheck.Checked = false;
                        _plumbCheck.Checked = false;
                        _fireCheck.Checked = false;
                        
                        _mechCheck.Enabled = false;
                        _elecCheck.Enabled = false;
                        _plumbCheck.Enabled = false;
                        _fireCheck.Enabled = false;
                    }
                }
                
                ValidateInput();
            }
            catch (Exception ex)
            {
                _validationLabel.Text = $"Validation error: {ex.Message}";
                _okButton.Enabled = false;
            }
        }

        private void OnMepCheckChanged(object sender, EventArgs e)
        {
            try
            {
                var anyMepChecked = _mechCheck.Checked || _elecCheck.Checked || _plumbCheck.Checked || _fireCheck.Checked;
                
                if (anyMepChecked)
                {
                    // Clear and disable primary radio buttons when MEP is selected (except Select None)
                    _archRadio.Checked = false;
                    _structRadio.Checked = false;
                    _techRadio.Checked = false;
                    _coordRadio.Checked = false;
                    _selectNoneRadio.Checked = true; // Set Select None as default
                    
                    // Gray out primary options (but keep Select None enabled)
                    _archRadio.Enabled = false;
                    _structRadio.Enabled = false;
                    _techRadio.Enabled = false;
                    _coordRadio.Enabled = false;
                    _selectNoneRadio.Enabled = true;
                }
                else
                {
                    // Re-enable primary options when no MEP selected
                    _archRadio.Enabled = true;
                    _structRadio.Enabled = true;
                    _techRadio.Enabled = true;
                    _coordRadio.Enabled = true;
                    _selectNoneRadio.Enabled = true;
                }
                
                ValidateInput();
            }
            catch (Exception ex)
            {
                _validationLabel.Text = $"Validation error: {ex.Message}";
                _okButton.Enabled = false;
            }
        }

        private void OnInputChanged(object sender, EventArgs e)
        {
            try
            {
                ValidateInput();
            }
            catch (Exception ex)
            {
                _validationLabel.Text = $"Validation error: {ex.Message}";
                _okButton.Enabled = false;
            }
        }

        private void ValidateInput()
        {
            try
            {
                var profileName = _profileNameTextBox.Text?.Trim() ?? "";
                var primarySelected = _archRadio.Checked || _structRadio.Checked || _techRadio.Checked || _coordRadio.Checked;
                var mepSelected = _mechCheck.Checked || _elecCheck.Checked || _plumbCheck.Checked || _fireCheck.Checked;
                var selectNoneChosen = _selectNoneRadio.Checked;

                if (string.IsNullOrWhiteSpace(profileName))
                {
                    _validationLabel.Text = "Profile name is required";
                    _okButton.Enabled = false;
                    return;
                }

                // Check if profile name already exists (with debug info)
                var existingProfile = _profileService.AvailableProfiles.FirstOrDefault(p => p.Name.Equals(profileName, StringComparison.OrdinalIgnoreCase));
                if (existingProfile != null)
                {
                    var currentDir = Environment.CurrentDirectory;
                    var projectName = System.IO.Path.GetFileName(currentDir);
                    _validationLabel.Text = $"Profile '{profileName}' exists in project '{projectName}'. Dir: {currentDir}";
                    _okButton.Enabled = false;
                    return;
                }

                if (selectNoneChosen && !mepSelected)
                {
                    _validationLabel.Text = "If 'Select None' is chosen for Primary, you must select at least one MEP discipline";
                    _okButton.Enabled = false;
                    return;
                }

                if (!primarySelected && !mepSelected && !selectNoneChosen)
                {
                    _validationLabel.Text = "Please select at least one discipline";
                    _okButton.Enabled = false;
                    return;
                }

                if (primarySelected && mepSelected)
                {
                    _validationLabel.Text = "Choose either ONE primary discipline OR one/more MEP disciplines";
                    _okButton.Enabled = false;
                    return;
                }

                // Valid
                _validationLabel.Text = "";
                _okButton.Enabled = true;
            }
            catch (Exception ex)
            {
                _validationLabel.Text = $"Validation error: {ex.Message}";
                _okButton.Enabled = false;
            }
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            try
            {
                var disciplines = new List<Discipline>();
                var profileName = _profileNameTextBox.Text.Trim();

                // Primary disciplines
                if (_archRadio.Checked)
                    disciplines.Add(new Discipline("Architectural", true, "Architectural design and coordination"));
                if (_structRadio.Checked)
                    disciplines.Add(new Discipline("Structural", true, "Structural engineering and design"));
                if (_techRadio.Checked)
                    disciplines.Add(new Discipline("Technical Planner", true, "Technical planning and coordination"));
                if (_coordRadio.Checked)
                    disciplines.Add(new Discipline("Coordination", true, "Project coordination and management"));

                // MEP disciplines
                if (_mechCheck.Checked)
                    disciplines.Add(new Discipline("Mechanical", false, "Mechanical systems (HVAC, plumbing)"));
                if (_elecCheck.Checked)
                    disciplines.Add(new Discipline("Electrical", false, "Electrical systems and power"));
                if (_plumbCheck.Checked)
                    disciplines.Add(new Discipline("Plumbing", false, "Plumbing and water systems"));
                if (_fireCheck.Checked)
                    disciplines.Add(new Discipline("Fire Fighting", false, "Fire fighting and safety systems"));

                // Create profile
                CreatedProfile = _profileService.CreateProfile(profileName, disciplines, "English");
                WasSuccessful = true;

                WinForms.MessageBox.Show($"Profile '{profileName}' created successfully!", "Success", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error creating profile: {ex.Message}", "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnCancelClick(object sender, EventArgs e)
        {
            this.DialogResult = WinForms.DialogResult.Cancel;
            this.Close();
        }
    }
}
