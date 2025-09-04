using System;
using System.Drawing;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using WinForms = System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Views
{
    /// <summary>
    /// Emergency WinForms Profile Management Dialog - NO WPF - CRASH SAFE for Revit
    /// Shows existing profiles and allows switching/creating new ones
    /// </summary>
    public partial class EmergencyProfileManagementDialog : WinForms.Form
    {
        private readonly ApplicationProfileService _appProfileService;
        
        // Controls
        private WinForms.Label _titleLabel = null!;
        private WinForms.Label _currentProfileLabel = null!;
        private WinForms.ListBox _profileListBox = null!;
        private WinForms.Button _switchProfileButton = null!;
        private WinForms.Button _openMainButton = null!;
        private WinForms.Button _newProfileButton = null!;
        private WinForms.Button _closeButton = null!;
        private WinForms.Label _instructionLabel = null!;

        public bool ShouldOpenMainDialog { get; private set; } = false;

        public EmergencyProfileManagementDialog(ApplicationProfileService appProfileService)
        {
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));
            InitializeComponent();
            LoadProfileData();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form properties
            this.Text = "JSE MEP Openings - Profile Management (Emergency WinForms Mode)";
            this.Size = new System.Drawing.Size(500, 400);
            this.StartPosition = WinForms.FormStartPosition.CenterScreen;
            this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int yPos = 20;

            // Title
            _titleLabel = new WinForms.Label
            {
                Text = "Profile Management",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(300, 30),
                AutoSize = false
            };
            this.Controls.Add(_titleLabel);
            yPos += 50;

            // Current Profile Info
            _currentProfileLabel = new WinForms.Label
            {
                Text = "Current Profile: Loading...",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(0, 122, 204),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(450, 40),
                AutoSize = false
            };
            this.Controls.Add(_currentProfileLabel);
            yPos += 50;

            // Available Profiles Label
            var availableLabel = new WinForms.Label
            {
                Text = "Available Profiles:",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.FromArgb(51, 51, 51),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(200, 20),
                AutoSize = false
            };
            this.Controls.Add(availableLabel);
            yPos += 25;

            // Profile ListBox
            _profileListBox = new WinForms.ListBox
            {
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(450, 120),
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular),
                BackColor = System.Drawing.Color.White,
                BorderStyle = WinForms.BorderStyle.FixedSingle
            };
            _profileListBox.SelectedIndexChanged += OnProfileSelectionChanged;
            this.Controls.Add(_profileListBox);
            yPos += 130;

            // Instruction Label
            _instructionLabel = new WinForms.Label
            {
                Text = "Click 'Open Main Interface' to start working, or select a profile and 'Switch Profile' to change.",
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Italic),
                ForeColor = System.Drawing.Color.FromArgb(102, 102, 102),
                Location = new System.Drawing.Point(20, yPos),
                Size = new System.Drawing.Size(450, 30),
                AutoSize = false
            };
            this.Controls.Add(_instructionLabel);
            yPos += 40;

            // Buttons
            int buttonY = yPos;
            int buttonWidth = 110;
            int buttonHeight = 35;
            int buttonSpacing = 120;

            _openMainButton = new WinForms.Button
            {
                Text = "Open Main Interface",
                Location = new System.Drawing.Point(20, buttonY),
                Size = new System.Drawing.Size(buttonWidth + 20, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold)
            };
            _openMainButton.Click += OnOpenMainClick;
            this.Controls.Add(_openMainButton);

            _newProfileButton = new WinForms.Button
            {
                Text = "New Profile",
                Location = new System.Drawing.Point(20 + buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(40, 167, 69),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold)
            };
            _newProfileButton.Click += OnNewProfileClick;
            this.Controls.Add(_newProfileButton);

            _switchProfileButton = new WinForms.Button
            {
                Text = "Switch Profile",
                Location = new System.Drawing.Point(20 + 2 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(255, 193, 7),
                ForeColor = System.Drawing.Color.Black,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold),
                Enabled = false
            };
            _switchProfileButton.Click += OnSwitchProfileClick;
            this.Controls.Add(_switchProfileButton);

            _closeButton = new WinForms.Button
            {
                Text = "Close",
                Location = new System.Drawing.Point(20 + 3 * buttonSpacing, buttonY),
                Size = new System.Drawing.Size(buttonWidth, buttonHeight),
                BackColor = System.Drawing.Color.FromArgb(108, 117, 125),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = WinForms.FlatStyle.Flat,
                Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold)
            };
            _closeButton.Click += OnCloseClick;
            this.Controls.Add(_closeButton);

            this.ResumeLayout(false);
        }

        private void LoadProfileData()
        {
            // Update current profile info
            if (_appProfileService.CurrentProfile != null)
            {
                var disciplines = string.Join(", ", _appProfileService.CurrentProfile.Disciplines.Select(d => d.Name));
                _currentProfileLabel.Text = $"Current Profile: {_appProfileService.CurrentProfile.Name}\nDisciplines: {disciplines}";
            }
            else
            {
                _currentProfileLabel.Text = "Current Profile: None selected";
            }

            // Load available profiles
            _profileListBox.Items.Clear();
            foreach (var profile in _appProfileService.ProfileService.AvailableProfiles)
            {
                var disciplines = string.Join(", ", profile.Disciplines.Select(d => d.Name));
                var displayText = $"{profile.Name} - {disciplines}";
                _profileListBox.Items.Add(new ProfileListItem { Profile = profile, DisplayText = displayText });
            }

            // Set display member
            _profileListBox.DisplayMember = "DisplayText";
        }

        private void OnProfileSelectionChanged(object sender, EventArgs e)
        {
            _switchProfileButton.Enabled = _profileListBox.SelectedItem != null;
        }

        private void OnNewProfileClick(object sender, EventArgs e)
        {
            try
            {
                // EMERGENCY MODE: Use WinForms ProfileSetup instead of WPF
                var emergencySetup = new EmergencyProfileSetup(_appProfileService.ProfileService, _appProfileService.StatusManager);
                
                var result = emergencySetup.ShowDialog();
                if (result == WinForms.DialogResult.OK && emergencySetup.CreatedProfile != null)
                {
                    // Set the new profile as current
                    _appProfileService.SetCurrentProfile(emergencySetup.CreatedProfile);
                    
                    // Refresh the data
                    LoadProfileData();
                    
                    // Show success message
                    WinForms.MessageBox.Show($"Profile '{emergencySetup.CreatedProfile.Name}' created and activated!", 
                        "Success", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    
                    // Set flag to open main dialog
                    ShouldOpenMainDialog = true;
                    this.DialogResult = WinForms.DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error creating profile: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnSwitchProfileClick(object sender, EventArgs e)
        {
            try
            {
                if (_profileListBox.SelectedItem is ProfileListItem selectedItem)
                {
                    // Switch to selected profile
                    _appProfileService.SetCurrentProfile(selectedItem.Profile);
                    
                    // Refresh the data
                    LoadProfileData();
                    
                    // Show success message
                    WinForms.MessageBox.Show($"Switched to profile '{selectedItem.Profile.Name}'!", 
                        "Success", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    
                    // Set flag to open main dialog
                    ShouldOpenMainDialog = true;
                    this.DialogResult = WinForms.DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error switching profile: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnOpenMainClick(object sender, EventArgs e)
        {
            try
            {
                // Set flag to open main dialog
                ShouldOpenMainDialog = true;
                this.DialogResult = WinForms.DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                WinForms.MessageBox.Show($"Error opening main interface: {ex.Message}", 
                    "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
        }

        private void OnCloseClick(object sender, EventArgs e)
        {
            this.DialogResult = WinForms.DialogResult.Cancel;
            this.Close();
        }

        private class ProfileListItem
        {
            public Models.UserProfile Profile { get; set; } = null!;
            public string DisplayText { get; set; } = string.Empty;
        }
    }
}
