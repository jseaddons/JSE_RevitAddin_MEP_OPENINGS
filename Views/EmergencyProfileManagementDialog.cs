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
            // RESTORE CURRENT PROFILE: Load from saved state
            try
            {
                // Use project-specific directory for current profile file
                var profileDir = System.IO.Path.GetDirectoryName(_appProfileService.ProfileService.ProfileFilePath);
                var currentProfileFile = System.IO.Path.Combine(profileDir ?? "", "current_profile.txt");
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_restore_debug.log");
                
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfileData: Checking for saved profile\n");
                
                if (System.IO.File.Exists(currentProfileFile))
                {
                    var savedProfileName = System.IO.File.ReadAllText(currentProfileFile).Trim();
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Found saved profile: {savedProfileName}\n");
                    
                    // If no current profile or different profile, restore the saved one
                    if (_appProfileService.CurrentProfile == null || _appProfileService.CurrentProfile.Name != savedProfileName)
                    {
                        // Create a working profile for the saved name
                        var restoredProfile = new JSE_RevitAddin_MEP_OPENINGS.Models.UserProfile
                        {
                            Id = System.Guid.NewGuid(),
                            Name = savedProfileName,
                            Disciplines = new System.Collections.Generic.List<JSE_RevitAddin_MEP_OPENINGS.Models.Discipline>
                            {
                                new JSE_RevitAddin_MEP_OPENINGS.Models.Discipline("Mechanical", false, "Mechanical systems")
                            },
                            Language = "English",
                            CreatedDate = DateTime.Now,
                            IsActive = true
                        };
                        
                        // Set as current profile using reflection (same as switching)
                        var currentProfileField = typeof(JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService)
                            .GetField("_currentProfile", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        
                        if (currentProfileField != null)
                        {
                            currentProfileField.SetValue(_appProfileService, restoredProfile);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Restored profile via reflection: {savedProfileName}\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_restore_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Profile restoration failed: {ex.Message}\n");
            }

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

            // Load available profiles DIRECTLY from file (bypassing broken ProfileManagementService)
            _profileListBox.Items.Clear();
            
            try
            {
                // Use the project-specific profile file path from the ApplicationProfileService
                var revitProfileFile = _appProfileService.ProfileService.ProfileFilePath;
                
                // Debug logging
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_load_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfileData: Checking project-specific file: {revitProfileFile}\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] File exists: {System.IO.File.Exists(revitProfileFile)}\n");
                
                if (System.IO.File.Exists(revitProfileFile))
                {
                    // SIMPLE TEXT PARSING: Parse profile names directly from XML content to avoid deserialization issues
                    var content = System.IO.File.ReadAllText(revitProfileFile);
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] File content length: {content.Length}\n");
                    
                    // Parse profile names using simple text matching
                    var profileNames = new System.Collections.Generic.List<string>();
                    var lines = content.Split('\n');
                    
                    foreach (var line in lines)
                    {
                        if (line.Trim().StartsWith("<Name>") && line.Trim().EndsWith("</Name>"))
                        {
                            var nameStart = line.IndexOf("<Name>") + 6;
                            var nameEnd = line.IndexOf("</Name>");
                            if (nameStart < nameEnd)
                            {
                                var profileName = line.Substring(nameStart, nameEnd - nameStart);
                                if (!profileNames.Contains(profileName))
                                {
                                    profileNames.Add(profileName);
                                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Found profile name: {profileName}\n");
                                }
                            }
                        }
                    }
                    
                    // Add profiles to list box (load actual profiles from XML)
                    foreach (var profileName in profileNames)
                    {
                        // Load the actual profile from XML instead of creating a fake one
                        var actualProfile = LoadProfileFromFile(profileName);
                        if (actualProfile != null)
                        {
                            var disciplines = string.Join(", ", actualProfile.Disciplines.Select(d => d.Name));
                            var displayText = $"{profileName} - {disciplines}";
                            _profileListBox.Items.Add(new ProfileListItem { Profile = actualProfile, DisplayText = displayText });
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Added actual profile to UI: {profileName}\n");
                        }
                        else
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to load profile: {profileName}\n");
                        }
                    }
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Total profiles added: {profileNames.Count}\n");
                }
                else
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Profile file does not exist\n");
                }
            }
            catch (Exception ex)
            {
                // Fallback to original method if direct loading fails
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_load_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Exception in LoadProfileData: {ex.Message}\n");
                
                // Try original method as fallback
                foreach (var profile in _appProfileService.ProfileService.AvailableProfiles)
                {
                    var disciplines = string.Join(", ", profile.Disciplines.Select(d => d.Name));
                    var displayText = $"{profile.Name} - {disciplines}";
                    _profileListBox.Items.Add(new ProfileListItem { Profile = profile, DisplayText = displayText });
                }
            }

            // Set display member
            _profileListBox.DisplayMember = "DisplayText";
        }

        /// <summary>
        /// Loads a specific profile from the XML file
        /// </summary>
        private JSE_RevitAddin_MEP_OPENINGS.Models.UserProfile? LoadProfileFromFile(string profileName)
        {
            try
            {
                var profileDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                var revitProfileFile = System.IO.Path.Combine(profileDir, "profiles_Revit 2023.xml");
                
                if (!System.IO.File.Exists(revitProfileFile))
                {
                    return null;
                }
                
                // Parse the XML file to find the specific profile
                var content = System.IO.File.ReadAllText(revitProfileFile);
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(System.Collections.Generic.List<JSE_RevitAddin_MEP_OPENINGS.Models.UserProfile>));
                
                using (var reader = new System.IO.StringReader(content))
                {
                    var profiles = (System.Collections.Generic.List<JSE_RevitAddin_MEP_OPENINGS.Models.UserProfile>?)serializer.Deserialize(reader);
                    if (profiles != null)
                    {
                        return profiles.FirstOrDefault(p => p.Name == profileName);
                    }
                }
            }
            catch (Exception ex)
            {
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_load_error.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Error loading profile '{profileName}': {ex.Message}\n");
            }
            
            return null;
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
                    
                    // IMPORTANT: Save the new profile as current to persistence file
                    try
                    {
                        var profileDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                        if (!System.IO.Directory.Exists(profileDir))
                            System.IO.Directory.CreateDirectory(profileDir);
                        
                        var currentProfileFile = System.IO.Path.Combine(profileDir, "current_profile.txt");
                        System.IO.File.WriteAllText(currentProfileFile, emergencySetup.CreatedProfile.Name);
                        
                        var debugLogPath = SafeFileLogger.GetLogFilePath("profile_creation_debug.log");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] New profile created and set as current: {emergencySetup.CreatedProfile.Name}\n");
                    }
                    catch (Exception saveEx)
                    {
                        var debugLogPath = SafeFileLogger.GetLogFilePath("profile_creation_debug.log");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to save new profile as current: {saveEx.Message}\n");
                    }
                    
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
                    // DIRECT PROFILE LOADING: Load the complete profile from XML file
                    var debugLogPath = SafeFileLogger.GetLogFilePath("profile_switch_debug.log");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Switching to profile: {selectedItem.Profile.Name}\n");
                    
                    // SIMPLIFIED APPROACH: Create a working profile and bypass all XML deserialization issues
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Creating functional profile for: {selectedItem.Profile.Name}\n");
                    
                    // Load the actual profile from the XML file instead of creating a fake one
                    var workingProfile = LoadProfileFromFile(selectedItem.Profile.Name);
                    if (workingProfile == null)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to load profile from file: {selectedItem.Profile.Name}\n");
                        WinForms.MessageBox.Show($"Failed to load profile '{selectedItem.Profile.Name}' from file.", 
                            "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                        return;
                    }
                    
                    // DIRECT PROFILE ASSIGNMENT: Bypass ApplicationProfileService validation
                    try
                    {
                        // Use reflection to directly set the current profile
                        var currentProfileField = typeof(JSE_RevitAddin_MEP_OPENINGS.Services.ApplicationProfileService)
                            .GetField("_currentProfile", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        
                        if (currentProfileField != null)
                        {
                            currentProfileField.SetValue(_appProfileService, workingProfile);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Successfully set current profile via reflection: {workingProfile.Name}\n");
                        }
                        else
                        {
                            // Fallback to normal method
                            _appProfileService.SetCurrentProfile(workingProfile);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Set profile using SetCurrentProfile method: {workingProfile.Name}\n");
                        }
                        
                        // IMPORTANT: Save the switched profile as current to persistence file
                        try
                        {
                            var profileDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                            var currentProfileFile = System.IO.Path.Combine(profileDir, "current_profile.txt");
                            System.IO.File.WriteAllText(currentProfileFile, workingProfile.Name);
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Updated current profile file to: {workingProfile.Name}\n");
                        }
                        catch (Exception fileEx)
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to update current profile file: {fileEx.Message}\n");
                        }
                        
                        // Refresh the data
                        LoadProfileData();
                        
                        // Show success message
                        WinForms.MessageBox.Show($"Switched to profile '{workingProfile.Name}'!", 
                            "Success", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    }
                    catch (Exception setEx)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to set profile: {setEx.Message}\n");
                        WinForms.MessageBox.Show($"Failed to switch to profile '{selectedItem.Profile.Name}': {setEx.Message}", 
                            "Error", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
                        return; // Don't proceed to close dialog
                    }
                    
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
