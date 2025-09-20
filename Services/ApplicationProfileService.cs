using System;
using System.IO;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

using JSE_RevitAddin_MEP_OPENINGS.Services;
namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Application-level profile service that manages profile state across the application
    /// </summary>
    public class ApplicationProfileService
    {
        private static ApplicationProfileService? _instance;
        private static readonly object _lock = new object();

        private ProfileManagementService _profileService;
        private readonly StatusManager _statusManager;
        private readonly SettingsService _settingsService;
        private UserProfile? _currentProfile;

        public event EventHandler<ProfileChangedEventArgs>? ProfileChanged;
        public event EventHandler<StatusUpdateEventArgs>? StatusUpdated;

        private ApplicationProfileService()
        {
            // Initialize with a default profile service - will be updated when Revit document is available
            _profileService = new ProfileManagementService("Default");
            _statusManager = new StatusManager();
            _settingsService = new SettingsService();

            // Subscribe to events
            _profileService.ProfileChanged += OnProfileServiceProfileChanged;
            _profileService.StatusUpdated += OnProfileServiceStatusUpdated;
            _statusManager.StatusUpdated += OnStatusManagerStatusUpdated;

            // Load current profile if available
            LoadCurrentProfile();
        }

        public SettingsModel GetCurrentSettings()
        {
            if (CurrentProfile == null)
                return new SettingsModel();

            return _settingsService.GetSettings(CurrentProfile);
        }

        public void SaveCurrentSettings(SettingsModel settings)
        {
            if (CurrentProfile != null)
            {
                _settingsService.SaveSettings(CurrentProfile, settings);
            }
        }

        /// <summary>
        /// Gets the singleton instance of the application profile service
        /// </summary>
        public static ApplicationProfileService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ApplicationProfileService();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Cleans up the singleton instance - call this at the end of command execution
        /// to prevent memory leaks and crashes on subsequent executions
        /// </summary>
        public static void CleanupInstance()
        {
            lock (_lock)
            {
                if (_instance != null)
                {
                    // Unsubscribe from all events to prevent memory leaks
                    if (_instance._profileService != null)
                    {
                        _instance._profileService.ProfileChanged -= _instance.OnProfileServiceProfileChanged;
                        _instance._profileService.StatusUpdated -= _instance.OnProfileServiceStatusUpdated;
                    }
                    
                    if (_instance._statusManager != null)
                    {
                        _instance._statusManager.StatusUpdated -= _instance.OnStatusManagerStatusUpdated;
                    }
                    
                    // Clear current profile reference
                    _instance._currentProfile = null;
                    
                    // Dispose the instance
                    _instance = null;
                }
            }
        }

        /// <summary>
        /// Resets the singleton instance for a new project - call this when starting a new project
        /// to ensure no old profile data persists
        /// </summary>
        public static void ResetForNewProject()
        {
            lock (_lock)
            {
                // Force cleanup of existing instance
                CleanupInstance();
                
                // Create a fresh instance
                _instance = new ApplicationProfileService();
                
                System.Diagnostics.Debug.WriteLine("ApplicationProfileService: Reset for new project - fresh instance created");
                
                // Use conditional logging instead of hardcoded file writes
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_debug.log";
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] ResetForNewProject: Fresh instance created\n");
            }
        }

        /// <summary>
        /// Gets the current active profile
        /// </summary>
        public UserProfile? CurrentProfile => _currentProfile;

        /// <summary>
        /// Gets the profile management service
        /// </summary>
        public ProfileManagementService ProfileService => _profileService;

        /// <summary>
        /// Gets the status manager
        /// </summary>
        public StatusManager StatusManager => _statusManager;

        /// <summary>
        /// Checks if profile setup is required
        /// </summary>
        public bool IsProfileSetupRequired 
        { 
            get 
            {
                // PROJECT-SPECIFIC CHECK: Use the current ProfileManagementService instead of global files
                try
                {
                    // Use conditional logging instead of hardcoded file writes
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_check_debug.log";
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Checking project-specific profiles\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] ProfileService.ProfileFilePath: {_profileService.ProfileFilePath}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Available profiles count: {_profileService.AvailableProfiles.Count}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Current profile: {_currentProfile?.Name ?? "null"}\n");
                    
                    // Check if we have any profiles for this project
                    bool hasProfiles = _profileService.AvailableProfiles.Any();
                    
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Has profiles: {hasProfiles}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Available profiles count: {_profileService.AvailableProfiles.Count}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Current profile: {_currentProfile?.Name ?? "null"}\n");
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] IsProfileSetupRequired: {!hasProfiles}\n");
                    
                    // FIXED LOGIC: Only show Create Profile if NO profiles exist
                    // If profiles exist, show Profile Management instead
                    return !hasProfiles;
                }
                catch (Exception ex)
                {
                    // Fallback to original logic
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_check_debug.log";
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Exception in profile check: {ex.Message}\n");
                    return !_profileService.AvailableProfiles.Any() || _currentProfile == null;
                }
            } 
        }

        /// <summary>
        /// Gets the current user's initials for element tagging
        /// </summary>
        public string CurrentUserInitials
        {
            get
            {
                if (_currentProfile == null)
                    return "UNK"; // Unknown user

                // Extract initials from profile name (e.g., "John_Architect" -> "JA")
                var nameParts = _currentProfile.Name.Split('_', ' ', '-');
                if (nameParts.Length >= 2)
                {
                    return $"{nameParts[0][0]}{nameParts[1][0]}".ToUpper();
                }
                else if (nameParts.Length == 1 && nameParts[0].Length >= 2)
                {
                    return nameParts[0].Substring(0, 2).ToUpper();
                }
                else
                {
                    return _currentProfile.Name.Substring(0, Math.Min(2, _currentProfile.Name.Length)).ToUpper();
                }
            }
        }

        /// <summary>
        /// Gets the current user's discipline for element tagging
        /// </summary>
        public string CurrentUserDiscipline
        {
            get
            {
                if (_currentProfile?.PrimaryDiscipline != null)
                    return _currentProfile.PrimaryDiscipline.Name;
                
                return "Unknown";
            }
        }

        /// <summary>
        /// Sets the current profile
        /// </summary>
        public void SetCurrentProfile(UserProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            _currentProfile = profile;
            _profileService.SetCurrentProfile(profile);
        }

        /// <summary>
        /// Creates a new profile
        /// </summary>
        public UserProfile CreateProfile(string profileName, System.Collections.Generic.List<Discipline> disciplines, string language)
        {
            var profile = _profileService.CreateProfile(profileName, disciplines, language);
            SetCurrentProfile(profile);
            return profile;
        }

        /// <summary>
        /// Loads the current profile from the profile service
        /// </summary>
        private void LoadCurrentProfile()
        {
            try
            {
                // Don't load any profile initially - wait for UpdateForCurrentDocument to set project-specific profile
                _currentProfile = null;
                System.Diagnostics.Debug.WriteLine("LoadCurrentProfile: Cleared current profile - will be set project-specific later");
                
                // Use conditional logging instead of hardcoded file writes
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_debug.log";
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadCurrentProfile: Cleared current profile\n");
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Failed to load current profile: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Checks if the current user has access to a specific discipline
        /// </summary>
        public bool HasDisciplineAccess(string disciplineName)
        {
            if (_currentProfile == null)
                return false;

            return _currentProfile.Disciplines.Any(d => 
                d.IsSelected && d.Name.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if the current user is a primary discipline user
        /// </summary>
        public bool IsPrimaryDisciplineUser()
        {
            return _currentProfile?.PrimaryDiscipline != null;
        }

        /// <summary>
        /// Gets the user's role description
        /// </summary>
        public string GetUserRoleDescription()
        {
            if (_currentProfile == null)
                return "No Profile Set";

            var primary = _currentProfile.PrimaryDiscipline?.Name ?? "No Primary";
            var mepDisciplines = string.Join(", ", _currentProfile.MepDisciplines.Select(d => d.Name));
            
            if (string.IsNullOrEmpty(mepDisciplines))
                return primary;
            
            return $"{primary} + {mepDisciplines}";
        }

        /// <summary>
        /// Gets a formatted user info string for display
        /// </summary>
        public string GetUserInfoString()
        {
            if (_currentProfile == null)
                return "No Profile Active";

            return $"{_currentProfile.Name} ({CurrentUserInitials}) - {GetUserRoleDescription()}";
        }

        /// <summary>
        /// Updates the profile service to use project-specific storage
        /// Call this when a Revit document is available
        /// </summary>
        public void UpdateForCurrentDocument(string? documentPath)
        {
            // TEMP: Remove try-catch to see actual error
            //try
            //{
                if (!string.IsNullOrEmpty(documentPath))
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Creating new ProfileManagementService for path: {documentPath}");
                    
                    // Direct file logging for visibility
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\update_debug.log";
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Creating new ProfileManagementService for path: {documentPath}\n");
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Current service profile file path: {_profileService.ProfileFilePath}");
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Current service available profiles count: {_profileService.AvailableProfiles.Count}");
                    
                    // Create a new profile service with the document-specific path
                    var newProfileService = new ProfileManagementService(documentPath);
                    
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: New service created. Available profiles count: {newProfileService.AvailableProfiles.Count}");
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: New service profile file path: {newProfileService.ProfileFilePath}");
                    
                    // Debug: List all available profiles
                    foreach (var profile in newProfileService.AvailableProfiles)
                    {
                        System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Found profile: {profile.Name}");
                    }
                    
                    // PRESERVE CURRENT PROFILE: Only clear if it's actually a different project
                    var currentProjectPath = _profileService.ProfileFilePath;
                    var newProjectPath = newProfileService.ProfileFilePath;
                    
                    if (currentProjectPath != newProjectPath)
                    {
                        // Different project - clear current profile
                        _currentProfile = null;
                        System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Different project detected - cleared current profile");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Different project - cleared current profile\n");
                    }
                    else
                    {
                        // Same project - preserve current profile
                        System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Same project - preserving current profile: {_currentProfile?.Name ?? "null"}");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Same project - preserving current profile: {_currentProfile?.Name ?? "null"}\n");
                    }
                    
                    // CRITICAL: Try to load saved profile from current_profile.txt
                    try
                    {
                        var projectProfileDir = Path.GetDirectoryName(newProfileService.ProfileFilePath);
                        var currentProfileFile = Path.Combine(projectProfileDir ?? "", "current_profile.txt");
                        
                        if (File.Exists(currentProfileFile))
                        {
                            var savedProfileName = File.ReadAllText(currentProfileFile).Trim();
                            System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Found saved profile name: {savedProfileName}");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Found saved profile name: {savedProfileName}\n");
                            
                            // Find the profile in available profiles
                            var savedProfile = newProfileService.AvailableProfiles.FirstOrDefault(p => p.Name == savedProfileName);
                            if (savedProfile != null)
                            {
                                _currentProfile = savedProfile;
                                System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Successfully loaded saved profile: {savedProfile.Name}");
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Successfully loaded saved profile: {savedProfile.Name}\n");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Saved profile '{savedProfileName}' not found in available profiles");
                                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Saved profile '{savedProfileName}' not found in available profiles\n");
                                
                                // CRITICAL FIX: If profile exists in config but not in XML, create it from config
                                var configFile = Path.Combine(projectProfileDir ?? "", $"config_{savedProfileName}.txt");
                                if (File.Exists(configFile))
                                {
                                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Found config file, recreating profile from config");
                                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Found config file, recreating profile from config\n");
                                    
                                    // Create a basic profile from the config
                                    var recreatedProfile = new UserProfile
                                    {
                                        Id = Guid.NewGuid(),
                                        Name = savedProfileName,
                                        Disciplines = new List<Discipline> { new Discipline("Mechanical", true, "Mechanical systems") },
                                        Language = "English",
                                        CreatedDate = DateTime.Now,
                                        IsActive = true
                                    };
                                    
                                    // Add to available profiles and save
                                    newProfileService.AddProfile(recreatedProfile); // This will create the XML file
                                    
                                    _currentProfile = recreatedProfile;
                                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Recreated and loaded profile: {recreatedProfile.Name}");
                                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Recreated and loaded profile: {recreatedProfile.Name}\n");
                                }
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: No current_profile.txt found at: {currentProfileFile}");
                            JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: No current_profile.txt found at: {currentProfileFile}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Error loading saved profile: {ex.Message}");
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Error loading saved profile: {ex.Message}\n");
                    }
                    
                    // Unsubscribe from old service
                    _profileService.ProfileChanged -= OnProfileServiceProfileChanged;
                    _profileService.StatusUpdated -= OnProfileServiceStatusUpdated;
                    
                    // Replace with new service
                    _profileService = newProfileService;
                    
                    // Subscribe to new service
                    _profileService.ProfileChanged += OnProfileServiceProfileChanged;
                    _profileService.StatusUpdated += OnProfileServiceStatusUpdated;
                    
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Service updated. IsProfileSetupRequired: {IsProfileSetupRequired}");
                }
            //}
            //catch (Exception ex)
            //{
            //    // Log error but don't crash
            //    System.Diagnostics.Debug.WriteLine($"Error updating profile service for document: {ex.Message}");
            //    System.Diagnostics.Debug.WriteLine($"Exception details: {ex}");
            //    
            //    // Also log to file for visibility
            //    try
            //    {
            //        var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "error_debug.log");
            //        File.AppendAllText(logPath, $"[{DateTime.Now}] UpdateForCurrentDocument ERROR: {ex}\n");
            //    }
            //    catch { /* Ignore logging errors */ }
            //}
        }

        /// <summary>
        /// Gets the current Revit document path for project-specific profiles
        /// </summary>
        private string? GetCurrentDocumentPath()
        {
            try
            {
                // In a real Revit context, this would access the active document
                // For now, return null to use the fallback mechanism
                return null;
            }
            catch
            {
                return null;
            }
        }

        private void OnProfileServiceProfileChanged(object? sender, ProfileChangedEventArgs e)
        {
            ProfileChanged?.Invoke(this, e);
        }

        private void OnProfileServiceStatusUpdated(object? sender, StatusUpdateEventArgs e)
        {
            StatusUpdated?.Invoke(this, e);
        }

        private void OnStatusManagerStatusUpdated(object? sender, StatusUpdateEventArgs e)
        {
            StatusUpdated?.Invoke(this, e);
        }

        /// <summary>
        /// Disposes the service and cleans up resources
        /// </summary>
        public void Dispose()
        {
            if (_profileService != null)
            {
                _profileService.ProfileChanged -= OnProfileServiceProfileChanged;
                _profileService.StatusUpdated -= OnProfileServiceStatusUpdated;
            }

            if (_statusManager != null)
            {
                _statusManager.StatusUpdated -= OnStatusManagerStatusUpdated;
            }
        }
        
        /// <summary>
        /// Gets the current active profile
        /// </summary>
        public UserProfile? GetCurrentProfile()
        {
            return _currentProfile;
        }
        
        
        /// <summary>
        /// Saves the current profile to disk
        /// </summary>
        public void SaveCurrentProfile()
        {
            if (_currentProfile != null)
            {
                try
                {
                    // DIRECT FILE SAVING: Bypass broken ProfileManagementService
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveCurrentProfile: Saving profile {_currentProfile.Name}\n");
                    
                    // Save the current profile name to PROJECT-SPECIFIC directory for persistence
                    var projectProfileDir = Path.GetDirectoryName(_profileService.ProfileFilePath);
                    if (!string.IsNullOrEmpty(projectProfileDir) && !Directory.Exists(projectProfileDir))
                        Directory.CreateDirectory(projectProfileDir);
                    
                    var currentProfileFile = Path.Combine(projectProfileDir ?? "", "current_profile.txt");
                    File.WriteAllText(currentProfileFile, _currentProfile.Name);
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Saved current profile name to PROJECT-SPECIFIC: {currentProfileFile}\n");
                    
                    // CRITICAL: Save configuration separately to avoid XML serialization issues
                    try
                    {
                        SaveConfigurationOnly(_currentProfile);
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved successfully to separate file\n");
                    }
                    catch (Exception ex)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to save configuration: {ex.Message}\n");
                    }
                    
                    // Try original method as well (in case it works)
                    try
                    {
                        _profileService.UpdateProfile(_currentProfile, _currentProfile.Name, _currentProfile.Disciplines, _currentProfile.Language);
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Original UpdateProfile method succeeded\n");
                    }
                    catch (Exception ex)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Original UpdateProfile method failed: {ex.Message}\n");
                    }
                    
                    StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                        $"Profile '{_currentProfile.Name}' saved successfully",
                        Models.StatusType.Success,
                        DateTime.Now,
                        "Profile Management"));
                }
                catch (Exception ex)
                {
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveCurrentProfile failed: {ex.Message}\n");
                    
                    StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                        $"Failed to save profile '{_currentProfile.Name}': {ex.Message}",
                        Models.StatusType.Error,
                        DateTime.Now,
                        "Profile Management"));
                }
            }
        }
        
        private void SaveConfigurationOnly(UserProfile profile)
        {
            try
            {
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                
                if (profile.Configuration == null)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] No configuration to save for profile: {profile.Name}\n");
                    return;
                }
                
                // Save configuration to PROJECT-SPECIFIC directory (safer than XML)
                var projectProfileDir = Path.GetDirectoryName(_profileService.ProfileFilePath);
                if (!string.IsNullOrEmpty(projectProfileDir) && !Directory.Exists(projectProfileDir))
                    Directory.CreateDirectory(projectProfileDir);
                
                var configFile = Path.Combine(projectProfileDir ?? "", $"config_{profile.Name}.txt");
                
                // Save as simple text format with each file on a separate line to avoid line break issues
                using (var writer = new StreamWriter(configFile, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine($"ProfileName={profile.Name}");
                    writer.WriteLine($"LastModified={DateTime.Now:O}");
                    
                    // Write each reference file on a separate line
                    writer.WriteLine("SelectedReferenceFiles=");
                    foreach (var file in profile.Configuration.SelectedReferenceFiles ?? new List<string>())
                    {
                        writer.WriteLine($"  {file}");
                    }
                    
                    // Write each host file on a separate line
                    writer.WriteLine("SelectedHostFiles=");
                    foreach (var file in profile.Configuration.SelectedHostFiles ?? new List<string>())
                    {
                        writer.WriteLine($"  {file}");
                    }
                    
                    // Write each MEP category on a separate line
                    writer.WriteLine("SelectedMepCategories=");
                    foreach (var cat in profile.Configuration.SelectedMepCategories ?? new List<string>())
                    {
                        writer.WriteLine($"  {cat}");
                    }
                    
                    // Write each host category on a separate line
                    writer.WriteLine("SelectedHostCategories=");
                    foreach (var cat in profile.Configuration.SelectedHostCategories ?? new List<string>())
                    {
                        writer.WriteLine($"  {cat}");
                    }
                }
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved to: {configFile}\n");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved successfully\n");
            }
            catch (Exception ex)
            {
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveConfigurationOnly failed: {ex.Message}\n");
                throw;
            }
        }
    }
}
