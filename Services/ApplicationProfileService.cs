using System;
using System.IO;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

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
        private UserProfile? _currentProfile;

        public event EventHandler<ProfileChangedEventArgs>? ProfileChanged;
        public event EventHandler<StatusUpdateEventArgs>? StatusUpdated;

        private ApplicationProfileService()
        {
            // Initialize with a default profile service - will be updated when Revit document is available
            _profileService = new ProfileManagementService("Default");
            _statusManager = new StatusManager();

            // Subscribe to events
            _profileService.ProfileChanged += OnProfileServiceProfileChanged;
            _profileService.StatusUpdated += OnProfileServiceStatusUpdated;
            _statusManager.StatusUpdated += OnStatusManagerStatusUpdated;

            // Load current profile if available
            LoadCurrentProfile();
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
                // DIRECT FILE CHECK: Bypass ProfileManagementService and check profile files directly
                try
                {
                    var profileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                    var revitProfileFile = Path.Combine(profileDir, "profiles_Revit 2023.xml");
                    
                    // Log this check
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_check_debug.log";
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Checking profile file: {revitProfileFile}\n");
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] File exists: {File.Exists(revitProfileFile)}\n");
                    
                    if (File.Exists(revitProfileFile))
                    {
                        var content = File.ReadAllText(revitProfileFile);
                        bool hasProfiles = content.Contains("<Name>ram</Name>") || content.Contains("<Name>mech</Name>") || content.Contains("<UserProfile>");
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] File has profiles: {hasProfiles}\n");
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] IsProfileSetupRequired: {!hasProfiles}\n");
                        return !hasProfiles;
                    }
                    
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] File doesn't exist, setup required: True\n");
                    return true;
                }
                catch (Exception ex)
                {
                    // Fallback to original logic
                    var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_check_debug.log";
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Exception in profile check: {ex.Message}\n");
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
                var defaultProfile = _profileService.GetDefaultProfile();
                if (defaultProfile != null)
                {
                    _currentProfile = defaultProfile;
                }
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
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Creating new ProfileManagementService for path: {documentPath}\n");
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
                    
                    // Transfer current profile if it exists
                    if (_currentProfile != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Transferring current profile: {_currentProfile.Name}");
                        newProfileService.SetCurrentProfile(_currentProfile);
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
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveCurrentProfile: Saving profile {_currentProfile.Name}\n");
                    
                    // Also save the current profile name to a simple text file for persistence
                    var profileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                    if (!Directory.Exists(profileDir))
                        Directory.CreateDirectory(profileDir);
                    
                    var currentProfileFile = Path.Combine(profileDir, "current_profile.txt");
                    File.WriteAllText(currentProfileFile, _currentProfile.Name);
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Saved current profile name to: {currentProfileFile}\n");
                    
                    // CRITICAL: Save configuration separately to avoid XML serialization issues
                    try
                    {
                        SaveConfigurationOnly(_currentProfile);
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved successfully to separate file\n");
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to save configuration: {ex.Message}\n");
                    }
                    
                    // Try original method as well (in case it works)
                    try
                    {
                        _profileService.UpdateProfile(_currentProfile, _currentProfile.Name, _currentProfile.Disciplines, _currentProfile.Language);
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Original UpdateProfile method succeeded\n");
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Original UpdateProfile method failed: {ex.Message}\n");
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
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveCurrentProfile failed: {ex.Message}\n");
                    
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
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] No configuration to save for profile: {profile.Name}\n");
                    return;
                }
                
                // Save configuration to a simple JSON file (safer than XML)
                var profileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                if (!Directory.Exists(profileDir))
                    Directory.CreateDirectory(profileDir);
                
                var configFile = Path.Combine(profileDir, $"config_{profile.Name}.txt");
                
                // Convert configuration to simple key-value pairs
                var configData = new Dictionary<string, object>
                {
                    ["ProfileName"] = profile.Name,
                    ["LastModified"] = DateTime.Now.ToString("O"),
                    ["SelectedReferenceFiles"] = profile.Configuration.SelectedReferenceFiles ?? new List<string>(),
                    ["SelectedHostFiles"] = profile.Configuration.SelectedHostFiles ?? new List<string>(),
                    ["SelectedMepCategories"] = profile.Configuration.SelectedMepCategories ?? new List<string>(),
                    ["SelectedHostCategories"] = profile.Configuration.SelectedHostCategories ?? new List<string>()
                };
                
                // Save as simple text format (much safer than XML serialization)
                var lines = new List<string>
                {
                    $"ProfileName={profile.Name}",
                    $"LastModified={DateTime.Now:O}",
                    $"SelectedReferenceFiles={string.Join("|", profile.Configuration.SelectedReferenceFiles ?? new List<string>())}",
                    $"SelectedHostFiles={string.Join("|", profile.Configuration.SelectedHostFiles ?? new List<string>())}",
                    $"SelectedMepCategories={string.Join("|", profile.Configuration.SelectedMepCategories ?? new List<string>())}",
                    $"SelectedHostCategories={string.Join("|", profile.Configuration.SelectedHostCategories ?? new List<string>())}"
                };
                
                File.WriteAllLines(configFile, lines);
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved to: {configFile}\n");
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] Saved {configData.Count} configuration settings\n");
            }
            catch (Exception ex)
            {
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveConfigurationOnly failed: {ex.Message}\n");
                throw;
            }
        }
    }
}
