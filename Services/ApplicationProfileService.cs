using System;
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
        public bool IsProfileSetupRequired => !_profileService.AvailableProfiles.Any() || _currentProfile == null;

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
            try
            {
                if (!string.IsNullOrEmpty(documentPath))
                {
                    // Create a new profile service with the document-specific path
                    var newProfileService = new ProfileManagementService(documentPath);
                    
                    // Transfer current profile if it exists
                    if (_currentProfile != null)
                    {
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
                }
            }
            catch (Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error updating profile service for document: {ex.Message}");
            }
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
    }
}
