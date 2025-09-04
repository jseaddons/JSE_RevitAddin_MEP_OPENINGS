using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing user profiles with discipline validation
    /// </summary>
    public class ProfileManagementService
    {
        private readonly string _profileDirectory;
        private readonly string _profileFilePath;
        private UserProfile? _currentProfile;
        private readonly List<UserProfile> _availableProfiles;

        public event EventHandler<ProfileChangedEventArgs>? ProfileChanged;
        public event EventHandler<StatusUpdateEventArgs>? StatusUpdated;

        public ProfileManagementService(string? projectPath = null)
        {
            _profileDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                           "JSE_MEP_Openings");
            
            // Make profiles project-specific - use Revit document path if available
            string currentProject;
            if (!string.IsNullOrEmpty(projectPath))
            {
                // Use Revit document name
                currentProject = Path.GetFileNameWithoutExtension(projectPath);
            }
            else
            {
                // Fallback to current directory
                currentProject = Path.GetFileName(Environment.CurrentDirectory) ?? "Default";
            }
            
            _profileFilePath = Path.Combine(_profileDirectory, $"profiles_{currentProject}.xml");
            _availableProfiles = new List<UserProfile>();

            // Ensure directory exists
            if (!Directory.Exists(_profileDirectory))
            {
                Directory.CreateDirectory(_profileDirectory);
            }

            LoadProfiles();
        }

        /// <summary>
        /// Gets the current active profile
        /// </summary>
        public UserProfile? CurrentProfile => _currentProfile;

        /// <summary>
        /// Gets all available profiles
        /// </summary>
        public IReadOnlyList<UserProfile> AvailableProfiles => _availableProfiles.AsReadOnly();

        /// <summary>
        /// Creates a new profile with discipline validation
        /// </summary>
        public UserProfile CreateProfile(string profileName, List<Discipline> disciplines, string language)
        {
            if (string.IsNullOrWhiteSpace(profileName))
                throw new ArgumentException("Profile name cannot be empty", nameof(profileName));

            if (disciplines == null || !disciplines.Any())
                throw new ArgumentException("At least one discipline must be selected", nameof(disciplines));

            // Validate discipline selection
            if (!ValidateDisciplineSelection(disciplines))
            {
                throw new InvalidOperationException("A combination of these disciplines is not allowed. Only one primary discipline can be selected.");
            }

            // Check if profile name already exists
            if (_availableProfiles.Any(p => p.Name.Equals(profileName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Profile '{profileName}' already exists. Please choose a different name.");
            }

            var profile = new UserProfile
            {
                Id = Guid.NewGuid(),
                Name = profileName.Trim(),
                Disciplines = new List<Discipline>(disciplines),
                Language = language,
                CreatedDate = DateTime.Now,
                IsActive = true
            };

            _availableProfiles.Add(profile);
            SaveProfiles();

            StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                $"Profile '{profileName}' created successfully", 
                StatusType.Success, 
                DateTime.Now, 
                "Profile Creation"));

            return profile;
        }

        /// <summary>
        /// Sets the current active profile
        /// </summary>
        public void SetCurrentProfile(UserProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            if (!_availableProfiles.Contains(profile))
                throw new InvalidOperationException("Profile not found in available profiles");

            _currentProfile = profile;
            ProfileChanged?.Invoke(this, new ProfileChangedEventArgs(profile, "Activated"));

            StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                $"Switched to profile '{profile.Name}'", 
                StatusType.Info, 
                DateTime.Now, 
                "Profile Switch"));
        }

        /// <summary>
        /// Updates an existing profile
        /// </summary>
        public void UpdateProfile(UserProfile profile, string newName, List<Discipline> newDisciplines, string newLanguage)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            if (string.IsNullOrWhiteSpace(newName))
                throw new ArgumentException("Profile name cannot be empty", nameof(newName));

            if (newDisciplines == null || !newDisciplines.Any())
                throw new ArgumentException("At least one discipline must be selected", nameof(newDisciplines));

            // Validate discipline selection
            if (!ValidateDisciplineSelection(newDisciplines))
            {
                throw new InvalidOperationException("A combination of these disciplines is not allowed. Only one primary discipline can be selected.");
            }

            // Check if new name conflicts with existing profiles (excluding current profile)
            if (_availableProfiles.Any(p => p.Id != profile.Id && p.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException($"Profile '{newName}' already exists. Please choose a different name.");
            }

            profile.Name = newName.Trim();
            profile.Disciplines = new List<Discipline>(newDisciplines);
            profile.Language = newLanguage;

            SaveProfiles();

            StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                $"Profile '{newName}' updated successfully", 
                StatusType.Success, 
                DateTime.Now, 
                "Profile Update"));
        }

        /// <summary>
        /// Deletes a profile
        /// </summary>
        public void DeleteProfile(UserProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            if (_currentProfile?.Id == profile.Id)
            {
                _currentProfile = null;
            }

            _availableProfiles.Remove(profile);
            SaveProfiles();

            StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                $"Profile '{profile.Name}' deleted", 
                StatusType.Info, 
                DateTime.Now, 
                "Profile Deletion"));
        }

        /// <summary>
        /// Validates discipline selection according to business rules
        /// </summary>
        private bool ValidateDisciplineSelection(List<Discipline> disciplines)
        {
            // Treat the incoming list as the actual selection set
            if (disciplines == null || disciplines.Count == 0)
                return false;

            int primaryCount = disciplines.Count(d => d.IsPrimary);
            int mepCount = disciplines.Count(d => !d.IsPrimary);

            // Only one primary is allowed
            if (primaryCount > 1)
                return false;

            // Mixing primary and MEP is not allowed
            if (primaryCount >= 1 && mepCount >= 1)
                return false;

            // At least one discipline overall
            return primaryCount + mepCount >= 1;
        }

        /// <summary>
        /// Gets predefined discipline options
        /// </summary>
        public List<Discipline> GetPredefinedDisciplines()
        {
            return new List<Discipline>
            {
                // Primary Disciplines (Single Selection)
                new Discipline("Architectural", true, "Architectural design and coordination"),
                new Discipline("Structural", true, "Structural engineering and design"),
                new Discipline("Technical Planner", true, "Technical planning and coordination"),
                new Discipline("Coordination", true, "Project coordination and management"),
                
                // MEP Disciplines (Multi Selection)
                new Discipline("Mechanical", false, "Mechanical systems (HVAC, plumbing)"),
                new Discipline("Electrical", false, "Electrical systems and power"),
                new Discipline("Plumbing", false, "Plumbing and water systems"),
                new Discipline("Fire Protection", false, "Fire protection and safety systems")
            };
        }

        /// <summary>
        /// Gets available languages
        /// </summary>
        public List<string> GetAvailableLanguages()
        {
            return new List<string>
            {
                "English",
                "Spanish",
                "French",
                "German",
                "Italian",
                "Portuguese",
                "Chinese",
                "Japanese",
                "Korean"
            };
        }

        /// <summary>
        /// Loads profiles from file
        /// </summary>
        private void LoadProfiles()
        {
            try
            {
                if (File.Exists(_profileFilePath))
                {
                    var serializer = new XmlSerializer(typeof(List<UserProfile>));
                    using (var reader = new FileStream(_profileFilePath, FileMode.Open))
                    {
                        var profiles = (List<UserProfile>?)serializer.Deserialize(reader);
                        if (profiles != null)
                        {
                            _availableProfiles.Clear();
                            _availableProfiles.AddRange(profiles);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                    $"Failed to load profiles: {ex.Message}", 
                    StatusType.Error, 
                    DateTime.Now, 
                    "Profile Loading"));
            }
        }

        /// <summary>
        /// Saves profiles to file
        /// </summary>
        private void SaveProfiles()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(List<UserProfile>));
                using (var writer = new FileStream(_profileFilePath, FileMode.Create))
                {
                    serializer.Serialize(writer, _availableProfiles);
                }
            }
            catch (Exception ex)
            {
                StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                    $"Failed to save profiles: {ex.Message}", 
                    StatusType.Error, 
                    DateTime.Now, 
                    "Profile Saving"));
            }
        }

        /// <summary>
        /// Checks if a profile setup is required
        /// </summary>
        public bool IsProfileSetupRequired()
        {
            return !_availableProfiles.Any() || _currentProfile == null;
        }

        /// <summary>
        /// Gets the default profile if available
        /// </summary>
        public UserProfile? GetDefaultProfile()
        {
            // First try to get an active profile
            var activeProfile = _availableProfiles.FirstOrDefault(p => p.IsActive);
            if (activeProfile != null)
                return activeProfile;
            
            // If no active profile, return the first available profile
            return _availableProfiles.FirstOrDefault();
        }
    }
}
