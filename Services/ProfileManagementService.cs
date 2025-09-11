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
            // Make profiles truly project-specific
            string currentProject;
            if (!string.IsNullOrEmpty(projectPath))
            {
                // Use the actual project path to create a unique identifier
                // Hash the project path to create a stable, unique identifier
                var projectHash = projectPath.GetHashCode().ToString("X8");
                currentProject = $"Project_{projectHash}";
                
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Creating project-specific directory for path: {projectPath}");
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Project identifier: {currentProject}");
            }
            else
            {
                currentProject = "Default"; // For testing without project path
            }
            
            // Create project-specific directory in the project's directory, not global AppData
            if (!string.IsNullOrEmpty(projectPath))
            {
                // Create profiles directory in the same directory as the project file
                var projectDir = Path.GetDirectoryName(projectPath);
                _profileDirectory = Path.Combine(projectDir ?? "", "JSE_MEP_Profiles");
            }
            else
            {
                // Fallback to AppData only for testing
                _profileDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                                               "JSE_MEP_Openings", currentProject);
            }
            
            _profileFilePath = Path.Combine(_profileDirectory, $"profiles_{currentProject}.xml");
            _availableProfiles = new List<UserProfile>();

            // Ensure directory exists
            if (!Directory.Exists(_profileDirectory))
            {
                Directory.CreateDirectory(_profileDirectory);
            }

            LoadProfiles();
            
            // If no profiles were loaded (XML failed), try fallback immediately
            if (_availableProfiles.Count == 0)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"ProfileManagementService constructor: No profiles loaded, attempting fallback from config files");
                    LoadProfilesFromConfigFiles();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ProfileManagementService constructor: Fallback failed - {ex.Message}");
                }
            }
            
            // Debug: Log profile loading results
            System.Diagnostics.Debug.WriteLine($"ProfileManagementService constructor: Profile file path: {_profileFilePath}");
            System.Diagnostics.Debug.WriteLine($"ProfileManagementService constructor: Loaded {_availableProfiles.Count} profiles");
            foreach (var profile in _availableProfiles)
            {
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService constructor: Loaded profile: {profile.Name}");
            }
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
        /// Gets the profile file path
        /// </summary>
        public string ProfileFilePath => _profileFilePath;

        /// <summary>
        /// Creates a new profile with discipline validation
        /// </summary>
        public UserProfile CreateProfile(string profileName, List<Discipline> disciplines, string language)
        {
            System.Diagnostics.Debug.WriteLine($"CreateProfile: Creating profile '{profileName}' with {disciplines?.Count ?? 0} disciplines");
            
            // Log to file for debugging
            var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
            File.AppendAllText(debugLogPath, $"[{DateTime.Now}] CreateProfile: Creating profile '{profileName}' with {disciplines?.Count ?? 0} disciplines\n");
            
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
            System.Diagnostics.Debug.WriteLine($"CreateProfile: Added profile to collection, now have {_availableProfiles.Count} profiles");
            File.AppendAllText(debugLogPath, $"[{DateTime.Now}] CreateProfile: Added profile to collection, now have {_availableProfiles.Count} profiles\n");
            
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
                System.Diagnostics.Debug.WriteLine($"LoadProfiles: Checking file: {_profileFilePath}");
                System.Diagnostics.Debug.WriteLine($"LoadProfiles: File exists: {File.Exists(_profileFilePath)}");
                
                // Log to file for debugging
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Checking file: {_profileFilePath}\n");
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: File exists: {File.Exists(_profileFilePath)}\n");
                
                if (File.Exists(_profileFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"LoadProfiles: File size: {new FileInfo(_profileFilePath).Length} bytes");
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: File size: {new FileInfo(_profileFilePath).Length} bytes\n");
                    
                    var serializer = new XmlSerializer(typeof(List<UserProfile>));
                    using (var reader = new FileStream(_profileFilePath, FileMode.Open))
                    {
                        var profiles = (List<UserProfile>?)serializer.Deserialize(reader);
                        if (profiles != null)
                        {
                            _availableProfiles.Clear();
                            _availableProfiles.AddRange(profiles);
                            System.Diagnostics.Debug.WriteLine($"LoadProfiles: Successfully loaded {profiles.Count} profiles");
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Successfully loaded {profiles.Count} profiles\n");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"LoadProfiles: Deserialized profiles is null");
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Deserialized profiles is null\n");
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"LoadProfiles: Profile file does not exist");
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Profile file does not exist\n");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadProfiles: ERROR - {ex.Message}");
                
                // Log to file for debugging
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: ERROR - {ex.Message}\n");
                
                // FALLBACK: Try to load profiles from config files when XML fails
                try
                {
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Attempting fallback from config files\n");
                    LoadProfilesFromConfigFiles();
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Fallback successful - loaded {_availableProfiles.Count} profiles\n");
                }
                catch (Exception fallbackEx)
                {
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfiles: Fallback also failed - {fallbackEx.Message}\n");
                    StatusUpdated?.Invoke(this, new StatusUpdateEventArgs(
                        $"Failed to load profiles: {ex.Message}", 
                        StatusType.Error, 
                        DateTime.Now, 
                        "Profile Loading"));
                }
            }
        }

        /// <summary>
        /// Fallback method to load profiles from config files when XML serialization fails
        /// </summary>
        private void LoadProfilesFromConfigFiles()
        {
            try
            {
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Starting fallback load\n");
                
                _availableProfiles.Clear();
                
                // Get the directory containing the profile files
                var profileDir = Path.GetDirectoryName(_profileFilePath);
                if (string.IsNullOrEmpty(profileDir) || !Directory.Exists(profileDir))
                {
                    File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Profile directory not found\n");
                    return;
                }
                
                // FIXED: Only look for config files in the CURRENT PROJECT directory
                // This ensures profiles are truly project-specific
                var configFiles = Directory.GetFiles(profileDir, "config_*.txt");
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Looking in project-specific directory: {profileDir}\n");
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Found {configFiles.Length} config files\n");
                
                foreach (var configFile in configFiles)
                {
                    try
                    {
                        var fileName = Path.GetFileNameWithoutExtension(configFile);
                        var profileName = fileName.Replace("config_", "");
                        
                        // Create a basic profile from the config file
                        var profile = new UserProfile
                        {
                            Id = Guid.NewGuid(),
                            Name = profileName,
                            Disciplines = new List<Discipline> { new Discipline("Mechanical", true, "Mechanical systems") },
                            Language = "English",
                            CreatedDate = DateTime.Now,
                            IsActive = true
                        };
                        
                        _availableProfiles.Add(profile);
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Created profile '{profileName}' from config file\n");
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Error processing config file {configFile}: {ex.Message}\n");
                    }
                }
                
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Successfully loaded {_availableProfiles.Count} profiles from config files\n");
            }
            catch (Exception ex)
            {
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: ERROR - {ex.Message}\n");
                throw;
            }
        }

        /// <summary>
        /// Adds a profile to the available profiles and saves to file
        /// </summary>
        public void AddProfile(UserProfile profile)
        {
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            _availableProfiles.Add(profile);
            SaveProfiles();
        }

        /// <summary>
        /// Saves profiles to file
        /// </summary>
        public void SaveProfiles()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"SaveProfiles: Saving {_availableProfiles.Count} profiles to: {_profileFilePath}");
                
                // Log to file for debugging
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: Saving {_availableProfiles.Count} profiles to: {_profileFilePath}\n");
                
                var serializer = new XmlSerializer(typeof(List<UserProfile>));
                using (var writer = new FileStream(_profileFilePath, FileMode.Create))
                {
                    serializer.Serialize(writer, _availableProfiles);
                }
                
                System.Diagnostics.Debug.WriteLine($"SaveProfiles: Successfully saved profiles to XML file");
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: Successfully saved profiles to XML file\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SaveProfiles: ERROR - {ex.Message}");
                
                // Log to file for debugging
                var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                File.AppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: ERROR - {ex.Message}\n");
                
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
