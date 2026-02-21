using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

using JSE_RevitAddin_MEP_OPENINGS.Services;
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

        public ProfileManagementService(string? projectRootDirectory = null)
        {
            // Make profiles truly project-specific
            // User Request: Use AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName] directly
            
            string currentProjectName = "Default";

            if (!string.IsNullOrEmpty(projectRootDirectory))
            {
                // ✅ CRITICAL FIX: Ensure path is absolute. If relative (e.g. "Default"), make it absolute under AppData.
                if (!Path.IsPathRooted(projectRootDirectory))
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    _profileDirectory = Path.Combine(appData, "JSE_MEP_Openings", "Projects", projectRootDirectory);
                    System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Converted relative path '{projectRootDirectory}' to absolute: {_profileDirectory}");
                }
                else
                {
                    _profileDirectory = projectRootDirectory;
                }
                
                // Extract project name from the directory path for the filename
                // Path is ...\Projects\[ProjectName]
                currentProjectName = new DirectoryInfo(_profileDirectory).Name;
                
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Using project directory: {_profileDirectory}");
            }
            else
            {
                // Fallback for testing/default
                 var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                _profileDirectory = Path.Combine(appData, "JSE_MEP_Openings", "Default");
                currentProjectName = "Default";
            }
            
            // Generate filename: profiles_[ProjectName].json
            _profileFilePath = Path.Combine(_profileDirectory, $"profiles_{currentProjectName}.json");
            _availableProfiles = new List<UserProfile>();

            // ✅ CRITICAL FIX: Ensure all parent directories are created step-by-step with proper error handling
            // This fixes issues where users with 3-digit IDs (jse***) can't create the folder
            try
            {
                EnsureDirectoryStructure(_profileDirectory);
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Directory verified/created: {_profileDirectory}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: ERROR creating directory {_profileDirectory}: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Stack trace: {ex.StackTrace}");
                // Don't throw - allow the service to continue with a fallback path or empty profile list
                // This prevents the UI from failing to appear
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Continuing with potentially invalid directory path");
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
        /// ✅ CRITICAL FIX: Ensures directory structure is created step-by-step
        /// This fixes issues where Directory.CreateDirectory() fails silently for some users (e.g., 3-digit user IDs)
        /// Creates: AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName] (or AppData\Roaming\JSE_MEP_Openings\Default)
        /// </summary>
        private void EnsureDirectoryStructure(string targetDirectory)
        {
            if (string.IsNullOrEmpty(targetDirectory))
            {
                System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Target directory is null/empty, skipping creation");
                return;
            }

            try
            {
                // Get AppData base path first
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                
                // Step 1: Ensure base JSE_MEP_Openings folder exists
                var baseFolder = Path.Combine(appData, "JSE_MEP_Openings");
                if (!Directory.Exists(baseFolder))
                {
                    try
                    {
                        Directory.CreateDirectory(baseFolder);
                        System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Created base folder: {baseFolder}");
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Cannot create base folder {baseFolder}. User may lack permissions. Error: {ex.Message}", ex);
                    }
                }

                // Step 2: If target is under Projects subfolder, ensure Projects folder exists
                if (targetDirectory.Contains("Projects", StringComparison.OrdinalIgnoreCase))
                {
                    var projectsFolder = Path.Combine(baseFolder, "Projects");
                    if (!Directory.Exists(projectsFolder))
                    {
                        try
                        {
                            Directory.CreateDirectory(projectsFolder);
                            System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Created Projects folder: {projectsFolder}");
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException($"Cannot create Projects folder {projectsFolder}. Error: {ex.Message}", ex);
                        }
                    }
                }

                // Step 3: Create the final target directory (Directory.CreateDirectory creates parents automatically)
                if (!Directory.Exists(targetDirectory))
                {
                    try
                    {
                        Directory.CreateDirectory(targetDirectory);
                        System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Created target directory: {targetDirectory}");
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        throw new InvalidOperationException($"Access denied creating directory: {targetDirectory}. Please check folder permissions. User: {Environment.UserName}", ex);
                    }
                    catch (DirectoryNotFoundException ex)
                    {
                        throw new InvalidOperationException($"Parent directory not found for: {targetDirectory}", ex);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Failed to create directory: {targetDirectory}. Error: {ex.Message}", ex);
                    }
                }

                // Step 4: Verify the directory exists and is writable
                if (!Directory.Exists(targetDirectory))
                {
                    throw new InvalidOperationException($"Directory creation failed: {targetDirectory} does not exist after creation attempt. User: {Environment.UserName}");
                }

                // Test write access
                var testFile = Path.Combine(targetDirectory, $"write_test_{Guid.NewGuid():N}.tmp");
                try
                {
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);
                    System.Diagnostics.Debug.WriteLine($"ProfileManagementService: Directory write test passed: {targetDirectory}");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Directory exists but is not writable: {targetDirectory}. User may lack write permissions. Error: {ex.Message}", ex);
                }
            }
            catch (InvalidOperationException)
            {
                throw; // Re-throw InvalidOperationException as-is
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unexpected error creating directory structure for: {targetDirectory}. User: {Environment.UserName}. Error: {ex.Message}", ex);
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
            
            // DISABLED: Hardcoded log write - use DebugLogger instead
            // var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
            // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] CreateProfile: Creating profile '{profileName}' with {disciplines?.Count ?? 0} disciplines\n");
            
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
            // DISABLED: Hardcoded log write
            // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] CreateProfile: Added profile to collection, now have {_availableProfiles.Count} profiles\n");
            
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
                
                if (File.Exists(_profileFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"LoadProfiles: File size: {new FileInfo(_profileFilePath).Length} bytes");
                    
                    var json = File.ReadAllText(_profileFilePath);
                    var profiles = Newtonsoft.Json.JsonConvert.DeserializeObject<List<UserProfile>>(json);
                    
                    if (profiles != null)
                    {
                        _availableProfiles.Clear();
                        _availableProfiles.AddRange(profiles);
                        System.Diagnostics.Debug.WriteLine($"LoadProfiles: Successfully loaded {profiles.Count} profiles");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"LoadProfiles: Deserialized profiles is null");
                    }
                }
                else
                {
                    // Check for XML file (migration)
                    var xmlPath = _profileFilePath.Replace(".json", ".xml");
                    if (File.Exists(xmlPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"LoadProfiles: Found legacy XML file, attempting migration: {xmlPath}");
                        try
                        {
                            var serializer = new XmlSerializer(typeof(List<UserProfile>));
                            using (var reader = new FileStream(xmlPath, FileMode.Open))
                            {
                                var profiles = (List<UserProfile>?)serializer.Deserialize(reader);
                                if (profiles != null)
                                {
                                    _availableProfiles.Clear();
                                    _availableProfiles.AddRange(profiles);
                                    System.Diagnostics.Debug.WriteLine($"LoadProfiles: Successfully migrated {profiles.Count} profiles from XML");
                                    
                                    // Save as JSON immediately
                                    SaveProfiles();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"LoadProfiles: XML migration failed: {ex.Message}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"LoadProfiles: Profile file does not exist");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadProfiles: ERROR - {ex.Message}");
                
                // FALLBACK: Try to load profiles from config files when JSON/XML fails
                try
                {
                    LoadProfilesFromConfigFiles();
                }
                catch (Exception fallbackEx)
                {
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
                // DISABLED: Hardcoded log write - use DebugLogger instead
                // var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Starting fallback load\n");
                
                _availableProfiles.Clear();
                
                // Get the directory containing the profile files
                var profileDir = Path.GetDirectoryName(_profileFilePath);
                if (string.IsNullOrEmpty(profileDir) || !Directory.Exists(profileDir))
                {
                    // DISABLED: Hardcoded log write
                    // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Profile directory not found\n");
                    return;
                }
                
                // FIXED: Only look for config files in the CURRENT PROJECT directory
                // This ensures profiles are truly project-specific
                var configFiles = Directory.GetFiles(profileDir, "config_*.txt");
                // DISABLED: Hardcoded log write
                // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Looking in project-specific directory: {profileDir}\n");
                // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Found {configFiles.Length} config files\n");
                
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
                        // DISABLED: Hardcoded log write
                        // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Created profile '{profileName}' from config file\n");
                    }
                    catch (Exception) // FIX: CS0168 - 'ex' commented to fix critical warning
                    {
                        // DISABLED: Hardcoded log write
                        // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Error processing config file {configFile}: {ex.Message}\n");
                    }
                }

                // DISABLED: Hardcoded log write
                // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: Successfully loaded {_availableProfiles.Count} profiles from config files\n");
            }
            catch (Exception) // FIX: CS0168 - 'ex' commented to fix critical warning
            {
                // DISABLED: Hardcoded log write - use DebugLogger instead
                // var debugLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\profile_save_debug.log";
                // JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] LoadProfilesFromConfigFiles: ERROR - {ex.Message}\n");
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
                
                // Ensure directory exists
                try
                {
                    EnsureDirectoryStructure(_profileDirectory);
                }
                catch (Exception dirEx)
                {
                    System.Diagnostics.Debug.WriteLine($"SaveProfiles: WARNING - Directory creation failed: {dirEx.Message}. Attempting to save anyway.");
                }
                
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_save_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: Saving {_availableProfiles.Count} profiles to: {_profileFilePath}\n");
                
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(_availableProfiles, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(_profileFilePath, json);
                
                System.Diagnostics.Debug.WriteLine($"SaveProfiles: Successfully saved profiles to JSON file");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: Successfully saved profiles to JSON file\n");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SaveProfiles: ERROR - {ex.Message}");
                
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_save_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveProfiles: ERROR - {ex.Message}\n");
                
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
