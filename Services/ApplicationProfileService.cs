using System;
using System.IO;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs;

using JSE_RevitAddin_MEP_OPENINGS.Services;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Application-level profile service that manages profile state across the application
    /// </summary>
    public class ApplicationProfileService : IDisposable
    {
        private static ApplicationProfileService? _instance;
        private static readonly object _lock = new object();

        private ProfileManagementService _profileService;
        private readonly StatusManager _statusManager;
        private SettingsService _settingsService;
        private UserProfile? _currentProfile;
        
        // Caching for CheckAndUpdateContext
        private static Type? _contextType;
        private static System.Reflection.PropertyInfo? _activeDocProp;
        private static bool _reflectionInitialized = false;
        private DateTime _lastContextCheck = DateTime.MinValue;
        private int _lastDocId = -1;
        private string? _lastDocTitle = null;
        private const int CONTEXT_CHECK_THROTTLE_MS = 500;
        
        public event EventHandler<ProfileChangedEventArgs>? ProfileChanged;
        public event EventHandler<StatusUpdateEventArgs>? StatusUpdated;


        private ApplicationProfileService()
        {
            // Initialize with a default profile service - will be updated when Revit document is available
            // ✅ CRITICAL FIX: Use parameterless constructor to trigger safe AppData fallback 
            // Passing "Default" caused relative path crash (C:\Program Files\Autodesk\Revit 2025\Default)
            _profileService = new ProfileManagementService(); 
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
            // DYNAMIC CONTEXT SWITCH: Ensure we are using settings for the active document
            CheckAndUpdateContext();

            // Always load settings from XML file, even if profile is null
            if (CurrentProfile == null)
            {
                // Load from XML file directly
                return _settingsService.LoadSettings();
            }

            return _settingsService.GetSettings(CurrentProfile);
        }

        public void SaveCurrentSettings(SettingsModel settings)
        {
            // DYNAMIC CONTEXT SWITCH: Ensure we are saving settings for the active document
            CheckAndUpdateContext();

            // Always save settings to XML file, even if profile is null
            if (CurrentProfile != null)
            {
                _settingsService.SaveSettings(CurrentProfile, settings);
            }
            else
            {
                // Save directly to XML file when no profile is set
                _settingsService.SaveSettings(settings);
            }
        }

        /// <summary>
        /// Checks the active document and updates the settings service path if it has changed.
        /// This ensures settings are saved in the correct project folder.
        /// Throttled and optimized to avoid UI thread blocks.
        /// </summary>
        private void CheckAndUpdateContext()
        {
            try
            {
                // 1. Throttle checks to avoid overhead on every property access
                TimeSpan elapsed = DateTime.Now - _lastContextCheck;
                if (elapsed.TotalMilliseconds < CONTEXT_CHECK_THROTTLE_MS)
                    return;
                
                _lastContextCheck = DateTime.Now;

                // 2. Initialize reflection once
                if (!_reflectionInitialized)
                {
                    InitializeReflection();
                }

                if (_contextType != null && _activeDocProp != null)
                {
                    var doc = _activeDocProp.GetValue(null) as Document;
                    if (doc != null && !doc.IsFamilyDocument)
                    {
                        // 3. Only proceed if document has actually changed
                        string currentTitle = doc.Title;
                        int currentId = doc.GetHashCode(); // Simple ID check

                        if (currentId == _lastDocId && currentTitle == _lastDocTitle)
                            return;

                        _lastDocId = currentId;
                        _lastDocTitle = currentTitle;

                        try 
                        {
                            // ✅ CRITICAL FIX: Use ProjectPathService to get standardized root path
                            string projectRoot = ProjectPathService.GetProjectRoot(doc);
                            
                            if (!string.IsNullOrEmpty(projectRoot)) 
                            {
                                UpdateForCurrentDocument(projectRoot);
                            }
                        } 
                        catch (Exception ex2)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error resolving project root: {ex2.Message}");
                        } 
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CheckAndUpdateContext: {ex.Message}");
            }
        }

        private void InitializeReflection()
        {
            try
            {
                // Try direct load/type retrieval first (fastest)
                _contextType = Type.GetType("Nice3point.Revit.Toolkit.Context, Nice3point.Revit.Toolkit");
                
                if (_contextType == null)
                {
                    // Only iterate if necessary, and use a faster check
                    var assemblies = AppDomain.CurrentDomain.GetAssemblies();
                    for (int i = 0; i < assemblies.Length; i++)
                    {
                        var asm = assemblies[i];
                        if (asm.FullName.StartsWith("Nice3point.Revit.Toolkit", StringComparison.OrdinalIgnoreCase))
                        {
                            _contextType = asm.GetType("Nice3point.Revit.Toolkit.Context");
                            if (_contextType != null) break;
                        }
                    }
                }

                if (_contextType != null)
                {
                    _activeDocProp = _contextType.GetProperty("ActiveDocument", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                }
                
                _reflectionInitialized = true;
            }
            catch { _reflectionInitialized = true; }
        }



        private string? _lastDocumentPath = null;

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
        /// Cleans up the singleton instance
        /// </summary>
        public static void CleanupInstance()
        {
            lock (_lock)
            {
                if (_instance != null)
                {
                    if (_instance._profileService != null)
                    {
                        _instance._profileService.ProfileChanged -= _instance.OnProfileServiceProfileChanged;
                        _instance._profileService.StatusUpdated -= _instance.OnProfileServiceStatusUpdated;
                    }
                    
                    if (_instance._statusManager != null)
                    {
                        _instance._statusManager.StatusUpdated -= _instance.OnStatusManagerStatusUpdated;
                    }
                    
                    _instance._currentProfile = null;
                    _instance = null;
                }
            }
        }

        /// <summary>
        /// Resets the singleton instance for a new project
        /// </summary>
        public static void ResetForNewProject()
        {
            lock (_lock)
            {
                CleanupInstance();
                _instance = new ApplicationProfileService();
                
                System.Diagnostics.Debug.WriteLine("ApplicationProfileService: Reset for new project - fresh instance created");
                
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] ResetForNewProject: Fresh instance created\n");
            }
        }

        public UserProfile? CurrentProfile 
        {
            get
            {
                CheckAndUpdateContext();
                return _currentProfile;
            }
        }

        public ProfileManagementService ProfileService 
        {
            get
            {
                CheckAndUpdateContext();
                return _profileService;
            }
        }
        public StatusManager StatusManager => _statusManager;

        public bool IsProfileSetupRequired 
        { 
            get 
            {
                try
                {
                    // Ensure context is correct before checking
                    CheckAndUpdateContext();
                    
                    var debugLogPath = SafeFileLogger.GetLogFilePath("profile_check_debug.log");
                    bool hasProfiles = _profileService.AvailableProfiles.Any();
                    return !hasProfiles;
                }
                catch (Exception ex)
                {
                    if (_profileService == null)
                    {
                        return true; 
                    }
                    return !_profileService.AvailableProfiles.Any() || _currentProfile == null;
                }
            } 
        }

        public string CurrentUserInitials
        {
            get
            {
                CheckAndUpdateContext();
                if (_currentProfile == null)
                    return "UNK";

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

        public string CurrentUserDiscipline
        {
            get
            {
                CheckAndUpdateContext();
                if (_currentProfile?.PrimaryDiscipline != null)
                    return _currentProfile.PrimaryDiscipline.Name;
                
                return "Unknown";
            }
        }

        public void SetCurrentProfile(UserProfile profile)
        {
            CheckAndUpdateContext();
            if (profile == null)
                throw new ArgumentNullException(nameof(profile));

            _currentProfile = profile;
            _profileService.SetCurrentProfile(profile);
        }

        public UserProfile CreateProfile(string profileName, System.Collections.Generic.List<Discipline> disciplines, string language)
        {
            CheckAndUpdateContext();
            var profile = _profileService.CreateProfile(profileName, disciplines, language);
            SetCurrentProfile(profile);
            return profile;
        }

        private void LoadCurrentProfile()
        {
            try
            {
                _currentProfile = null;
                System.Diagnostics.Debug.WriteLine("LoadCurrentProfile: Cleared current profile - will be set project-specific later");
            }
            catch (Exception ex)
            {
                _statusManager.UpdateStatus($"Failed to load current profile: {ex.Message}", StatusType.Error);
            }
        }

        public bool HasDisciplineAccess(string disciplineName)
        {
            CheckAndUpdateContext();
            if (_currentProfile == null)
                return false;

            return _currentProfile.Disciplines.Any(d => 
                d.IsSelected && d.Name.Equals(disciplineName, StringComparison.OrdinalIgnoreCase));
        }

        public bool IsPrimaryDisciplineUser()
        {
            CheckAndUpdateContext();
            return _currentProfile?.PrimaryDiscipline != null;
        }

        public string GetUserRoleDescription()
        {
            CheckAndUpdateContext();
            if (_currentProfile == null)
                return "No Profile Set";

            var primary = _currentProfile.PrimaryDiscipline?.Name ?? "No Primary";
            var mepDisciplines = string.Join(", ", _currentProfile.MepDisciplines.Select(d => d.Name));
            
            if (string.IsNullOrEmpty(mepDisciplines))
                return primary;
            
            return $"{primary} + {mepDisciplines}";
        }

        public string GetUserInfoString()
        {
            CheckAndUpdateContext();
            if (_currentProfile == null)
                return "No Profile Active";

            return $"{_currentProfile.Name} ({CurrentUserInitials}) - {GetUserRoleDescription()}";
        }

        /// <summary>
        /// Updates the profile and settings context for the given Revit document.
        /// Uses ProjectPathService to resolve the standardized AppData project path.
        /// Prefer this overload over the string-based one to ensure correct path resolution.
        /// </summary>
        public void UpdateForCurrentDocument(Document doc)
        {
            if (doc == null || doc.IsFamilyDocument) return;
            string projectRoot = ProjectPathService.GetProjectRoot(doc);
            UpdateForCurrentDocument(projectRoot);
        }

        public void UpdateForCurrentDocument(string? documentPath)
        {
            if (string.Equals(_lastDocumentPath, documentPath, StringComparison.OrdinalIgnoreCase) && _profileService != null)
            {
                return;
            }

            _lastDocumentPath = documentPath;

            if (!string.IsNullOrEmpty(documentPath))
            {
                System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Creating new ProfileManagementService for path: {documentPath}");
                
                var debugLogPath = SafeFileLogger.GetLogFilePath("update_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Creating new ProfileManagementService for path: {documentPath}\n");
                
                var newProfileService = new ProfileManagementService(documentPath);
                
                var currentProjectPath = _profileService.ProfileFilePath;
                var newProjectPath = newProfileService.ProfileFilePath;
                
                if (currentProjectPath != newProjectPath)
                {
                    _currentProfile = null;
                    JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Different project - cleared current profile\n");
                }
                
                try
                {
                    var projectProfileDir = Path.GetDirectoryName(newProfileService.ProfileFilePath);
                    var currentProfileFile = Path.Combine(projectProfileDir ?? "", "current_profile.txt");
                    
                    if (File.Exists(currentProfileFile))
                    {
                        var savedProfileName = File.ReadAllText(currentProfileFile).Trim();
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] UpdateForCurrentDocument: Found saved profile name: {savedProfileName}\n");
                        
                        var savedProfile = newProfileService.AvailableProfiles.FirstOrDefault(p => p.Name == savedProfileName);
                        if (savedProfile != null)
                        {
                            _currentProfile = savedProfile;
                        }
                        else
                        {
                            var configFile = Path.Combine(projectProfileDir ?? "", $"config_{savedProfileName}.txt");
                            if (File.Exists(configFile))
                            {
                                var recreatedProfile = new UserProfile
                                {
                                    Id = Guid.NewGuid(),
                                    Name = savedProfileName,
                                    Disciplines = new List<Discipline> { new Discipline("Mechanical", true, "Mechanical systems") },
                                    Language = "English",
                                    CreatedDate = DateTime.Now,
                                    IsActive = true
                                };
                                
                                newProfileService.AddProfile(recreatedProfile);
                                _currentProfile = recreatedProfile;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: Error loading saved profile: {ex.Message}");
                }
                
                _profileService.ProfileChanged -= OnProfileServiceProfileChanged;
                _profileService.StatusUpdated -= OnProfileServiceStatusUpdated;
                
                _profileService = newProfileService;
                
                _profileService.ProfileChanged += OnProfileServiceProfileChanged;
                _profileService.StatusUpdated += OnProfileServiceStatusUpdated;
                
                _settingsService = new SettingsService(documentPath);
                System.Diagnostics.Debug.WriteLine($"UpdateForCurrentDocument: SettingsService updated for path: {documentPath}");
            }
        }

        private string? GetCurrentDocumentPath()
        {
            try
            {
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

        public static void ForceProfileSetupForNewFile()
        {
            lock (_lock)
            {
                CleanupInstance();
                _instance = new ApplicationProfileService();
                _instance._currentProfile = null;
                
                System.Diagnostics.Debug.WriteLine("ApplicationProfileService: Forced profile setup for new file");
            }
        }

        public UserProfile? GetCurrentProfile()
        {
            return _currentProfile;
        }
        
        public void SaveCurrentProfile()
        {
            if (_currentProfile != null)
            {
                try
                {
                    var debugLogPath = SafeFileLogger.GetLogFilePath("profile_save_debug.log");
                    
                    var projectProfileDir = Path.GetDirectoryName(_profileService.ProfileFilePath);
                    if (!string.IsNullOrEmpty(projectProfileDir) && !Directory.Exists(projectProfileDir))
                        Directory.CreateDirectory(projectProfileDir);
                    
                    var currentProfileFile = Path.Combine(projectProfileDir ?? "", "current_profile.txt");
                    File.WriteAllText(currentProfileFile, _currentProfile.Name);
                    
                    try
                    {
                        SaveConfigurationOnly(_currentProfile);
                    }
                    catch (Exception ex)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Failed to save configuration: {ex.Message}\n");
                    }
                    
                    try
                    {
                        _profileService.UpdateProfile(_currentProfile, _currentProfile.Name, _currentProfile.Disciplines, _currentProfile.Language);
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
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_save_debug.log");
                
                if (profile.Configuration == null)
                    return;
                
                var projectProfileDir = Path.GetDirectoryName(_profileService.ProfileFilePath);
                
                if (string.IsNullOrEmpty(projectProfileDir))
                {
                    projectProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings");
                }
                
                if (!Directory.Exists(projectProfileDir))
                {
                    Directory.CreateDirectory(projectProfileDir);
                }
                
                var configFile = Path.Combine(projectProfileDir, $"config_{profile.Name}.txt");
                
                using (var writer = new StreamWriter(configFile, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine($"ProfileName={profile.Name}");
                    writer.WriteLine($"LastModified={DateTime.Now:O}");
                    
                    writer.WriteLine("SelectedReferenceFiles=");
                    foreach (var file in profile.Configuration.SelectedReferenceFiles ?? new List<string>())
                    {
                        writer.WriteLine($"  {file}");
                    }
                    
                    writer.WriteLine("SelectedHostFiles=");
                    foreach (var file in profile.Configuration.SelectedHostFiles ?? new List<string>())
                    {
                        writer.WriteLine($"  {file}");
                    }
                    
                    writer.WriteLine("SelectedMepCategories=");
                    foreach (var cat in profile.Configuration.SelectedMepCategories ?? new List<string>())
                    {
                        writer.WriteLine($"  {cat}");
                    }
                    
                    writer.WriteLine("SelectedHostCategories=");
                    foreach (var cat in profile.Configuration.SelectedHostCategories ?? new List<string>())
                    {
                        writer.WriteLine($"  {cat}");
                    }
                    
                    writer.WriteLine("ClearanceSettings=");
                    if (profile.Configuration.OpeningSettings?.ClearanceSettings != null)
                    {
                        foreach (var kvp in profile.Configuration.OpeningSettings.ClearanceSettings)
                        {
                            writer.WriteLine($"  {kvp.Key}={kvp.Value}");
                        }
                    }
                    
                    writer.WriteLine("ClashZoneStorage=");
                    if (profile.Configuration.ClashZoneStorage?.AllZones != null)
                    {
                        writer.WriteLine($"  LastUpdated={profile.Configuration.ClashZoneStorage.LastUpdated:O}");
                        writer.WriteLine($"  DocumentHash={profile.Configuration.ClashZoneStorage.DocumentHash}");
                        writer.WriteLine($"  ClashZonesCount={profile.Configuration.ClashZoneStorage.AllZones.Count}");
                        foreach (var clashZone in profile.Configuration.ClashZoneStorage.AllZones)
                        {
                            writer.WriteLine($"  ClashZone: MEP={clashZone.MepElementId}, Structural={clashZone.StructuralElementId}, Resolved={clashZone.IsResolved}");
                        }
                    }
                }
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] Configuration saved to: {configFile}\n");
            }
            catch (Exception ex)
            {
                var debugLogPath = SafeFileLogger.GetLogFilePath("profile_save_debug.log");
                JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText(debugLogPath, $"[{DateTime.Now}] SaveConfigurationOnly failed: {ex.Message}\n");
                throw;
            }
        }
    }
}
