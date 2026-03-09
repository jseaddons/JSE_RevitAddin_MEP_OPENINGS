using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class SettingsService
    {
        private readonly string _settingsPath;
        
        // projectRootDirectory: expected to be "...\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]"
        public SettingsService(string? projectRootDirectory = null)
        {
            if (!string.IsNullOrEmpty(projectRootDirectory))
            {
                try
                {
                    // projectRootDirectory is ...\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]
                    // Save Settings.json DIRECTLY in this folder.
                    
                    if (!Directory.Exists(projectRootDirectory))
                    {
                        Directory.CreateDirectory(projectRootDirectory);
                    }
                    
                    _settingsPath = Path.Combine(projectRootDirectory, "Settings.json");
                    System.Diagnostics.Debug.WriteLine($"SettingsService: Using settings path: {_settingsPath}");
                    
                    // FORCE CREATION: If settings file doesn't exist, create it with defaults immediately
                    // This ensures the user sees the file in the folder right away.
                    if (!File.Exists(_settingsPath))
                    {
                        try 
                        {
                            SaveSettings(new SettingsModel());
                            System.Diagnostics.Debug.WriteLine($"SettingsService: Created default settings file at {_settingsPath}");
                        }
                        catch (Exception saveEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"SettingsService: Failed to create default settings: {saveEx.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SettingsService: Error with provided path: {ex.Message}. Using default fallback.");
                     _settingsPath = GetGlobalFallbackPath();
                }
            }
            else
            {
                // Global/Fallback settings
               _settingsPath = GetGlobalFallbackPath();
            }
        }
        
        private string GetGlobalFallbackPath()
        {
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appDataPath, "JSE_MEP_Openings");
            Directory.CreateDirectory(configPath);
            return Path.Combine(configPath, "Settings.json");
        }

        public SettingsModel GetSettings(UserProfile profile)
        {
            // If profile has specific settings, return them (though currently we don't fully sync this)
            // Ideally, we load from the file which is the source of truth for this project/context
            return LoadSettings();
        }

        public void SaveSettings(UserProfile profile, SettingsModel settings)
        {
            SaveSettings(settings);
            
            // Also update the profile's in-memory copy
            if (profile != null && profile.Configuration != null)
            {
                profile.Configuration.AdvancedSettings = settings;
            }
        }
        
        public SettingsModel LoadSettings()
        {
            SettingsModel projectSettings = null;
            bool isProjectSpecificPath = !string.Equals(_settingsPath, GetGlobalFallbackPath(), StringComparison.OrdinalIgnoreCase);

            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    projectSettings = Newtonsoft.Json.JsonConvert.DeserializeObject<SettingsModel>(json);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to load settings (JSON): {ex.Message}");
            }

            // ✅ SETTINGS MERGE FIX: If this is a project-specific settings file and key threshold
            // settings are at their defaults (0), merge in values from the global fallback settings.
            // This prevents the per-project Settings.json (auto-created with defaults) from silently
            // overriding values the user set via the UI (which may have saved to the global path).
            if (isProjectSpecificPath)
            {
                var globalPath = GetGlobalFallbackPath();
                if (File.Exists(globalPath))
                {
                    try
                    {
                        var globalJson = File.ReadAllText(globalPath);
                        var globalSettings = Newtonsoft.Json.JsonConvert.DeserializeObject<SettingsModel>(globalJson);

                        if (globalSettings != null)
                        {
                            if (projectSettings == null)
                            {
                                // No project-specific file at all → use global directly
                                projectSettings = globalSettings;
                            }
                            else
                            {
                                // Merge: copy global values for fields that are still at default (0) in project settings
                                if (projectSettings.MinWallThickness <= 0 && globalSettings.MinWallThickness > 0)
                                    projectSettings.MinWallThickness = globalSettings.MinWallThickness;
                                if (projectSettings.IgnoreArchitecturalFloors == false && globalSettings.IgnoreArchitecturalFloors)
                                    projectSettings.IgnoreArchitecturalFloors = globalSettings.IgnoreArchitecturalFloors;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"[SETTINGS-MERGE] Failed to load global settings for merge: {ex.Message}");
                    }
                }
            }

            if (projectSettings != null)
            {
                // ✅ ALWAYS log which path was used and key threshold values for diagnosability
                SafeFileLogger.SafeAppendText("settings_load_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] Settings loaded from: {_settingsPath} | MinWallThickness={projectSettings.MinWallThickness}mm | IgnoreArchFloors={projectSettings.IgnoreArchitecturalFloors}\n");
                return projectSettings;
            }

            // --- legacy JSON load failed, keep original fallback flow ---
            try
            {
                // (no-op — errors already handled above)
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to load settings (JSON): {ex.Message}");
            }

            // Fallback: Check for old XML file if JSON doesn't exist
            if (_settingsPath.EndsWith(".json"))
            {
                var xmlPath = _settingsPath.Replace(".json", ".xml");
                if (File.Exists(xmlPath))
                {
                    try
                    {
                        var serializer = new XmlSerializer(typeof(SettingsModel));
                        using (var reader = new StreamReader(xmlPath))
                        {
                            var settings = (SettingsModel)serializer.Deserialize(reader)!;
                            // Save as JSON for next time
                            SaveSettings(settings); 
                            return settings;
                        }
                    }
                    catch { /* Ignore XML load failure */ }
                }
            }

            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log("Settings file not found or invalid, returning default settings");
                
            return new SettingsModel();
        }
        
        public void SaveSettings(SettingsModel settings)
        {
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(settings, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(_settingsPath, json);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Settings saved to file: {_settingsPath}");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to save settings: {ex.Message}");
                throw;
            }
        }
    }
}
