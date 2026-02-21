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
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    var settings = Newtonsoft.Json.JsonConvert.DeserializeObject<SettingsModel>(json);
                    
                    if (settings != null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"Settings loaded from file: {_settingsPath}");
                        return settings;
                    }
                }
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
