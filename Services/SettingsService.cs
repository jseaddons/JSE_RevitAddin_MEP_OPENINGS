using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class SettingsService
    {
        private readonly string _settingsPath;
        
        public SettingsService()
        {
            // Settings file location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Settings.xml
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appDataPath, "JSE_MEP_Openings");
            _settingsPath = Path.Combine(configPath, "Settings.xml");
            
            // Ensure directory exists
            Directory.CreateDirectory(configPath);
        }

        public SettingsModel GetSettings(UserProfile profile)
        {
            try
            {
                // Try to load from file first
                if (File.Exists(_settingsPath))
                {
                    var serializer = new XmlSerializer(typeof(SettingsModel));
                    using (var reader = new StreamReader(_settingsPath))
                    {
                        var settings = (SettingsModel)serializer.Deserialize(reader)!;
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"Settings loaded from file: {_settingsPath}");
                        return settings;
                    }
                }
                else
                {
                    // Return default settings if file doesn't exist
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log("Settings file not found, returning default settings");
                    return new SettingsModel();
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to load settings: {ex.Message}");
                return new SettingsModel();
            }
        }

        public void SaveSettings(UserProfile profile, SettingsModel settings)
        {
            try
            {
                // Save to file
                var serializer = new XmlSerializer(typeof(SettingsModel));
                using (var writer = new StreamWriter(_settingsPath))
                {
                    serializer.Serialize(writer, settings);

                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Settings saved to file: {_settingsPath}");
                
                // Also save to profile if available
                if (profile.Configuration != null)
                {
                    profile.Configuration.AdvancedSettings = settings;
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to save settings: {ex.Message}");
                throw;
            }
        }
        
        public SettingsModel LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var serializer = new XmlSerializer(typeof(SettingsModel));
                    using (var reader = new StreamReader(_settingsPath))
                    {
                        var settings = (SettingsModel)serializer.Deserialize(reader)!;
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Log($"Settings loaded from file: {_settingsPath}");
                        return settings;
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log("Settings file not found, returning default settings");
                    return new SettingsModel();
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Log($"Failed to load settings: {ex.Message}");
                return new SettingsModel();
            }
        }
        
        public void SaveSettings(SettingsModel settings)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(SettingsModel));
                using (var writer = new StreamWriter(_settingsPath))
                {
                    serializer.Serialize(writer, settings);

                }
                
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
