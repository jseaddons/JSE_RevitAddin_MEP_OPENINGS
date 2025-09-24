using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Configuration
{
    /// <summary>
    /// Singleton service for managing global application configuration
    /// Handles loading, saving, and accessing global settings
    /// </summary>
    public class GlobalConfigurationManager
    {
        private static GlobalConfigurationManager _instance;
        private static readonly object _lock = new object();
        
        private GlobalApplicationSettings _settings;
        private readonly string _settingsPath;
        
        private GlobalConfigurationManager()
        {
            // Global Settings Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\GlobalSettings.xml
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appDataPath, "JSE_MEP_Openings");
            _settingsPath = Path.Combine(configPath, "GlobalSettings.xml");
            
            LoadSettings();
        }
        
        public static GlobalConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new GlobalConfigurationManager();
                        }
                    }
                }
                return _instance;
            }
        }
        
        public GlobalApplicationSettings Settings => _settings;
        
        /// <summary>
        /// Load global settings from XML file
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var serializer = new XmlSerializer(typeof(GlobalApplicationSettings));
                    using (var reader = new FileStream(_settingsPath, FileMode.Open))
                    {
                        _settings = (GlobalApplicationSettings)serializer.Deserialize(reader);
                        DebugLogger.Info($"[GlobalConfig] Loaded settings from: {_settingsPath}");
                    }
                }
                else
                {
                    // Create default settings if file doesn't exist
                    _settings = new GlobalApplicationSettings();
                    DebugLogger.Info("[GlobalConfig] Created default settings (file not found)");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GlobalConfig] Error loading settings: {ex.Message}");
                _settings = new GlobalApplicationSettings(); // Fallback to defaults
            }
        }
        
        /// <summary>
        /// Save global settings to XML file
        /// </summary>
        public void SaveSettings()
        {
            try
            {
                // Ensure directory exists
                var configDir = Path.GetDirectoryName(_settingsPath);
                if (!Directory.Exists(configDir))
                {
                    Directory.CreateDirectory(configDir);
                    DebugLogger.Info($"[GlobalConfig] Created directory: {configDir}");
                }
                
                var serializer = new XmlSerializer(typeof(GlobalApplicationSettings));
                using (var writer = new FileStream(_settingsPath, FileMode.Create))
                {
                    serializer.Serialize(writer, _settings);
                    DebugLogger.Info($"[GlobalConfig] Saved settings to: {_settingsPath}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GlobalConfig] Error saving settings: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Update a specific configuration setting
        /// </summary>
        public void UpdateConfigurationSetting(string settingName, object value)
        {
            try
            {
                var property = typeof(ConfigurationDefaults).GetProperty(settingName);
                if (property != null && property.CanWrite)
                {
                    property.SetValue(_settings.ConfigurationDefaults, value);
                    DebugLogger.Info($"[GlobalConfig] Updated {settingName} = {value}");
                }
                else
                {
                    DebugLogger.Warning($"[GlobalConfig] Property {settingName} not found or not writable");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GlobalConfig] Error updating {settingName}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get a specific configuration setting value
        /// </summary>
        public T GetConfigurationSetting<T>(string settingName)
        {
            try
            {
                var property = typeof(ConfigurationDefaults).GetProperty(settingName);
                if (property != null)
                {
                    var value = property.GetValue(_settings.ConfigurationDefaults);
                    if (value is T)
                    {
                        return (T)value;
                    }
                }
                DebugLogger.Warning($"[GlobalConfig] Property {settingName} not found or wrong type");
                return default(T);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GlobalConfig] Error getting {settingName}: {ex.Message}");
                return default(T);
            }
        }
        
        /// <summary>
        /// Reset all configuration settings to defaults
        /// </summary>
        public void ResetToDefaults()
        {
            _settings.ConfigurationDefaults = new ConfigurationDefaults();
            DebugLogger.Info("[GlobalConfig] Reset all configuration settings to defaults");
        }
        
        /// <summary>
        /// Get the full path to the settings file
        /// </summary>
        public string GetSettingsFilePath()
        {
            return _settingsPath;
        }
    }
}








