using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Handles persistence of System Type Override settings to/from JSON file
    /// Saves to: %AppData%\JSE_RevitAddin_MEP_OPENINGS\MarkPrefixSettings.json
    /// </summary>
    public static class SystemTypeOverridePersistenceService
    {
        private static readonly string AppDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "JSE_RevitAddin_MEP_OPENINGS"
        );

        private static readonly string SettingsFilePath = Path.Combine(AppDataFolder, "MarkPrefixSettings.json");

        /// <summary>
        /// Load settings from JSON file, or return defaults if file doesn't exist
        /// </summary>
        public static MarkPrefixSettings LoadSettings()
        {
            RemarkDebugLogger.LogInfo("[UI-PERSIST] === LoadSettings CALLED ===");
            RemarkDebugLogger.LogInfo($"[UI-PERSIST] Settings file path: {SettingsFilePath}");
            
            try
            {
                if (!File.Exists(SettingsFilePath))
                {
                    RemarkDebugLogger.LogInfo("[UI-PERSIST] Settings file does NOT exist, returning defaults");
                    return new MarkPrefixSettings();
                }

                RemarkDebugLogger.LogInfo("[UI-PERSIST] Settings file EXISTS, reading...");
                string json = File.ReadAllText(SettingsFilePath);
                RemarkDebugLogger.LogInfo($"[UI-PERSIST] JSON content length: {json.Length} characters");
                
                var settings = JsonSerializer.Deserialize<MarkPrefixSettings>(json);

                if (settings != null)
                {
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST] Deserialization SUCCESS");
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkDuctPrefix: {settings.RemarkDuctPrefix}");
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkPipePrefix: {settings.RemarkPipePrefix}");
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkCableTrayPrefix: {settings.RemarkCableTrayPrefix}");
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkDamperPrefix: {settings.RemarkDamperPrefix}");
                    
                    return settings;
                }
                else
                {
                    RemarkDebugLogger.LogInfo("[UI-PERSIST] Deserialization returned NULL");
                }
            }
            catch (Exception ex)
            {
                RemarkDebugLogger.LogError($"[UI-PERSIST] Exception during load: {ex.Message}");
            }

            RemarkDebugLogger.LogInfo("[UI-PERSIST] Returning default settings due to error");
            return new MarkPrefixSettings();
        }

        /// <summary>
        /// Save settings to JSON file
        /// </summary>
        public static void SaveSettings(MarkPrefixSettings settings)
        {
            RemarkDebugLogger.LogInfo("[UI-PERSIST] === SaveSettings CALLED ===");
            RemarkDebugLogger.LogInfo($"[UI-PERSIST] Settings file path: {SettingsFilePath}");
            
            try
            {
                RemarkDebugLogger.LogInfo("[UI-PERSIST] Checking/creating directory...");
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                    RemarkDebugLogger.LogInfo($"[UI-PERSIST] Created directory: {AppDataFolder}");
                }

                RemarkDebugLogger.LogInfo("[UI-PERSIST] Serializing settings...");
                RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkDuctPrefix: {settings.RemarkDuctPrefix}");
                RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkPipePrefix: {settings.RemarkPipePrefix}");
                RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkCableTrayPrefix: {settings.RemarkCableTrayPrefix}");
                RemarkDebugLogger.LogInfo($"[UI-PERSIST]   RemarkDamperPrefix: {settings.RemarkDamperPrefix}");
                
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(settings, options);
                RemarkDebugLogger.LogInfo($"[UI-PERSIST] JSON length: {json.Length} characters");
                
                File.WriteAllText(SettingsFilePath, json);
                RemarkDebugLogger.LogInfo("[UI-PERSIST] File written successfully!");

                RemarkDebugLogger.LogInfo($"[MarkPrefixSettings] Saved settings: ProjectPrefix='{settings.ProjectPrefix}', NumberFormat='{settings.NumberFormat}', " +
                    $"Overrides: {settings.DuctSystemTypeOverrides?.Count ?? 0} duct, {settings.PipeSystemTypeOverrides?.Count ?? 0} pipe");
            }
            catch (Exception ex)
            {
                RemarkDebugLogger.LogError($"[UI-PERSIST] Exception during save: {ex.Message}");
                RemarkDebugLogger.LogError($"[UI-PERSIST] Stack trace: {ex.StackTrace}");
            }
        }
    }
}
