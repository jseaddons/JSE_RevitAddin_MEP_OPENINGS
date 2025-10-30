using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ⚠️ CRITICAL SERVICE - DO NOT REMOVE ⚠️
    /// Manages saving and loading of OpeningConditions (clearances, opening types, etc.)
    /// Saved as FilterName_CONDITIONS.xml alongside clash zone data
    /// This implements proper separation: CLASH ZONES (data) vs CONDITIONS (configuration)
    /// </summary>
    public class ConditionsService
    {
        private readonly string _projectDirectory;
        private readonly Action<string> _log;

        public ConditionsService(Action<string> log = null)
        {
            _log = log ?? (msg => { });
            
            // Use same directory structure as clash zones
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _projectDirectory = Path.Combine(appData, "JSE_MEP_Openings", "Projects", "Default", "Filters");
            
            // Ensure directory exists
            if (!Directory.Exists(_projectDirectory))
            {
                Directory.CreateDirectory(_projectDirectory);
            }
        }

        /// <summary>
        /// Overload: allow callers to specify the exact project filters directory
        /// </summary>
        public ConditionsService(string projectFiltersDirectory, Action<string> log = null)
        {
            _log = log ?? (msg => { });
            _projectDirectory = string.IsNullOrWhiteSpace(projectFiltersDirectory)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters")
                : projectFiltersDirectory;
            if (!Directory.Exists(_projectDirectory))
            {
                Directory.CreateDirectory(_projectDirectory);
            }
        }

        /// <summary>
        /// Save opening conditions to XML file
        /// File name: FilterName_CONDITIONS.xml (e.g., Ventilation_ducts_CONDITIONS.xml)
        /// </summary>
        public bool SaveConditions(OpeningConditions conditions, string customKey = null)
        {
            try
            {
                // Use custom key if provided, otherwise use FilterName
                string key = customKey ?? conditions.FilterName;
                
                if (string.IsNullOrEmpty(key))
                {
                    _log("Cannot save conditions: Key is empty");
                    return false;
                }

                // Update timestamps
                conditions.LastModified = DateTime.Now;

                // Build file path
                var fileName = $"{key}_CONDITIONS.xml";
                var filePath = Path.Combine(_projectDirectory, fileName);

                // Serialize to XML
                var serializer = new XmlSerializer(typeof(OpeningConditions));
                using (var writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, conditions);
                }

                _log($"[ConditionsService] Saved conditions to: {filePath}");
                DebugLogger.Info($"[ConditionsService] Saved conditions for key '{key}' - Clearances: Rect={conditions.ClearanceSettings.RectangularNormal}/{conditions.ClearanceSettings.RectangularInsulated}mm, Round={conditions.ClearanceSettings.RoundNormal}/{conditions.ClearanceSettings.RoundInsulated}mm");
                
                return true;
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] Error saving conditions: {ex.Message}");
                DebugLogger.Error($"[ConditionsService] Error saving conditions: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Load opening conditions from XML file
        /// File name: FilterName_CONDITIONS.xml
        /// Returns null if file doesn't exist
        /// </summary>
        public OpeningConditions LoadConditions(string filterName)
        {
            try
            {
                if (string.IsNullOrEmpty(filterName))
                {
                    _log("[ConditionsService] Cannot load conditions: FilterName is empty");
                    return CreateDefaultConditions(filterName);
                }

                // Build file path
                var fileName = $"{filterName}_CONDITIONS.xml";
                var filePath = Path.Combine(_projectDirectory, fileName);

                // Check if file exists
                if (!File.Exists(filePath))
                {
                    _log($"[ConditionsService] Conditions file not found: {filePath}, creating defaults");
                    return CreateDefaultConditions(filterName);
                }

                // Deserialize from XML
                var serializer = new XmlSerializer(typeof(OpeningConditions));
                using (var reader = new StreamReader(filePath))
                {
                    var conditions = (OpeningConditions)serializer.Deserialize(reader);
                    _log($"[ConditionsService] Loaded conditions from: {filePath}");
                    DebugLogger.Info($"[ConditionsService] Loaded conditions for filter '{filterName}' - Clearances: Rect={conditions.ClearanceSettings.RectangularNormal}/{conditions.ClearanceSettings.RectangularInsulated}mm, Round={conditions.ClearanceSettings.RoundNormal}/{conditions.ClearanceSettings.RoundInsulated}mm");
                    return conditions;
                }
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] Error loading conditions: {ex.Message}");
                DebugLogger.Error($"[ConditionsService] Error loading conditions: {ex.Message}");
                return CreateDefaultConditions(filterName);
            }
        }

        /// <summary>
        /// Create default conditions for a filter
        /// </summary>
        private OpeningConditions CreateDefaultConditions(string filterName)
        {
            return new OpeningConditions
            {
                FilterName = filterName,
                Category = "Ducts", // Will be updated based on actual filter
                ClearanceSettings = new ClearanceSettings
                {
                    RectangularNormal = 50.0,
                    RectangularInsulated = 25.0,
                    RoundNormal = 50.0,
                    RoundInsulated = 50.0
                },
                OpeningTypePreferences = new OpeningTypePreferences
                {
                    RoundDucts = "Circular",
                    Pipes = "Circular"
                }
            };
        }

        /// <summary>
        /// Get the file path for a filter's conditions XML
        /// </summary>
        public string GetConditionsFilePath(string filterName)
        {
            return Path.Combine(_projectDirectory, $"{filterName}_CONDITIONS.xml");
        }

        /// <summary>
        /// Check if conditions file exists for a filter
        /// </summary>
        public bool ConditionsExist(string filterName)
        {
            var filePath = GetConditionsFilePath(filterName);
            return File.Exists(filePath);
        }
    }
}


