using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Represents a catalog of unique system and service types discovered from persisted clash data.
    /// </summary>
    public class SystemTypeCatalog
    {
        public HashSet<string> SystemTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> ServiceTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public bool IsEmpty => SystemTypes.Count == 0 && ServiceTypes.Count == 0;
    }

    /// <summary>
    /// Loads unique System Type / Service Type values from stored filter XML (and future persistence sources).
    /// Used to drive UI dropdowns for System Type overrides.
    /// </summary>
    public class SystemTypeCatalogService
    {
        private readonly Action<string> _logger;

        public SystemTypeCatalogService(Action<string> logger = null)
        {
            _logger = logger ?? (_ => { });
        }

        public SystemTypeCatalog Load(Document document)
        {
            var catalog = new SystemTypeCatalog();

            if (document == null)
            {
                _logger("[SystemTypeCatalog] No document context supplied – returning empty catalog.");
                return catalog;
            }

            try
            {
                // ✅ STEP 1: Load from persisted XML files (clash zones)
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(document);
                if (Directory.Exists(filtersDirectory))
                {
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                    var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                    foreach (var xmlFile in xmlFiles)
                    {
                        try
                        {
                            using (var reader = new StreamReader(xmlFile))
                            {
                                if (!(serializer.Deserialize(reader) is OpeningFilter filter) ||
                                    filter?.ClashZoneStorage?.AllZones == null ||
                                    filter.ClashZoneStorage.AllZones.Count == 0)
                                {
                                    continue;
                                }

                                foreach (var zone in filter.ClashZoneStorage.AllZones)
                                {
                                    // Extract System Type (for Ducts, Pipes)
                                    var systemType = ExtractParameterValue(zone, "System Type", "System Classification");
                                    if (!string.IsNullOrWhiteSpace(systemType))
                                    {
                                        catalog.SystemTypes.Add(systemType.Trim());
                                    }

                                    // ✅ Extract Service Type (for Cable Trays, Conduits) - separate from System Type
                                    var serviceType = ExtractParameterValue(zone, "Service Type");
                                    if (!string.IsNullOrWhiteSpace(serviceType))
                                    {
                                        catalog.ServiceTypes.Add(serviceType.Trim());
                                    }

                                    // Also extract System Abbreviation (can be used as service type fallback)
                                    var systemAbbr = ExtractParameterValue(zone, "System Abbreviation");
                                    if (!string.IsNullOrWhiteSpace(systemAbbr))
                                    {
                                        catalog.ServiceTypes.Add(systemAbbr.Trim());
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger($"[SystemTypeCatalog] Failed to parse '{Path.GetFileName(xmlFile)}': {ex.Message}");
                        }
                    }
                }

                // ✅ STEP 2: Load from SQLite database (MepParametersJson in SleeveSnapshots)
                // ✅ PERFORMANCE: No Revit API calls - use database that was populated during refresh
                LoadSystemTypesFromDatabase(document, catalog);
            }
            catch (Exception ex)
            {
                _logger($"[SystemTypeCatalog] Unexpected error: {ex.Message}");
            }

            return catalog;
        }

        /// <summary>
        /// ✅ PERFORMANCE FIX: Loads system types from SQLite database (no Revit API calls)
        /// Uses "dumb once, use many times" principle - data was collected during refresh and stored in DB
        /// </summary>
        private void LoadSystemTypesFromDatabase(Document document, SystemTypeCatalog catalog)
        {
            try
            {
                using (var context = new Data.SleeveDbContext(document, msg =>
                {
                    _logger($"[SystemTypeCatalog][SQLite] {msg}");
                }))
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        // Query all MEP parameters from SleeveSnapshots table
                        cmd.CommandText = @"
                            SELECT DISTINCT MepParametersJson 
                            FROM SleeveSnapshots 
                            WHERE MepParametersJson IS NOT NULL 
                              AND MepParametersJson != '{}'
                              AND MepParametersJson != ''";

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    var jsonText = reader["MepParametersJson"]?.ToString();
                                    if (string.IsNullOrWhiteSpace(jsonText))
                                        continue;

                                    // Parse JSON dictionary
                                    var mepParams = DeserializeDictionary(jsonText);
                                    if (mepParams == null || mepParams.Count == 0)
                                        continue;

                                    // Extract System Type (for Ducts, Pipes)
                                    var systemType = GetParameterValue(mepParams, "System Type", "MEP System Type", "System Classification");
                                    if (!string.IsNullOrWhiteSpace(systemType))
                                    {
                                        catalog.SystemTypes.Add(systemType.Trim());
                                    }

                                    // ✅ Extract Service Type (for Cable Trays, Conduits) - separate from System Type
                                    var serviceType = GetParameterValue(mepParams, "Service Type", "MEP Service Type");
                                    if (!string.IsNullOrWhiteSpace(serviceType))
                                    {
                                        catalog.ServiceTypes.Add(serviceType.Trim());
                                    }

                                    // Also extract System Abbreviation (can be used as service type fallback)
                                    var systemAbbr = GetParameterValue(mepParams, "System Abbreviation");
                                    if (!string.IsNullOrWhiteSpace(systemAbbr))
                                    {
                                        catalog.ServiceTypes.Add(systemAbbr.Trim());
                                    }
                                }
                                catch (Exception jsonEx)
                                {
                                    _logger($"[SystemTypeCatalog] Error parsing MepParametersJson: {jsonEx.Message}");
                                }
                            }
                        }
                    }
                }

                _logger($"[SystemTypeCatalog] ✅ Loaded {catalog.SystemTypes.Count} unique system types and {catalog.ServiceTypes.Count} unique service types from database");
            }
            catch (Exception ex)
            {
                _logger($"[SystemTypeCatalog] Error loading from database: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper to get parameter value from dictionary with multiple key options
        /// </summary>
        private static string GetParameterValue(Dictionary<string, string> dict, params string[] keys)
        {
            if (dict == null || keys == null)
                return null;

            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (dict.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return null;
        }

        /// <summary>
        /// Deserialize JSON dictionary using System.Text.Json (same as ClashZoneRepository)
        /// </summary>
        private static Dictionary<string, string> DeserializeDictionary(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json == "{}")
                return new Dictionary<string, string>();

            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(json) 
                    ?? new Dictionary<string, string>();
            }
            catch
            {
                return new Dictionary<string, string>();
            }
        }

        private static string ExtractParameterValue(ClashZone zone, params string[] keys)
        {
            if (zone?.MepParameterValues == null || zone.MepParameterValues.Count == 0 || keys == null)
                return null;

            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                var match = zone.MepParameterValues
                    .FirstOrDefault(kv => kv != null &&
                                           kv.Key != null &&
                                           kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

                if (match != null && !string.IsNullOrWhiteSpace(match.Value))
                {
                    return match.Value;
                }
            }

            return null;
        }
    }
}

