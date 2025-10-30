using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for reading parameter values from stored XML files
    /// </summary>
    public class ParameterValueReader
    {
        /// <summary>
        /// Reads parameter values from XML files for specified openings
        /// </summary>
        public Dictionary<ElementId, Dictionary<string, string>> ReadParameterValuesFromXml(
            List<ElementId> openingIds,
            string parameterName)
        {
            var parameterValues = new Dictionary<ElementId, Dictionary<string, string>>();

            try
            {
                DebugLogger.Info($"[ParameterValueReader] Reading {parameterName} values for {openingIds.Count} openings from XML");

                // Look for XML files in the standard locations
                var xmlFiles = FindParameterXmlFiles(parameterName);

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var values = ParseParameterXmlFile(xmlFile, parameterName);

                        // Match values to opening IDs
                        foreach (var openingId in openingIds)
                        {
                            if (values.TryGetValue(openingId, out var value))
                            {
                                if (!parameterValues.ContainsKey(openingId))
                                {
                                    parameterValues[openingId] = new Dictionary<string, string>();
                                }
                                parameterValues[openingId][parameterName] = value;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[ParameterValueReader] Error reading XML file {xmlFile}: {ex.Message}");
                    }
                }

                DebugLogger.Info($"[ParameterValueReader] Found values for {parameterValues.Count} openings");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterValueReader] Error reading parameter values: {ex.Message}");
            }

            return parameterValues;
        }

        private List<string> FindParameterXmlFiles(string parameterName)
        {
            var xmlFiles = new List<string>();

            try
            {
                // Look in multiple locations
                var searchPaths = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Parameters"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Parameters"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Parameters")
                };

                foreach (var searchPath in searchPaths)
                {
                    if (Directory.Exists(searchPath))
                    {
                        try
                        {
                            var files = Directory.EnumerateFiles(searchPath, "*.xml", SearchOption.AllDirectories);
                            xmlFiles.AddRange(files);
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[ParameterValueReader] Error enumerating XML files in {searchPath}: {ex.Message}");
                        }
                    }
                }

                DebugLogger.Info($"[ParameterValueReader] Found {xmlFiles.Count} XML files to scan for parameter '{parameterName}'");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterValueReader] Error searching for XML files: {ex.Message}");
            }

            return xmlFiles;
        }

        private Dictionary<ElementId, string> ParseParameterXmlFile(string xmlFile, string parameterName)
        {
            var result = new Dictionary<ElementId, string>();

            try
            {
                // Attempt to deserialize as OpeningFilter, which carries ClashZoneStorage
                var serializer = new XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFile))
                {
                    var filter = serializer.Deserialize(reader) as OpeningFilter;
                    var zones = filter?.ClashZoneStorage?.ClashZones;
                    if (zones == null || zones.Count == 0) return result;

                    foreach (var cz in zones)
                    {
                        // Use SleeveInstanceId as the opening identifier if present
                        var openingIdInt = cz.SleeveInstanceId;
                        if (openingIdInt <= 0) continue;

                        // Prefer MEP bag; fall back to Host bag
                        var value = TryGetParamValueFromBags(cz.MepParameterValues, parameterName);
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            value = TryGetParamValueFromBags(cz.HostParameterValues, parameterName);
                        }

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            var eid = new ElementId(openingIdInt);
                            if (!result.ContainsKey(eid)) result[eid] = value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[ParameterValueReader] Failed to parse {Path.GetFileName(xmlFile)}: {ex.Message}");
            }

            return result;
        }

        private static string TryGetParamValueFromBags(List<SerializableKeyValue> bag, string parameterName)
        {
            if (bag == null || bag.Count == 0 || string.IsNullOrWhiteSpace(parameterName)) return string.Empty;
            var kv = bag.FirstOrDefault(k => string.Equals(k.Key, parameterName, StringComparison.OrdinalIgnoreCase));
            return kv?.Value ?? string.Empty;
        }
    }
}
