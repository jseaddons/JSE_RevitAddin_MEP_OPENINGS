using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// OOP service responsible for building parameter whitelists and capturing element parameter snapshots.
    /// Keeps Refresh integration minimal and isolated.
    /// </summary>
    public class ParameterSnapshotService
    {
        // ✅ FIX 1: Essential parameters whitelist - only capture these to reduce memory by 90%
        private static readonly HashSet<string> ESSENTIAL_PARAMETERS = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            // MEP Element essentials
            "System Name", "System Abbreviation", "System Type",
            "MEP System Name", "MEP System Abbreviation", "MEP System Type", "MEP System Classification",
            "MEP Size",
            "Width", "Height", "Diameter", "Size",
            "Level", "Offset",
            "Insulation Thickness",
            
            // Host essentials  
            "Family", "Family Name",
            "Width", "Thickness", "Height",
            "Structural", "Function",
            "Level", "Base Offset", "Top Offset",
            
            // Common
            "Mark", "Comments", "Phase Created",
            
            // Legacy compatibility (from _commonMepKeys and _commonHostKeys)
            "Nominal Diameter", "Outside Diameter",
            "Reference Level", "Schedule Level", "Reference Level Elevation",
            "System Classification", "Service Type",
            "Fire Rating"
        };

        private readonly ISet<string> _commonMepKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Size","Diameter","Nominal Diameter","Outside Diameter","Width","Height",
            "Reference Level","Level","Schedule Level","Reference Level Elevation",
            "System Type","System Classification","Service Type","System Abbreviation",
            "MEP System Type","MEP System Name","MEP System Abbreviation","MEP Size"
        };

        private readonly ISet<string> _commonHostKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Per user request: only capture Fire Rating for host elements
            "Fire Rating"
        };

        /// <summary>
        /// Build a whitelist for the current run by combining curated keys and previously learned keys from storage.
        /// ✅ MEMORY OPTIMIZATION: Cap learned parameters at 20 to prevent unbounded growth.
        /// </summary>
        public HashSet<string> BuildWhitelist(ClashZoneStorage storage, IEnumerable<(Element mep, Element host)> sample)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var k in _commonMepKeys) keys.Add(k);
            foreach (var k in _commonHostKeys) keys.Add(k);

            if (storage?.ParameterKeyWhitelist != null)
            {
                // ✅ MEMORY OPTIMIZATION: Limit to first 50 user-defined parameters
                var limitedWhitelist = storage.ParameterKeyWhitelist.Take(50).ToList();
                foreach (var k in limitedWhitelist) keys.Add(k);
            }
            if (storage?.LearnedParameterKeys != null)
            {
                // ✅ MEMORY OPTIMIZATION: Cap learned parameters at 20 to prevent unbounded growth
                var limitedLearned = storage.LearnedParameterKeys.Take(20).ToList();
                foreach (var k in limitedLearned) keys.Add(k);
            }

            // Merge disk-learned keys (project-level) - also limit these
            var diskKeys = LoadLearnedKeysFromDisk();
            foreach (var k in diskKeys.Take(20)) keys.Add(k);

            // ✅ MEMORY OPTIMIZATION: Log total whitelist size for debugging
            if (keys.Count > 30)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PARAM_SNAPSHOT] ⚠️ Large parameter whitelist detected: {keys.Count} parameters. This may increase memory usage significantly.");
            }

            return keys;
        }

        /// <summary>
        /// Capture whitelisted parameter values for an element (tries instance, then type if missing).
        /// ✅ FIX 1 & 6: Only capture ESSENTIAL parameters to reduce memory by 90%.
        /// </summary>
        public List<SerializableKeyValue> CaptureParams(Element element, HashSet<string> whitelist)
        {
            var result = new List<SerializableKeyValue>();
            if (element == null || whitelist == null || whitelist.Count == 0) return result;
            
            // ✅ FIX 6: Emergency parameter limit
            const int MAX_PARAMETERS = 30;
            
            // DEBUG: Log all available parameters for duct accessories
            if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Starting parameter capture\n");
                
                var allParams = element.Parameters.Cast<Parameter>().Where(p => p != null && !string.IsNullOrEmpty(p.Definition?.Name)).ToList();
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Found {allParams.Count} total parameters\n");
                
                foreach (var param in allParams.Take(10)) // Log first 10 parameters
                {
                    var paramValue = ConvertParameterToString(element, param);
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Parameter '{param.Definition.Name}' = '{paramValue}'\n");
                }
            }

            foreach (var key in whitelist)
            {
                // ✅ FIX 6: Emergency brake - stop if limit reached
                if (result.Count >= MAX_PARAMETERS)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[PARAM_SNAPSHOT] Parameter limit ({MAX_PARAMETERS}) reached for element {element.Id}");
                    break;
                }

                // ✅ FIX 1: Only capture ESSENTIAL parameters OR user-defined/learned parameters from whitelist
                // User-defined params (from ParameterKeyWhitelist/LearnedParameterKeys) should always be captured
                // Essential params are always captured
                // Common system params (from _commonMepKeys/_commonHostKeys) are only captured if they're in ESSENTIAL_PARAMETERS
                bool isEssential = ESSENTIAL_PARAMETERS.Contains(key);
                bool isCommonKey = _commonMepKeys.Contains(key) || _commonHostKeys.Contains(key);
                
                // Only capture if:
                // 1. It's an essential parameter, OR
                // 2. It's a user-defined/learned parameter (not in common keys)
                if (!isEssential && isCommonKey)
                {
                    continue; // Skip common system params that aren't essential
                }

                var p = LookupParam(element, key);
                
                // Special fallback for System Type when not found by name
                if (p == null && key.Equals("System Type", StringComparison.OrdinalIgnoreCase))
                {
                    // Try built-in parameters for ducts/pipes
                    p = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) ??
                        element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                        element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                }
                // Special fallback for System Abbreviation when not found by name
                if (p == null && key.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase))
                {
                    // Try built-in parameters for system abbreviation
                    // System Abbreviation is typically a custom parameter, not a built-in parameter
                    // Try common parameter names for system abbreviation
                    p = element.LookupParameter("System Abbreviation") ??
                        element.LookupParameter("System Abbr") ??
                        element.LookupParameter("Abbreviation") ??
                        element.LookupParameter("Abbr");
                    
                    // DEBUG: Log System Abbreviation search for duct accessories
                    if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: System Abbreviation fallback search - Parameter found: {p != null}\n");
                        
                        if (p != null)
                        {
                            var testValue = ConvertParameterToString(element, p);
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: System Abbreviation value = '{testValue}'\n");
                        }
                    }
                }
                
                if (p == null) 
                {
                    // DEBUG: Log missing parameters
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Parameter '{key}' not found\n");
                    continue;
                }

                var value = ConvertParameterToString(element, p);
                if (string.IsNullOrWhiteSpace(value)) 
                {
                    // DEBUG: Log empty parameter values
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Parameter '{key}' found but value is empty\n");
                    continue;
                }

                // ✅ MEMORY OPTIMIZATION: Truncate very long parameter values to prevent memory bloat
                // Long parameter values (e.g., comments, descriptions) can be 500+ bytes
                const int MAX_PARAM_VALUE_LENGTH = 200; // Cap at 200 bytes to prevent bloat
                if (value != null && value.Length > MAX_PARAM_VALUE_LENGTH)
                {
                    value = value.Substring(0, MAX_PARAM_VALUE_LENGTH) + "...[truncated]";
                }
                
                // ✅ FIX 1: Intern strings to share memory across clash zones
                result.Add(new SerializableKeyValue 
                { 
                    Key = string.Intern(key), 
                    Value = string.Intern(value) 
                });
                
                // DEBUG: Log successful parameter capture (only in non-deployment mode)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Captured '{key}' = '{value}'\n");
                }
            }

            return result;
        }

        /// <summary>
        /// Convert a parameter value to a robust invariant string.
        /// </summary>
        private string ConvertParameterToString(Element owner, Parameter p)
        {
            if (p == null) return string.Empty;

            string value = p.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            value = p.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            switch (p.StorageType)
            {
                case StorageType.Integer:
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                case StorageType.Double:
                    // ✅ FIX 1: Round doubles to 3 decimals to reduce string size
                    return Math.Round(p.AsDouble(), 3).ToString(CultureInfo.InvariantCulture);
                case StorageType.ElementId:
                    try
                    {
                        var id = p.AsElementId();
                        if (id == null) return string.Empty;
                        // Prefer referenced element name for readability if available
                        var e = owner?.Document?.GetElement(id);
                        var name = e?.Name;
                        if (!string.IsNullOrWhiteSpace(name)) return name;
                        return id.IntegerValue.ToString(CultureInfo.InvariantCulture);
                    }
                    catch { return string.Empty; }
                default:
                    return string.Empty;
            }
        }

        private Parameter LookupParam(Element e, string key)
        {
            var p = e.LookupParameter(key);
            if (p != null) return p;

            if (e is FamilyInstance fi)
            {
                var sp = fi.Symbol?.LookupParameter(key);
                if (sp != null) return sp;
            }

            return null;
        }

        /// <summary>
        /// Utility to produce a doc key (host vs linked distinction).
        /// </summary>
        public string GetDocKey(Element e)
        {
            try { return e?.Document?.PathName ?? "Unknown"; }
            catch { return "Unknown"; }
        }

        // === Learned Keys (Project-level Persistence) ===
        private static string GetLearnedKeysFilePath()
        {
            // Use null for document since this is a static method - will use default path
            var dir = ProjectPathService.GetFiltersDirectory(null);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "learned_parameter_keys.xml");
        }

        public IEnumerable<string> LoadLearnedKeysFromDisk()
        {
            try
            {
                var file = GetLearnedKeysFilePath();
                if (!File.Exists(file)) return Enumerable.Empty<string>();

                var doc = new System.Xml.XmlDocument();
                doc.Load(file);
                var nodes = doc.SelectNodes("/LearnedParameterKeys/Key");
                var list = new List<string>();
                if (nodes != null)
                {
                    foreach (System.Xml.XmlNode n in nodes)
                    {
                        var v = n.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(v)) list.Add(v);
                    }
                }
                return list;
            }
            catch { return Enumerable.Empty<string>(); }
        }

        public static void AddLearnedKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            try
            {
                var file = GetLearnedKeysFilePath();
                var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (File.Exists(file))
                {
                    var doc = new System.Xml.XmlDocument();
                    doc.Load(file);
                    var nodes = doc.SelectNodes("/LearnedParameterKeys/Key");
                    if (nodes != null)
                    {
                        foreach (System.Xml.XmlNode n in nodes)
                        {
                            var v = n.InnerText?.Trim();
                            if (!string.IsNullOrEmpty(v)) keys.Add(v);
                        }
                    }
                }

                if (!keys.Contains(key)) keys.Add(key);

                // Write out
                var xml = new System.Xml.XmlDocument();
                var root = xml.CreateElement("LearnedParameterKeys");
                xml.AppendChild(root);
                foreach (var k in keys.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                {
                    var e = xml.CreateElement("Key");
                    e.InnerText = k;
                    root.AppendChild(e);
                }
                xml.Save(file);
            }
            catch { /* non-fatal */ }
        }
    }
}


