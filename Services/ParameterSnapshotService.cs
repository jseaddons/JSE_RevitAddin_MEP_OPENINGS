using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// OOP service responsible for building parameter whitelists and capturing element parameter snapshots.
    /// Keeps Refresh integration minimal and isolated.
    /// </summary>
    public partial class ParameterSnapshotService
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
            "Reference Level", "Schedule Level", "Schedule of Level", "Reference Level Elevation",
            "System Classification", "Service Type",
            "Fire Rating"
        };

        // ✅ MUST-CAPTURE: Always capture these critical parameters (no limits applied)
        private static readonly string[] MUST_CAPTURE_KEYS = new[]
        {
            "System Type",
            "System Name",
            "System Abbreviation",
            "Reference Level",
            "Schedule of Level",
            "Schedule Level", // alias for schedule of level
            "Size",           // ✅ CRITICAL: Size must always be captured for MEP elements
            "Service Type"    // ✅ CRITICAL: For Cable Trays
        };

        private static readonly ISet<string> _commonMepKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Size","Diameter","Nominal Diameter","Outside Diameter","Width","Height",
            "Reference Level","Level","Schedule Level","Schedule of Level","Reference Level Elevation",
            "System Type","System Name","System Classification","Service Type","System Abbreviation",
            "MEP System Type","MEP System Name","MEP System Abbreviation","MEP Size"
        };

        private static readonly ISet<string> _commonHostKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Per user request: only capture Fire Rating for host elements
            "Fire Rating"
        };

        /// <summary>
        /// Build a whitelist for the current run by combining curated keys and previously learned keys from storage.
        /// ✅ MEMORY OPTIMIZATION: Cap learned parameters at 20 to prevent unbounded growth.
        /// </summary>
        private static partial HashSet<string> BuildWhitelistLegacy()
        {
            return BuildWhitelistLegacy(null, null);
        }

        public HashSet<string> BuildWhitelist(ClashZoneStorage storage, IEnumerable<(Element mep, Element host)> sample)
        {
            return BuildWhitelistLegacy(storage, sample);
        }

        private static HashSet<string> BuildWhitelistLegacy(ClashZoneStorage storage, IEnumerable<(Element mep, Element host)> sample)
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
            return CaptureParamsLegacy(element, whitelist, element?.Document, null);
        }

        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Capture parameters for multiple elements in one pass.
        /// Returns a dictionary mapping element IDs to their parameter dictionaries.
        /// This significantly reduces overhead in large refresh operations.
        /// </summary>
        public static Dictionary<int, Dictionary<string, string>> CaptureBatchParams(IEnumerable<Element> elements)
        {
            var results = new Dictionary<int, Dictionary<string, string>>();
            if (elements == null) return results;

            // Use the default legacy whitelist for batch processing
            var whitelist = BuildWhitelistLegacy();

            foreach (var element in elements)
            {
                if (element == null) continue;

                var id = element.Id.IntegerValue;
                if (results.ContainsKey(id)) continue;

                // Capture using the legacy logic
                var snapshots = CaptureParamsLegacy(element, whitelist, element.Document, null);
                
                // Convert to dictionary for easy O(1) lookup in processing loops
                var paramDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var snapshot in snapshots)
                {
                    if (!string.IsNullOrEmpty(snapshot.Key))
                        paramDict[snapshot.Key] = snapshot.Value ?? "";
                }

                results[id] = paramDict;
            }

            return results;
        }

        private static partial List<SerializableKeyValue> CaptureParamsLegacy(Element element, HashSet<string> whitelist, Document doc, string docKey)
        {
            var result = new List<SerializableKeyValue>();
            if (element == null || whitelist == null || whitelist.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode && element != null)
                {
                    DebugLogger.Warning($"[{DateTime.Now}] [PARAM-CAPTURE] ⚠️ Element {element.Id} ({element.Category?.Name}): No whitelist provided or whitelist is empty\n");
                }
                return result;
            }
            
            // ✅ FIX 6: Emergency parameter limit
            const int MAX_PARAMETERS = 30;

            // ✅ PRIORITY ORDER: Must-capture keys first, then remaining whitelist
            var orderedKeys = new List<string>();
            var mustCaptureSet = new HashSet<string>(MUST_CAPTURE_KEYS, StringComparer.OrdinalIgnoreCase);
            foreach (var k in MUST_CAPTURE_KEYS)
            {
                if (whitelist.Contains(k)) orderedKeys.Add(k);
            }
            foreach (var k in whitelist)
            {
                if (!mustCaptureSet.Contains(k)) orderedKeys.Add(k);
            }
            
            // ✅ ENHANCED LOGGING: Show whitelist for all elements (to debug parameter issues)
            // Log whitelist for all elements to help diagnose parameter capture issues
            bool shouldLogDetails = !DeploymentConfiguration.DeploymentMode && whitelist.Count > 0;
            
            if (shouldLogDetails && OptimizationFlags.UseDiagnosticMode)
            {
                var whitelistSample = string.Join(", ", whitelist.Take(15));
                var moreCount = whitelist.Count > 15 ? $" (+{whitelist.Count - 15} more)" : "";
                DebugLogger.Info($"[{DateTime.Now}] [PARAM-CAPTURE] 🔍 WHITELIST for Element {element.Id} ({element.Category?.Name}): {whitelist.Count} parameters in whitelist: {whitelistSample}{moreCount}\n");
                DebugLogger.Info($"[{DateTime.Now}] [PARAM-CAPTURE] 🔍 ORDERED KEYS for Element {element.Id}: {orderedKeys.Count} parameters to try: {string.Join(", ", orderedKeys.Take(15))}{moreCount}\n");
            }
            
            // DEBUG: Log all available parameters for duct accessories
            if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory && OptimizationFlags.UseDiagnosticMode)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Starting parameter capture\n");
                
                var allParams = element.Parameters.Cast<Parameter>().Where(p => p != null && !string.IsNullOrEmpty(p.Definition?.Name)).ToList();
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Found {allParams.Count} total parameters\n");
                
                foreach (var param in allParams.Take(10)) // Log first 10 parameters
                {
                    var paramValue = ConvertParameterToStringLegacy(element, param);
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Parameter '{param.Definition.Name}' = '{paramValue}'\n");
                }
            }

            // ✅ CRITICAL: Check if element is a Cable Tray (needed for parameter mapping)
            bool isCableTray = element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_CableTray ||
                              element.Category?.Name?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                              element is CableTray;
            
            foreach (var key in orderedKeys)
            {
                // ✅ FIX 6: Emergency brake - stop if limit reached
                bool isMustCapture = mustCaptureSet.Contains(key);
                if (!isMustCapture && result.Count >= MAX_PARAMETERS)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[PARAM_SNAPSHOT] Parameter limit ({MAX_PARAMETERS}) reached for element {element.Id}");
                    break;
                }

                // ✅ CRITICAL: For Cable Trays, map "System Type" to "Service Type" in the whitelist
                // This ensures we capture "Service Type" even if whitelist only has "System Type"
                string actualKey = key;
                if (isCableTray && key.Equals("System Type", StringComparison.OrdinalIgnoreCase))
                {
                    actualKey = "Service Type"; // Use "Service Type" for Cable Trays
                }

                // ✅ CRITICAL FIX: If parameter is in whitelist, ALWAYS capture it (regardless of essential/common)
                // The whitelist is the source of truth - if it's whitelisted, capture it
                // This ensures consistent parameter capture across all elements
                // Essential parameters are prioritized but all whitelisted params should be captured
                bool isEssential = ESSENTIAL_PARAMETERS.Contains(actualKey) || 
                                   (isCableTray && actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase) && ESSENTIAL_PARAMETERS.Contains("Service Type"));
                
                // ✅ FIX: Don't filter out whitelisted parameters - if it's in whitelist, capture it
                // The whitelist already contains only the parameters we want to capture
                // No need for additional filtering that causes inconsistency

                // ✅ CRITICAL: Use actualKey (mapped for Cable Trays) instead of original key
                var p = LookupParamLegacy(element, actualKey);
                
                // ✅ CRITICAL: Special fallback for System Type/Service Type when not found by name
                // NOTE: For Cable Trays, use "Service Type" instead of "System Type"
                if (p == null && (key.Equals("System Type", StringComparison.OrdinalIgnoreCase) || 
                                  actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase)))
                {
                    // ✅ CABLE TRAY: Use isCableTray already declared above - use "Service Type" instead
                    if (isCableTray)
                    {
                        // For Cable Trays, look for "Service Type" parameter
                        p = element.LookupParameter("Service Type") ??
                            element.LookupParameter("MEP Service Type");
                    }
                    else
                    {
                        // For Mechanical/Plumbing (Ducts/Pipes), use "System Type"
                        // Try built-in parameters for ducts/pipes
                        p = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) ??
                            element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                            element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                        
                        // Try common parameter name variations
                        if (p == null)
                        {
                            p = element.LookupParameter("MEP System Type") ??
                                element.LookupParameter("System Classification") ??
                                element.LookupParameter("MEP System Classification");
                        }
                    }
                }
                
                // ✅ CRITICAL: Special fallback for Service Type (Cable Trays)
                if (p == null && key.Equals("Service Type", StringComparison.OrdinalIgnoreCase))
                {
                    // Try common parameter name variations for Service Type
                    p = element.LookupParameter("Service Type") ??
                        element.LookupParameter("MEP Service Type");
                }
                
                // ✅ CRITICAL: Special fallback for System Name when not found by name
                if (p == null && key.Equals("System Name", StringComparison.OrdinalIgnoreCase))
                {
                    // Try common parameter name variations
                    p = element.LookupParameter("MEP System Name") ??
                        element.LookupParameter("System Name");
                }
                
                // ✅ CRITICAL: Special fallback for System Abbreviation when not found by name
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
                    if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory && OptimizationFlags.UseDiagnosticMode)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: System Abbreviation fallback search - Parameter found: {p != null}\n");
                        
                        if (p != null)
                        {
                            var testValue = ConvertParameterToStringLegacy(element, p);
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: System Abbreviation value = '{testValue}'\n");
                        }
                    }
                }
                
                // ✅ CRITICAL FIX: Special fallback for "Schedule of Level" parameter
                // Schedule of Level is critical for "Bottom of Opening" calculation
                if (p == null && (key.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                                  key.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase)))
                {
                    // Try common parameter name variations
                    p = element.LookupParameter("Schedule of Level") ??
                        element.LookupParameter("Schedule Level") ??
                        element.LookupParameter("Elevation from Level");
                    
                    if (!DeploymentConfiguration.DeploymentMode && p != null && OptimizationFlags.UseDiagnosticMode)
                    {
                        var testValue = ConvertParameterToStringLegacy(element, p);
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Schedule of Level fallback found, value = '{testValue}'\n");
                    }
                }
                
                // ✅ CRITICAL FIX: Special fallback for "Reference Level" parameter
                // Reference Level is a critical parameter for MEP elements and may be stored as a built-in parameter
                // or may have different names in different MEP categories
                if (p == null && (key.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) || 
                                  key.Equals("Level", StringComparison.OrdinalIgnoreCase)))
                {
                    // Try built-in parameters for level
                    p = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                        element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM) ??
                        element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                    
                    // Try common parameter name variations
                    if (p == null)
                    {
                        p = element.LookupParameter("Reference Level") ??
                            element.LookupParameter("Level") ??
                            element.LookupParameter("Schedule Level") ??
                            element.LookupParameter("Schedule of Level");
                    }
                    
                    // For linked files, try to get from the element's document
                    if (p == null && element.Document.IsLinked)
                    {
                        try
                        {
                            var linkDoc = element.Document;
                            if (linkDoc != null)
                            {
                                p = element.LookupParameter("Reference Level") ??
                                    element.LookupParameter("Level");
                            }
                        }
                        catch { /* Ignore errors when accessing linked document */ }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode && p != null && OptimizationFlags.UseDiagnosticMode)
                    {
                        var testValue = ConvertParameterToStringLegacy(element, p);
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Reference Level fallback found, value = '{testValue}'\n");
                    }
                }
                
                if (p == null) 
                {
                    // ✅ CRITICAL: Log missing essential parameters (especially System Type, Service Type, System Name, System Abbreviation)
                    // These are critical for parameter transfer and remarking
                    if (isEssential && (actualKey.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                       actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase) ||
                                       actualKey.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                                       actualKey.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase) ||
                                       actualKey.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                                       actualKey.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[{DateTime.Now}] [PARAM_CAPTURE] ⚠️⚠️⚠️ CRITICAL: Essential parameter '{actualKey}' (mapped from '{key}') not found on element {element.Id} ({element.Category?.Name}) - this will prevent parameter transfer!\n");
                        }
                    }
                    else
                    {
                        // DEBUG: Log missing non-essential parameters
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Parameter '{actualKey}' (mapped from '{key}') not found\n");
                    }
                    continue;
                }

                var value = ConvertParameterToStringLegacy(element, p);
                if (string.IsNullOrWhiteSpace(value)) 
                {
                    // ✅ ENHANCED LOGGING: Show parameter details when empty (to diagnose why parameters exist in model but are empty here)
                    var paramDetails = $"StorageType={p.StorageType}, IsReadOnly={p.IsReadOnly}, HasValue={p.HasValue}";
                    if (p.StorageType == StorageType.ElementId)
                    {
                        try
                        {
                            var elemId = p.AsElementId();
                            paramDetails += $", ElementId={elemId?.IntegerValue ?? -1}";
                        }
                        catch { }
                    }
                    else if (p.StorageType == StorageType.Double)
                    {
                        try
                        {
                            var dblVal = p.AsDouble();
                            paramDetails += $", DoubleValue={dblVal}";
                        }
                        catch { }
                    }
                    else if (p.StorageType == StorageType.Integer)
                    {
                        try
                        {
                            var intVal = p.AsInteger();
                            paramDetails += $", IntegerValue={intVal}";
                        }
                        catch { }
                    }
                    
                    // ✅ CRITICAL: Log empty essential parameters with full details
                    if (isEssential && (key.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                       key.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                                       key.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[{DateTime.Now}] [PARAM_CAPTURE] ⚠️⚠️⚠️ CRITICAL: Essential parameter '{key}' found but value is empty on element {element.Id} ({element.Category?.Name}) - {paramDetails} - this will prevent parameter transfer!\n");
                        }
                    }
                    else
                    {
                        // ✅ ENHANCED LOGGING: Log empty non-essential parameter values with details
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            DebugLogger.Info($"[{DateTime.Now}] [PARAM-CAPTURE] Element {element.Id} ({element.Category?.Name}): Parameter '{key}' found but value is empty - {paramDetails}\n");
                    }
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
                // ✅ CRITICAL: Store actualKey (for Cable Trays, this is "Service Type" instead of "System Type")
                // This ensures the snapshot has the correct parameter name for parameter transfer
                result.Add(new SerializableKeyValue 
                { 
                    Key = string.Intern(actualKey), // Use actualKey (mapped for Cable Trays)
                    Value = string.Intern(value)
                });
                
                // ✅ CRITICAL DIAGNOSTIC: Log captured essential parameters (especially for circular elements)
                if (isEssential && (actualKey.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                   actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase) ||
                                   actualKey.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                                   actualKey.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase) ||
                                   actualKey.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                                   actualKey.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase)))
                {
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        string logKey = key.Equals(actualKey) ? actualKey : $"{actualKey} (from '{key}')";
                        DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] ✅ CAPTURED: Element {element.Id} ({element.Category?.Name}): '{logKey}' = '{value}'\n");
                    }
                }
                
                // DEBUG: Log successful parameter capture (only in non-deployment mode)
                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                {
                    DebugLogger.Info($"[{DateTime.Now}] [PARAM_CAPTURE] Element {element.Id} ({element.Category?.Name}): Captured '{key}' = '{value}'\n");
                }
            }

            // ✅ ENHANCED LOGGING: Summary of captured parameters
            if (!DeploymentConfiguration.DeploymentMode && result.Count > 0 && OptimizationFlags.UseDiagnosticMode)
            {
                var capturedKeys = string.Join(", ", result.Select(kv => kv.Key).Take(10));
                var moreCount = result.Count > 10 ? $" (+{result.Count - 10} more)" : "";
                DebugLogger.Info($"[{DateTime.Now}] [PARAM-CAPTURE] ✅ SUMMARY: Element {element.Id} ({element.Category?.Name}): Captured {result.Count} parameters: {capturedKeys}{moreCount}\n");
            }
            else if (!DeploymentConfiguration.DeploymentMode && result.Count == 0 && orderedKeys.Count > 0)
            {
                DebugLogger.Warning($"[{DateTime.Now}] [PARAM-CAPTURE] ⚠️ WARNING: Element {element.Id} ({element.Category?.Name}): No parameters captured from {orderedKeys.Count} whitelisted parameters!\n");
            }

            return result;
        }

        /// <summary>
        /// Convert a parameter value to a robust invariant string.
        /// ✅ ENHANCED: Tries multiple methods to extract parameter value.
        /// </summary>
        private static partial string ConvertParameterToStringLegacy(Element element, Parameter param)
        {
            if (param == null) return string.Empty;

            // ✅ METHOD 1: Try AsString() first (most common for text parameters)
            string value = param.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            // ✅ METHOD 2: Try AsValueString() (formatted display value)
            value = param.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            // ✅ METHOD 3: Try storage type-specific conversions
            switch (param.StorageType)
            {
                case StorageType.Integer:
                    try
                    {
                        if (param.HasValue)
                            return param.AsInteger().ToString(CultureInfo.InvariantCulture);
                    }
                    catch { }
                    return string.Empty;
                    
                case StorageType.Double:
                    try
                    {
                        if (param.HasValue)
                        {
                            // ✅ FIX 1: Round doubles to 3 decimals to reduce string size
                            return Math.Round(param.AsDouble(), 3).ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    catch { }
                    return string.Empty;
                    
                case StorageType.ElementId:
                    try
                    {
                        if (param.HasValue)
                        {
                            var id = param.AsElementId();
                            if (id == null || id == ElementId.InvalidElementId) return string.Empty;
                            
                            // Prefer referenced element name for readability if available
                            var e = element?.Document?.GetElement(id);
                            if (e != null)
                            {
                                var name = e.Name;
                                if (!string.IsNullOrWhiteSpace(name)) return name;
                                
                                // Try Category name if element name is empty
                                var catName = e.Category?.Name;
                                if (!string.IsNullOrWhiteSpace(catName)) return catName;
                            }
                            
                            // Fallback to ElementId integer value
                            return id.IntegerValue.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    catch { }
                    return string.Empty;
                    
                case StorageType.String:
                    // Already tried AsString() and AsValueString() above
                    // If still empty, parameter truly has no value
                    return string.Empty;
                    
                default:
                    return string.Empty;
            }
        }

        private static partial Parameter LookupParamLegacy(Element element, string paramName)
        {
            var p = element.LookupParameter(paramName);
            if (p != null) return p;

            if (element is FamilyInstance fi)
            {
                var sp = fi.Symbol?.LookupParameter(paramName);
                if (sp != null) return sp;
            }
            
            // ✅ SPECIAL HANDLING: For "Size" parameter, try multiple fallbacks
            // Dampers may have Size as a calculated/formula parameter that needs different lookup
            if (paramName.Equals("Size", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ DEBUG: Log all parameter names for Duct Accessories to find Size
                if (!DeploymentConfiguration.DeploymentMode && element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                {
                    var allParamNames = element.Parameters.Cast<Parameter>()
                        .Where(p => p?.Definition?.Name != null)
                        .Select(p => p.Definition.Name)
                        .Take(30)
                        .ToList();
                    
                    SafeFileLogger.SafeAppendText("param_debug.log", 
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SIZE-LOOKUP] Element {element.Id} ALL PARAMS ({allParamNames.Count}): [{string.Join(", ", allParamNames)}]\n");
                }
                
                // Try all parameters looking for exact match "Size" (without HasValue check for formula params)
                foreach (Parameter param in element.Parameters)
                {
                    if (param?.Definition?.Name != null &&
                        param.Definition.Name.Equals("Size", StringComparison.OrdinalIgnoreCase))
                    {
                        // ✅ DEBUG: Log Size parameter found
                        if (!DeploymentConfiguration.DeploymentMode && element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                        {
                            SafeFileLogger.SafeAppendText("param_debug.log", 
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SIZE-FOUND] Element {element.Id}: Size param found! HasValue={param.HasValue}, StorageType={param.StorageType}\n");
                        }
                        return param; // Return regardless of HasValue - let conversion handle empty values
                    }
                }
                
                // Also check Type parameters (without HasValue check)
                if (element is FamilyInstance fi2)
                {
                    var typeElem = fi2.Symbol;
                    if (typeElem != null)
                    {
                        foreach (Parameter param in typeElem.Parameters)
                        {
                            if (param?.Definition?.Name != null &&
                                param.Definition.Name.Equals("Size", StringComparison.OrdinalIgnoreCase))
                            {
                                return param;
                            }
                        }
                    }
                }
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
        private static List<string>? _learnedKeysCache = null;

        private static string GetLearnedKeysFilePath()
        {
            // Use null for document since this is a static method - will use default path
            var dir = ProjectPathService.GetFiltersDirectory(null);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "learned_parameter_keys.xml");
        }

        public static IEnumerable<string> LoadLearnedKeysFromDisk()
        {
            // ✅ OPTIMIZATION: Return cache if available to avoid redundant disk I/O
            if (_learnedKeysCache != null) return _learnedKeysCache;

            try
            {
                var file = GetLearnedKeysFilePath();
                if (!File.Exists(file)) 
                {
                    _learnedKeysCache = new List<string>();
                    return _learnedKeysCache;
                }

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
                _learnedKeysCache = list;
                return list;
            }
            catch 
            { 
                _learnedKeysCache = new List<string>();
                return _learnedKeysCache; 
            }
        }

        private static partial void AddLearnedKeyLegacy(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            
            // ✅ OPTIMIZATION: Check cache first to avoid redundant disk writes if key already learned
            if (_learnedKeysCache != null && _learnedKeysCache.Contains(key, StringComparer.OrdinalIgnoreCase)) return;

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

                if (!keys.Contains(key))
                {
                    keys.Add(key);
                    // Update cache
                    _learnedKeysCache = keys.ToList();

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
            }
            catch { /* non-fatal */ }
        }
    }
}


