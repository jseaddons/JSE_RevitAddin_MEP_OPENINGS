
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

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
            "System Name", "System Abbreviation", "System Type", "Service Type",
            "Size", "Width", "Height", "Diameter",
            "Level", "Elevation from Level", "Insulation Thickness",
            
            // Host essentials  
            "Thickness",
            "Structural",
            "Fire Rating",
            "b", "B", "Breadth",
            
            // Common - Removed per User Request
            // "Mark", "Comments", "Phase Created",
            
            // Legacy compatibility
            "Nominal Diameter", "Outside Diameter",
            "Reference Level", "Reference Level Elevation", // ✅ Essential for accessories (Damper/Valve) level data
            "Schedule Level", "Schedule of Level",
            // "Type Name", "Family Name" — Accessories only, kept in _commonAccessoryKeys
            
            // ✅ Orientation Angles: Captured for structured DB columns
            "MepAngleToXRad", "MepAngleToXDeg", "MepAngleToYRad", "MepAngleToYDeg",
            "Angle to X", "Angle to Y", "Angle to X-axis", "Angle to Y-axis",
            
            // Structural Framing Type Params
            "b", "B", "Breadth"
        };

        // ✅ MUST-CAPTURE: Always capture these critical parameters (no limits applied)
        // ⚠️ Only add keys that belong to ALL element categories here.
        // Host-only keys ("b") stay in _commonHostKeys. MEP-only keys stay in _ductKeys/_pipeKeys etc.
        // MUST_CAPTURE only bypasses the param-count limit — it does NOT add extra keys outside the whitelist.
        private static readonly string[] MUST_CAPTURE_KEYS = new[]
        {
            "System Type",
            "System Name",
            "System Abbreviation",
            "Reference Level",
            "Reference Level Elevation",
            "Schedule of Level",
            "Schedule Level",
            "Size",
            "Service Type",
            // "b" REMOVED: Host-only (structural framing). It's in _commonHostKeys already.
        };

        // ✅ CATEGORY-SPECIFIC MEP whitelists — zero wasted LookupParameter misses
        private static readonly HashSet<string> _ductKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Ducts: Width+Height for rectangular, Size for all. No Diameter (round duct uses Size).
            // Type Name / Family Name removed: not needed for ducts, only for Accessories
            "Size", "Width", "Height",
            "Level", "Schedule Level", "Schedule of Level",
            "System Type", "System Name", "System Abbreviation", "Insulation Thickness",
            "Offset", "Elevation from Level"
        };

        private static readonly HashSet<string> _pipeKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Pipes: Diameter + Nominal + Outside. No Width/Height (pipes are round).
            // Type Name / Family Name removed: not needed for pipes, only for Accessories
            "Size", "Diameter", "Nominal Diameter", "Outside Diameter",
            "Level", "Schedule Level", "Schedule of Level",
            "System Type", "System Name", "System Abbreviation", "Insulation Thickness",
            "Offset", "Elevation from Level"
        };

        private static readonly HashSet<string> _cableTrayConduitKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Cable Trays & Conduits: Width+Height. Service Type (not System Type). No pipe-only diameter params.
            // Type Name / Family Name removed: not needed for cable trays, only for Accessories
            "Size", "Width", "Height",
            "Level", "Schedule Level", "Schedule of Level",
            "Service Type", "System Abbreviation",
            "Offset", "Elevation from Level"
        };

        // ✅ KEPT for backward-compat with any non-CaptureBatchParams callers
        private static readonly HashSet<string> _commonMepKeys = _ductKeys;

        // ✅ HIGH-SPEED LOOKUP: Map whitelisted strings to BuiltInParameters to avoid slow string-based lookups
        // NOTE: These are safe for strings because ConvertParameterToStringLegacy handles them correctly.
        // "Size" is excluded as it returns formatted strings (e.g. 150x200) that can break numeric logic.
        private static readonly Dictionary<string, BuiltInParameter> _bipMap = new Dictionary<string, BuiltInParameter>(StringComparer.OrdinalIgnoreCase)
        {
            { "Type Name", BuiltInParameter.SYMBOL_NAME_PARAM },
            { "Family Name", BuiltInParameter.ALL_MODEL_FAMILY_NAME },
            { "System Name", BuiltInParameter.RBS_SYSTEM_NAME_PARAM },
            { "System Type", BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM }, 
            { "Piping System Type", BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM },
            { "System Abbreviation", BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM },
            { "Service Type", BuiltInParameter.RBS_CTC_SERVICE_TYPE },
            { "Width", BuiltInParameter.RBS_CURVE_WIDTH_PARAM },
            { "Height", BuiltInParameter.RBS_CURVE_HEIGHT_PARAM },
            { "Diameter", BuiltInParameter.RBS_CURVE_DIAMETER_PARAM },
            { "Nominal Diameter", BuiltInParameter.RBS_PIPE_DIAMETER_PARAM }, 
            { "Outside Diameter", BuiltInParameter.RBS_PIPE_OUTER_DIAMETER },
            { "Level", BuiltInParameter.RBS_START_LEVEL_PARAM }, 
            { "Offset", BuiltInParameter.RBS_OFFSET_PARAM },
            { "Elevation from Level", BuiltInParameter.RBS_OFFSET_PARAM },
            { "Schedule Level", BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM },
            { "Reference Level", BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM },
            { "Insulation Thickness", BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS },
            
            // Host Structural Thickness BIPs
            // "Thickness" maps to Wall Width by default (most common), logic handles fallback for floors
            { "Thickness", BuiltInParameter.WALL_ATTR_WIDTH_PARAM }, 
            { "Wall Width", BuiltInParameter.WALL_ATTR_WIDTH_PARAM },
            { "Floor Thickness", BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM },
            { "Structural", BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM },
            
            // Host Structural Framing BIPs
            { "Fire Rating", BuiltInParameter.FIRE_RATING },
            { "b", BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH } // Case-insensitive dict: covers 'b', 'B', etc.
        };

        private static readonly HashSet<string> _commonHostKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // ✅ Host-only: Walls, Floors, Structural Framing.
            // These are needed for dedicated columns (Thickness, etc.) but will be filtered from JSON.
            "Fire Rating", "b", "B", "Breadth",
            "Wall Width", "Floor Thickness", "Thickness",
            "Structural"
        };

        // ✅ NEW: Accessory specific keys (Loadable families like Dampers)
        private static readonly HashSet<string> _commonAccessoryKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Width", "Height",
            "Damper Width", "Damper Height", // ✅ User requested specific damper param support
            "Diameter",
            "Level", "Schedule Level", "Schedule of Level",
            "System Type", "System Name", "System Abbreviation",
            "Type Name", "Family Name",      // ✅ Identification for Accessories
            "Offset", "Elevation from Level", "Reference Level Elevation" // ✅ Added for accurate elevation capture
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
                    SafeFileLogger.SafeAppendText("parameter_capture.log", $"[PARAM_SNAPSHOT] ⚠️ Large parameter whitelist detected: {keys.Count} parameters. This may increase memory usage significantly.\n");
            }

            return keys;
        }

        /// <summary>
        /// Capture whitelisted parameter values for an element (tries instance, then type if missing).
        /// ✅ FIX 1 & 6: Only capture ESSENTIAL parameters to reduce memory by 90%.
        /// </summary>
        public List<SerializableKeyValue> CaptureParams(Element element, HashSet<string> whitelist)
        {
            return CaptureParamsLegacy(element, whitelist, element?.Document, null, null, null);
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

            // Group by Document for thread-safety and ElementId caching
            var elementsByDoc = elements
                .Where(e => e != null)
                .GroupBy(e => e.Document);

            foreach (var docGroup in elementsByDoc)
            {
                var doc = docGroup.Key;
                if (doc == null) continue;

                var elementCache = new Dictionary<ElementId, string>();
                var typeElementCache = new Dictionary<ElementId, Element>();
                // ✅ TYPE PARAM CACHE: Cache lookup results per type to avoid repetitive string searching
                var typeParamCache = new Dictionary<ElementId, Dictionary<string, Parameter>>();
                // ✅ MISSING PARAM CACHE: Cache keys that are known to be missing for a given type to avoid repetitive string lookups
                var missingTypeKeysCache = new Dictionary<ElementId, HashSet<string>>();

                // Split elements into MEP Curves, Accessories, and Host for optimized whitelisting
                var mepCurveElements = new List<(Element Element, int? CatId, string CatName, string LevelName, double? LevelElevation)>(); // Ducts, Pipes, Trays, Conduits
                var accessoryElements = new List<(Element Element, int? CatId, string CatName, string LevelName, double? LevelElevation)>(); // Dampers, Valves, etc.
                var hostElements = new List<(Element Element, int? CatId, string CatName, string LevelName, double? LevelElevation)>();

                var levelCache = new Dictionary<ElementId, (string Name, double Elevation)>();

                foreach (var element in docGroup)
                {
                    // ✅ USER OPTIMIZATION: Extract Category ONCE in the batch group loop
                    int? catId = element.Category?.Id?.GetIntegerValue();
                    string catName = element.Category?.Name;

                    // ✅ LEVEL CACHE OPTIMIZATION: Extract Level info ONCE per Floor in the batch group loop
                    string levelName = null;
                    double? levelElevation = null;
                    ElementId levelId = element.LevelId;

                    if (levelId != null && levelId != ElementId.InvalidElementId)
                    {
                        if (levelCache.TryGetValue(levelId, out var cachedLevel))
                        {
                            levelName = cachedLevel.Name;
                            levelElevation = cachedLevel.Elevation;
                        }
                        else
                        {
                            var levelInstance = doc.GetElement(levelId) as Level;
                            if (levelInstance != null)
                            {
                                levelName = levelInstance.Name;
                                levelElevation = levelInstance.Elevation;
                                levelCache[levelId] = (levelName, levelElevation.Value);
                            }
                        }
                    }

                    if (element is Duct || element is Pipe || element is CableTray || element is Conduit)
                    {
                        mepCurveElements.Add((element, catId, catName, levelName, levelElevation));
                    }
                    else if (catId == (int)BuiltInCategory.OST_DuctAccessory ||
                             catId == (int)BuiltInCategory.OST_PipeAccessory)
                    {
                        accessoryElements.Add((element, catId, catName, levelName, levelElevation));
                    }
                    else
                    {
                        hostElements.Add((element, catId, catName, levelName, levelElevation));
                    }
                }

                // ✅ CATEGORY-SPLIT MEP Capture: each category only looks up its own parameters
                var swMep = System.Diagnostics.Stopwatch.StartNew();
                if (mepCurveElements.Count > 0)
                {
                    foreach (var tuple in mepCurveElements)
                    {
                        var element = tuple.Element;
                        // Pick the tightest whitelist for this element's category
                        HashSet<string> mepWhitelist;
                        if (element is Duct)
                            mepWhitelist = _ductKeys;
                        else if (element is Pipe)
                            mepWhitelist = _pipeKeys;
                        else // CableTray, Conduit, FamilyInstance MEP
                            mepWhitelist = _cableTrayConduitKeys;

                        var snapshots = CaptureParamsLegacy(element, mepWhitelist, doc, null, elementCache, typeElementCache, typeParamCache, tuple.CatId, tuple.CatName, tuple.LevelName, tuple.LevelElevation, missingTypeKeysCache);
                        var paramDict = results[element.Id.GetIntegerValue()] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var snapshot in snapshots)
                        {
                            if (snapshot != null && !string.IsNullOrEmpty(snapshot.Key))
                                paramDict[snapshot.Key] = snapshot.Value ?? "";
                        }
                    }
                }
                swMep.Stop();

                // ✅ NEW: Process Accessories with Accessory whitelist (Dampers, etc.)
                var swAcc = System.Diagnostics.Stopwatch.StartNew();
                if (accessoryElements.Count > 0)
                {
                    var accessoryWhitelist = new HashSet<string>(_commonAccessoryKeys, StringComparer.OrdinalIgnoreCase);
                    
                    // Add learned keys to accessories too (important for custom damper params)
                    var learnedKeys = GetLearnedKeysLegacy();
                    if (learnedKeys != null)
                        foreach (var k in learnedKeys.Take(20)) accessoryWhitelist.Add(k);

                    foreach (var tuple in accessoryElements)
                    {
                        var element = tuple.Element;
                        var snapshots = CaptureParamsLegacy(element, accessoryWhitelist, doc, null, elementCache, typeElementCache, typeParamCache, tuple.CatId, tuple.CatName, tuple.LevelName, tuple.LevelElevation, missingTypeKeysCache);
                        var paramDict = results[element.Id.GetIntegerValue()] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var snapshot in snapshots)
                        {
                            if (snapshot != null && !string.IsNullOrEmpty(snapshot.Key))
                                paramDict[snapshot.Key] = snapshot.Value ?? "";
                        }
                    }
                }
                swAcc.Stop();

                // Process Host Elements with Host whitelist (Fire Rating only)
                var swHost = System.Diagnostics.Stopwatch.StartNew();
                if (hostElements.Count > 0)
                {
                    var hostWhitelist = new HashSet<string>(_commonHostKeys, StringComparer.OrdinalIgnoreCase);
                    foreach (var tuple in hostElements)
                    {
                        var element = tuple.Element;
                        var snapshots = CaptureParamsLegacy(element, hostWhitelist, doc, null, elementCache, typeElementCache, typeParamCache, tuple.CatId, tuple.CatName, tuple.LevelName, tuple.LevelElevation, missingTypeKeysCache);
                        var paramDict = results[element.Id.GetIntegerValue()] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var snapshot in snapshots)
                        {
                            if (snapshot != null && !string.IsNullOrEmpty(snapshot.Key))
                                paramDict[snapshot.Key] = snapshot.Value ?? "";
                        }
                    }
                }
                swHost.Stop();

                // ✅ ALWAYS-ON: Log group timing breakdown (lightweight — no per-param overhead)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    double mepAvg  = mepCurveElements.Count  > 0 ? (double)swMep.ElapsedMilliseconds  / mepCurveElements.Count  : 0;
                    double accAvg  = accessoryElements.Count > 0 ? (double)swAcc.ElapsedMilliseconds  / accessoryElements.Count  : 0;
                    double hostAvg = hostElements.Count      > 0 ? (double)swHost.ElapsedMilliseconds / hostElements.Count       : 0;
                    SafeFileLogger.SafeAppendText("parameter_perf.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAM-TIMING] " +
                        $"MEP={swMep.ElapsedMilliseconds}ms ({mepCurveElements.Count} elems, {mepAvg:F1}ms/ea, whitelist={_commonMepKeys.Count} keys) | " +
                        $"Acc={swAcc.ElapsedMilliseconds}ms ({accessoryElements.Count} elems, {accAvg:F1}ms/ea) | " +
                        $"Host={swHost.ElapsedMilliseconds}ms ({hostElements.Count} elems, {hostAvg:F1}ms/ea, whitelist={_commonHostKeys.Count} keys)\n");
                }
            }

            return results;
        }

        private static partial List<SerializableKeyValue> CaptureParamsLegacy(Element element, HashSet<string> whitelist, Document doc, string docKey, Dictionary<ElementId, string> elementCache, Dictionary<ElementId, Element> typeElementCache, Dictionary<ElementId, Dictionary<string, Parameter>> typeParamCache = null, int? preCategoryId = null, string preCategoryName = null, string preLevelName = null, double? preLevelElevation = null, Dictionary<ElementId, HashSet<string>> missingTypeKeysCache = null)
        {
            var result = new List<SerializableKeyValue>();

            // ✅ CRASH PROOF: Wrap entire method to prevent single element failure from crashing the process
            try
            {
                // ✅ OPTIMIZATION: Extract category info once, optionally using pre-computed values from batching
                int? categoryId = preCategoryId ?? element?.Category?.Id?.GetIntegerValue();
                string categoryName = preCategoryName ?? element?.Category?.Name;

                if (element == null || whitelist == null || whitelist.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode && element != null)
                    {
                        SafeFileLogger.SafeAppendText("parameter_capture.log", $"[{DateTime.Now}] [PARAM-CAPTURE] ⚠️ Element {element.Id} ({categoryName}): No whitelist provided or whitelist is empty\n");
                    }
                    return result;
                }

                // NOTE: UniqueId and Category are NOT stored in snapshot JSON.
                // Parameter transfer is keyed by SleeveInstanceId; Category is stored directly on ClashZone.

                // ✅ LEVEL CACHE OPTIMIZATION: Output pre-cached level info directly to avoid ClashZoneService re-calculating it
                if (!string.IsNullOrEmpty(preLevelName))
                {
                    result.Add(new SerializableKeyValue { Key = "Reference Level", Value = preLevelName });
                    if (preLevelElevation.HasValue)
                    {
                        // ✅ UNIT: Store in FEET (Revit internal units) - consistent with ElevationFromLevel
                        result.Add(new SerializableKeyValue { Key = "Reference Level Elevation", Value = preLevelElevation.Value.ToString("F6", CultureInfo.InvariantCulture) });
                    }
                }

                // NOTE: UniqueId, Category, Type Name, and Family Name are NOT stored in the parameter list.
                // These are stored in dedicated columns in the ClashZones table.
                // Parameter transfer is handled separately.

                // (Synthetic parameters moved to late processing if needed)
                // ✅ NEW: Synthetic Parameters for Accessories (User Request)
                if (element is FamilyInstance fi && fi.MEPModel?.ConnectorManager != null)
                {
                    try
                    {
                        var connectors = fi.MEPModel.ConnectorManager.Connectors;
                        result.Add(new SerializableKeyValue { Key = "Connector Count", Value = connectors.Size.ToString() });
                        result.Add(new SerializableKeyValue { Key = "Has Connectors", Value = (connectors.Size > 0).ToString() });

                        // ✅ REUSE LEGACY CODE: Use DamperConnectorDetector to get side (User Request)
                        // "Don't scare me with new code" -> We reuse the existing trusted detector
                        var detector = new JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection.DamperConnectorDetector();
                        Connector c;
                        // Use World Coordinates (true) as that is what placement uses
                        string side = detector.DetectConnectorSide(fi, true, out c);

                        result.Add(new SerializableKeyValue { Key = "Connector Side", Value = side ?? "Unknown" });
                    }
                    catch { /* Ignore connector access errors */ }
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

                // Skip redundant diagnostics in tight loop

                // DEBUG: Log all available parameters for duct accessories
                if (categoryId == (int)BuiltInCategory.OST_DuctAccessory && OptimizationFlags.UseDiagnosticMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var msg1 = $"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Starting parameter capture\n";
                        SafeFileLogger.SafeAppendText("parameter_capture.log", msg1);
                    }

                    var allParams = element.Parameters.Cast<Parameter>().Where(p => p != null && !string.IsNullOrEmpty(p.Definition?.Name)).ToList();
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var msg2 = $"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Found {allParams.Count} total parameters\n";
                        SafeFileLogger.SafeAppendText("parameter_capture.log", msg2);
                    }

                    foreach (var param in allParams.Take(10)) // Log first 10 parameters
                    {
                        var paramValue = ConvertParameterToStringLegacy(element, param, elementCache);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var msg3 = $"[{DateTime.Now}] [PARAM_CAPTURE] DUCT ACCESSORY {element.Id}: Parameter '{param.Definition.Name}' = '{paramValue}'\n";
                            SafeFileLogger.SafeAppendText("parameter_capture.log", msg3);
                        }
                    }
                }

                // ✅ CRITICAL: Check if element is a Cable Tray (needed for parameter mapping)
                bool isCableTray = categoryId == (int)BuiltInCategory.OST_CableTray ||
                                  categoryName?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                                  element is CableTray;

                // (Deleted duplicate loop block)

                    // ✅ PERFORMANCE DEBUG: Track time per parameter
                    var swElement = System.Diagnostics.Stopwatch.StartNew();
                    int capturedCount = 0;
                    bool hasFramingWidth = false; // Optimization: If 'b' is found, skip B, Breadth, d, D, Depth
                    bool hasScheduleLevel = false; // Optimization: If Schedule Level found, skip Schedule of Level

                    // ✅ CACHE OPTIMIZATION: Initialize missing keys cache for this element's TypeId
                    ElementId typeIdOfElem = element.GetTypeId();
                    HashSet<string> missingKeysForType = null;
                    if (missingTypeKeysCache != null && typeIdOfElem != null && typeIdOfElem != ElementId.InvalidElementId)
                    {
                        if (!missingTypeKeysCache.TryGetValue(typeIdOfElem, out missingKeysForType))
                        {
                            missingKeysForType = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                            missingTypeKeysCache[typeIdOfElem] = missingKeysForType;
                        }
                    }

                    foreach (var key in orderedKeys)
                    {
                        var pSw = System.Diagnostics.Stopwatch.StartNew();

                        // ✅ MISSING CACHE: Skip if we already proved this key+fallbacks don't exist for this type
                        if (missingKeysForType != null && missingKeysForType.Contains(key))
                        {
                            continue;
                        }

                        // ✅ FIX 6: Emergency brake - stop if limit reached
                        bool isMustCapture = mustCaptureSet.Contains(key);
                        if (!isMustCapture && result.Count >= MAX_PARAMETERS)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("parameter_capture.log", $"[PARAM_SNAPSHOT] Parameter limit ({MAX_PARAMETERS}) reached for element {element.Id}");
                            break;
                        }

                        // ✅ OPTIMIZATION: Skip redundant framing parameters if 'b' was already captured
                        if (hasFramingWidth && (key.Equals("B", StringComparison.OrdinalIgnoreCase) ||
                                               key.Equals("Breadth", StringComparison.OrdinalIgnoreCase) ||
                                               key.Equals("d", StringComparison.OrdinalIgnoreCase) ||
                                               key.Equals("D", StringComparison.OrdinalIgnoreCase) ||
                                               key.Equals("Depth", StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        // ✅ OPTIMIZATION: Level redundancy check
                        if (hasScheduleLevel && (key.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                                                key.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase)))
                        {
                            continue;
                        }

                        // ✅ CRITICAL: For Cable Trays, map "System Type" to "Service Type" in the whitelist
                        string actualKey = key;
                        if (isCableTray && key.Equals("System Type", StringComparison.OrdinalIgnoreCase))
                        {
                            actualKey = "Service Type"; // Use "Service Type" for Cable Trays
                        }

                        // ✅ CRITICAL FIX: If parameter is in whitelist, ALWAYS capture it (regardless of essential/common)
                        bool isEssential = ESSENTIAL_PARAMETERS.Contains(actualKey) ||
                                           (isCableTray && actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase) && ESSENTIAL_PARAMETERS.Contains("Service Type"));

                        // ✅ CRITICAL: Use actualKey (mapped for Cable Trays) instead of original key
                        var p = LookupParamLegacy(element, actualKey, typeElementCache, typeParamCache, typeIdOfElem);

                        // ✅ CRITICAL: Special fallback for System Type/Service Type when not found by name
                        if (p == null && (key.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                          actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase)))
                        {
                            if (isCableTray)
                            {
                                p = element.LookupParameter("Service Type") ?? element.LookupParameter("MEP Service Type");
                            }
                            else
                            {
                                p = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) ??
                                    element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                                    element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);

                                if (p == null)
                                {
                                    p = element.LookupParameter("MEP System Type") ??
                                        element.LookupParameter("System Classification") ??
                                        element.LookupParameter("MEP System Classification");
                                }

                                // ✅ DUCT/PIPE ACCESSORY: traverse connectors to get system type from connected MEPSystem
                                if (p == null && (categoryId == (int)BuiltInCategory.OST_DuctAccessory ||
                                                  categoryId == (int)BuiltInCategory.OST_PipeAccessory) &&
                                    element is FamilyInstance fiST)
                                {
                                    try
                                    {
                                        var connMgrST = fiST.MEPModel?.ConnectorManager;
                                        if (connMgrST != null)
                                        {
                                            foreach (Connector conn in connMgrST.Connectors)
                                            {
                                                var mepSys = conn.MEPSystem;
                                                if (mepSys == null) continue;
                                                var sysPar = mepSys.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
                                                          ?? mepSys.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                                                          ?? mepSys.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM)
                                                          ?? mepSys.LookupParameter("System Type");
                                                if (sysPar != null)
                                                {
                                                    var sysTypeStr = sysPar.AsValueString() ?? sysPar.AsString();
                                                    if (!string.IsNullOrEmpty(sysTypeStr))
                                                    {
                                                        result.Add(new SerializableKeyValue { Key = actualKey, Value = sysTypeStr });
                                                        goto nextKey_SystemType;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                    nextKey_SystemType:;
                                }
                            }
                        }

                        // ✅ CRITICAL: Special fallback for Service Type (Cable Trays)
                        if (p == null && key.Equals("Service Type", StringComparison.OrdinalIgnoreCase))
                        {
                            p = element.LookupParameter("Service Type") ?? element.LookupParameter("MEP Service Type");
                        }

                        // ✅ CRITICAL: Special fallback for System Name
                        if (p == null && key.Equals("System Name", StringComparison.OrdinalIgnoreCase))
                        {
                            p = element.LookupParameter("MEP System Name") ??
                                element.LookupParameter("System Name") ??
                                element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);

                            // ✅ DUCT/PIPE ACCESSORY: traverse connectors to get system name from connected MEPSystem
                            if (p == null && (categoryId == (int)BuiltInCategory.OST_DuctAccessory ||
                                              categoryId == (int)BuiltInCategory.OST_PipeAccessory) &&
                                element is FamilyInstance fiSN)
                            {
                                try
                                {
                                    var connMgrSN = fiSN.MEPModel?.ConnectorManager;
                                    if (connMgrSN != null)
                                    {
                                        foreach (Connector conn in connMgrSN.Connectors)
                                        {
                                            var mepSys = conn.MEPSystem;
                                            if (mepSys == null) continue;
                                            var sysName = mepSys.Name;
                                            if (!string.IsNullOrEmpty(sysName))
                                            {
                                                result.Add(new SerializableKeyValue { Key = key, Value = sysName });
                                                goto nextKey_SystemName;
                                            }
                                        }
                                    }
                                }
                                catch { }
                                nextKey_SystemName:;
                            }
                        }

                        // ✅ CRITICAL: Special fallback for System Abbreviation
                        if (p == null && key.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase))
                        {
                            p = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
                             ?? element.LookupParameter("System Abbreviation") ??
                                element.LookupParameter("System Abbr") ??
                                element.LookupParameter("Abbreviation") ??
                                element.LookupParameter("Abbr");

                            // ✅ DUCT/PIPE ACCESSORY: traverse connectors to get system abbreviation from connected MEPSystem
                            if (p == null && (categoryId == (int)BuiltInCategory.OST_DuctAccessory ||
                                              categoryId == (int)BuiltInCategory.OST_PipeAccessory) &&
                                element is FamilyInstance fiSA)
                            {
                                try
                                {
                                    var connMgrSA = fiSA.MEPModel?.ConnectorManager;
                                    if (connMgrSA != null)
                                    {
                                        foreach (Connector conn in connMgrSA.Connectors)
                                        {
                                            var mepSys = conn.MEPSystem;
                                            if (mepSys == null) continue;
                                            var abbrPar = mepSys.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM)
                                                       ?? mepSys.LookupParameter("System Abbreviation")
                                                       ?? mepSys.LookupParameter("System Abbr");
                                            if (abbrPar != null)
                                            {
                                                var abbrStr = abbrPar.AsString();
                                                if (!string.IsNullOrEmpty(abbrStr))
                                                {
                                                    result.Add(new SerializableKeyValue { Key = key, Value = abbrStr });
                                                    goto nextKey_SystemAbbr;
                                                }
                                            }
                                        }
                                    }
                                }
                                catch { }
                                nextKey_SystemAbbr:;
                            }
                        }

                        // ✅ CRITICAL FIX: "Elevation from Level" mapping per User Request
                        // For Ducts, Pipes, Cable Trays -> Use RBS_OFFSET_PARAM (Middle Elevation)
                        // For Dampers -> Use "Elevation from Level" (keep existing lookup)
                        if (key.Equals("Elevation from Level", StringComparison.OrdinalIgnoreCase))
                        {
                            var catId = categoryId;
                            // RBS_OFFSET_PARAM is valid for Curves and Cable Trays
                            if (catId == (int)BuiltInCategory.OST_DuctCurves ||
                                catId == (int)BuiltInCategory.OST_PipeCurves ||
                                catId == (int)BuiltInCategory.OST_CableTray)
                            {
                                p = element.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
                            }

                            // fallback for everyone (including dampers and failed curve lookups)
                            if (p == null)
                            {
                                p = element.LookupParameter("Elevation from Level") ??
                                    element.LookupParameter("Middle Elevation") ??
                                    element.LookupParameter("Offset");
                            }
                        }

                        // ✅ CRITICAL FIX: Damper Accessory Fallback (User Request)
                        // If "Width" is requested for a Damper, try "Damper Width"
                        if (p == null &&
                           (categoryId == (int)BuiltInCategory.OST_DuctAccessory ||
                            categoryId == (int)BuiltInCategory.OST_PipeAccessory))
                        {
                            if (key.Equals("Width", StringComparison.OrdinalIgnoreCase))
                            {
                                p = element.LookupParameter("Damper Width") ??
                                    element.LookupParameter("Valve Width");
                            }
                            else if (key.Equals("Height", StringComparison.OrdinalIgnoreCase))
                            {
                                p = element.LookupParameter("Damper Height") ??
                                    element.LookupParameter("Valve Height");
                            }
                        }

                        // ✅ CRITICAL FIX: Wall Width is a TYPE parameter (User Request)
                        // If "Width" is requested for a Wall, we must look at the Type Parameter WALL_ATTR_WIDTH_PARAM
                        if (p == null && key.Equals("Width", StringComparison.OrdinalIgnoreCase) &&
                            categoryId == (int)BuiltInCategory.OST_Walls)
                        {
                            var innerTypeId = typeIdOfElem;
                            if (innerTypeId != null && innerTypeId != ElementId.InvalidElementId)
                            {
                                // Try to get type from cache first
                                Element typeElem = null;
                                if (typeElementCache != null && typeElementCache.TryGetValue(innerTypeId, out var cachedType))
                                {
                                    typeElem = cachedType;
                                }
                                else
                                {
                                    typeElem = element.Document.GetElement(innerTypeId);
                                    if (typeElementCache != null && typeElem != null)
                                    {
                                        typeElementCache[innerTypeId] = typeElem;
                                    }
                                }

                                if (typeElem != null)
                                {
                                    p = typeElem.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                                }
                            }
                        }

                        // ✅ CRITICAL FIX: Special fallback for "Schedule of Level"
                        if (p == null && (key.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                                          key.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase)))
                        {
                            p = element.LookupParameter("Schedule of Level") ??
                                element.LookupParameter("Schedule Level") ??
                                element.LookupParameter("Elevation from Level");
                        }

                        // ✅ CRITICAL FIX: Special fallback for "Reference Level"
                        if (p == null && (key.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) ||
                                          key.Equals("Level", StringComparison.OrdinalIgnoreCase)))
                        {
                            p = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM) ??
                                element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM) ??
                                element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);

                            if (p == null)
                            {
                                p = element.LookupParameter("Reference Level") ??
                                    element.LookupParameter("Level") ??
                                    element.LookupParameter("Schedule Level") ??
                                    element.LookupParameter("Schedule of Level");
                            }
                        }

                        if (p == null)
                        {
                            // ✅ MISSING CACHE: Remember that this key and all its fallbacks completely failed for this type
                            if (missingKeysForType != null)
                            {
                                missingKeysForType.Add(key);
                            }
                            continue;
                        }

                        var value = ConvertParameterToStringLegacy(element, p, elementCache);

                        // Handle empty values for essential parameters
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            if (isEssential && (actualKey.Equals("System Type", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Service Type", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("System Name", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("System Abbreviation", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Level", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Reference Level", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Size", StringComparison.OrdinalIgnoreCase) ||
                                              actualKey.Equals("Fire Rating", StringComparison.OrdinalIgnoreCase)))
                            {
                                value = ""; // Capture empty string for essential 
                            }
                            else
                            {
                                continue; // Skip empty non-essential
                            }
                        }

                        // Truncate
                        const int MAX_PARAM_VALUE_LENGTH = 200;
                        if (value != null && value.Length > MAX_PARAM_VALUE_LENGTH)
                        {
                            value = value.Substring(0, MAX_PARAM_VALUE_LENGTH);
                        }

                        result.Add(new SerializableKeyValue
                        {
                            Key = actualKey,
                            Value = value ?? ""
                        });
                        capturedCount++;

                        // ✅ OPTIMIZATION: Track successful 'b' capture
                        if (!hasFramingWidth && actualKey.Equals("b", StringComparison.OrdinalIgnoreCase))
                        {
                            hasFramingWidth = true;
                        }

                        // ✅ OPTIMIZATION: Track successful Schedule Level capture
                        if (!hasScheduleLevel && (actualKey.Equals("Schedule Level", StringComparison.OrdinalIgnoreCase) ||
                                                 actualKey.Equals("Schedule of Level", StringComparison.OrdinalIgnoreCase)))
                        {
                            hasScheduleLevel = true;
                        }

                        // ✅ PERFORMANCE LOGGING
                        pSw.Stop();
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode && pSw.ElapsedMilliseconds > 5)
                        {
                            SafeFileLogger.SafeAppendText("parameter_perf.log", $"[{DateTime.Now:HH:mm:ss.fff}] SLOW PARAM: {key} ({actualKey}) on Ems {element.Id} took {pSw.ElapsedMilliseconds}ms\n");
                        }
                    }

                    // ✅ ELEMENT-LEVEL PERFORMANCE LOGGING
                    swElement.Stop();
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode && swElement.ElapsedMilliseconds > 20)
                    {
                        SafeFileLogger.SafeAppendText("parameter_perf.log", $"[{DateTime.Now:HH:mm:ss.fff}] SLOW ELEMENT: {element.Id} ({categoryName}) captured {capturedCount} params in {swElement.ElapsedMilliseconds}ms\n");
                    }

                    // ✅ ENHANCED LOGGING: Summary of captured parameters
                    if (!DeploymentConfiguration.DeploymentMode && result.Count > 0 && OptimizationFlags.UseDiagnosticMode)
                    {
                        var capturedKeys = string.Join(", ", result.Select(kv => kv.Key).Take(10));
                        var moreCount = result.Count > 10 ? $" (+{result.Count - 10} more)" : "";
                        var msg = $"[{DateTime.Now}] [PARAM-CAPTURE] ✅ SUMMARY: Element {element.Id} ({categoryName}): Captured {result.Count} parameters: {capturedKeys}{moreCount}\n";
                        SafeFileLogger.SafeAppendText("parameter_capture.log", msg);
                    }
                    else if (!DeploymentConfiguration.DeploymentMode && result.Count == 0 && orderedKeys.Count > 0 && OptimizationFlags.UseDiagnosticMode)
                    {
                        SafeFileLogger.SafeAppendText("parameter_capture.log", $"[{DateTime.Now}] [PARAM-CAPTURE] ⚠️ WARNING: Element {element.Id} ({categoryName}): No parameters captured from {orderedKeys.Count} whitelisted parameters!\n");
                    }

                    return result;


                }
            catch (Exception ex)
            {
                // ✅ CRASH PROOF: Log error but return empty result to prevent process crash
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var msg = $"[{DateTime.Now}] [PARAM-CAPTURE] ❌ ERROR capturing element {element?.Id}: {ex.Message}\n";
                    SafeFileLogger.SafeAppendText("parameter_capture.log", msg);
                }
                return result;
            }
        }
        
        /// <summary>
        /// Convert a parameter value to a robust invariant string.
        /// ✅ ENHANCED: Tries multiple methods to extract parameter value.
        /// ✅ OPTIMIZED: Uses cache to prevent redundant Element lookups (round-trips).
        /// </summary>
        private static partial string ConvertParameterToStringLegacy(Element element, Parameter param, Dictionary<ElementId, string> elementCache)
        {
            if (param == null) return string.Empty;

            // ✅ PRIORITY 1: Check StorageType first to return raw data for Doubles/Integers (Unit Consistency)
            // This prevents "1200 mm" strings which cause parsing errors in consumers expecting feet
            if (param.StorageType == StorageType.Double)
            {
                try
                {
                    if (param.HasValue)
                    {
                        // ✅ FIX 1: Return raw internal units (Feet) with high precision (nop rounding per plan, or 5+ dec)
                        // Use 5 decimals to be safe but avoid noise
                        return param.AsDouble().ToString("F5", CultureInfo.InvariantCulture);
                    }
                }
                catch { }
                return string.Empty;
            }
            else if (param.StorageType == StorageType.Integer)
            {
                try
                {
                    if (param.HasValue)
                        return param.AsInteger().ToString(CultureInfo.InvariantCulture);
                }
                catch { }
                return string.Empty;
            }

            // ✅ PRIORITY 2: Try AsString() (text parameters)
            string value = param.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            // ✅ PRIORITY 3: Try AsValueString() (formatted display value - last resort for non-numeric)
            value = param.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            // Handle ElementId
            if (param.StorageType == StorageType.ElementId)
            {
                try
                {
                    ElementId id = param.AsElementId();
                    if (id != null && id != ElementId.InvalidElementId)
                    {
                        // ✅ CACHE OPTIMIZATION: Check cache first to avoid Revit API round-trip
                        if (elementCache != null && elementCache.TryGetValue(id, out var cachedName))
                        {
                            return cachedName;
                        }

                        // Prefer referenced element name for readability if available
                        var e = element?.Document?.GetElement(id);
                        if (e != null)
                        {
                            var name = e.Name;
                            if (!string.IsNullOrWhiteSpace(name)) 
                            {
                                if (elementCache != null) elementCache[id] = name; // seamless update
                                return name;
                            }
                            
                            // Try Category name if element name is empty
                            var catName = e.Category?.Name;
                            if (!string.IsNullOrWhiteSpace(catName)) 
                            {
                                if (elementCache != null) elementCache[id] = catName;
                                return catName;
                            }
                        }
                        
                        // Fallback to ElementId integer value
                        var idStr = id.GetIntegerValue().ToString(CultureInfo.InvariantCulture);
                        if (elementCache != null) elementCache[id] = idStr; // cache the numeric string too
                        return idStr;
                    }
                }
                catch { }
                return string.Empty;
            }
            else if (param.StorageType == StorageType.String)
            {
                // Already tried AsString() and AsValueString() above
                return string.Empty;
            }

            return string.Empty;
        }

        private static partial Parameter LookupParamLegacy(Element element, string paramName, Dictionary<ElementId, Element> typeElementCache, Dictionary<ElementId, Dictionary<string, Parameter>> typeParamCache = null, ElementId optionalTypeId = null)
        {
            // ✅ HIGH-SPEED PATH: Try BuiltInParameter mapping first (instant, no string search)
            if (_bipMap.TryGetValue(paramName, out var bip))
            {
                var p = element.get_Parameter(bip);
                if (p != null) return p;

                // ✅ BIP FALLBACKS: Handle cross-category naming differences without slower Name-based lookup
                if (paramName.Equals("Level", StringComparison.OrdinalIgnoreCase))
                {
                    p = element.get_Parameter(BuiltInParameter.LEVEL_PARAM) ??
                        element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM) ??
                        element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    if (p != null) return p;
                }
                else if (paramName.Equals("System Type", StringComparison.OrdinalIgnoreCase))
                {
                    p = element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                        element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                    if (p != null) return p;
                }
                else if (paramName.Equals("Width", StringComparison.OrdinalIgnoreCase))
                {
                    p = element.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM) ??
                        element.get_Parameter(BuiltInParameter.GENERIC_WIDTH); // For accessories
                    if (p != null) return p;
                }
                else if (paramName.Equals("Height", StringComparison.OrdinalIgnoreCase))
                {
                    p = element.get_Parameter(BuiltInParameter.GENERIC_HEIGHT);
                    if (p != null) return p;
                }
            }

            // Fallback 1: Standard string-based lookup (slower)
            var sp = element.LookupParameter(paramName);
            if (sp != null) return sp;

            // Fallback 2: Check type parameters
            ElementId typeId = optionalTypeId ?? element.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                // ✅ CACHE OPTIMIZATION: Check cache first to avoid redundant GetElement(typeId) calls
                Element typeElem = null;
                if (typeElementCache != null && typeElementCache.TryGetValue(typeId, out var cachedTypeElem))
                {
                    typeElem = cachedTypeElem;
                }
                else
                {
                    typeElem = element.Document.GetElement(typeId);
                    if (typeElem != null && typeElementCache != null)
                    {
                        typeElementCache[typeId] = typeElem; 
                    }
                }
                
                if (typeElem != null)
                {
                    // Try BIP on type first
                    if (_bipMap.TryGetValue(paramName, out var typeBip))
                    {
                        var tp = typeElem.get_Parameter(typeBip);
                        if (tp != null) return tp;
                    }
                    
                    // ✅ TYPE PARAM CATCHING: Check if we've already looked up this parameter on this type
                    if (typeParamCache != null)
                    {
                        if (!typeParamCache.ContainsKey(typeId))
                        {
                            typeParamCache[typeId] = new Dictionary<string, Parameter>(StringComparer.OrdinalIgnoreCase);
                        }
                        
                        if (typeParamCache[typeId].TryGetValue(paramName, out var cachedParam))
                        {
                            return cachedParam;
                        }
                    }

                    var tsp = typeElem.LookupParameter(paramName);
                    
                    // Cache the result (even if null, to avoid re-lookup? No, LookUpParameter might be flaky, store found only for now)
                    if (tsp != null && typeParamCache != null)
                    {
                        typeParamCache[typeId][paramName] = tsp;
                    }

                    if (tsp != null) return tsp;
                }
            }
            
            // ✅ SPECIAL HANDLING: For "Size" parameter, try multiple fallbacks
            // Dampers may have Size as a calculated/formula parameter that needs different lookup
            if (paramName.Equals("Size", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ DEBUG: Log all parameter names for Duct Accessories to find Size
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

            return GetLearnedKeysLegacy().AsEnumerable();
        }

        private static List<string> GetLearnedKeysLegacy()
        {
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

