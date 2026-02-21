using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Responsibility: Define WHAT parameters to capture for full snapshots.
    /// 
    /// Features Preserved (28-point compliance):
    /// - FEATURE 22: Essential parameters whitelist (memory optimization)
    /// - FEATURE 23: Must-capture keys (System Type, Name, Abbr, Levels)
    /// - FEATURE 24: User-defined parameter support
    /// - FEATURE 25: Learned keys integration
    /// - FEATURE 26: Cable Tray Service Type mapping
    /// - FEATURE 27: Parameter limit enforcement
    /// </summary>
    public class SnapshotParameterPolicy : IParameterPolicy
    {
        private readonly IParameterKeyStore _keyStore;
        private readonly ClashZoneStorage _storage;
        
        // FEATURE 22: Essential parameters - only capture these to reduce memory by 90%
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
            "Thickness",
            "Structural",
            "Wall Width", "Floor Thickness", // ✅ Added per BIP Map
            
            // Common
            // "Mark", "Comments", "Phase Created", // Removed
            
            // Legacy compatibility
            "Nominal Diameter", "Outside Diameter",
            "Schedule Level", "Schedule of Level", // Removed Reference Level
            "System Classification", "Service Type",
            "Fire Rating",
            "b", "B", "Breadth" // Sync framing params
        };

        // FEATURE 23: MUST-CAPTURE - Always capture these critical parameters (no limits applied)
        private static readonly HashSet<string> MUST_CAPTURE_KEYS = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "System Type",
            "System Name",
            "System Abbreviation",
            // "Reference Level", // Removed per user request
            "Schedule of Level",
            "Schedule Level",
            "Size",           // ✅ CRITICAL
            "Service Type",   // ✅ CRITICAL
            "b"               // ✅ CRITICAL: Structural Framing Width
        };

        public SnapshotParameterPolicy(IParameterKeyStore keyStore, ClashZoneStorage storage = null)
        {
            _keyStore = keyStore ?? throw new ArgumentNullException(nameof(keyStore));
            _storage = storage;
        }

        /// <summary>
        /// Build whitelist by combining essential params + user-defined + learned keys.
        /// FEATURE 24: User-defined parameter support (limited to 50).
        /// FEATURE 25: Learned keys integration (limited to 20).
        /// </summary>
        public HashSet<string> GetWhitelist()
        {
            var keys = new HashSet<string>(ESSENTIAL_PARAMETERS, StringComparer.OrdinalIgnoreCase);

            // FEATURE 24: Add user-defined parameters (limited to 50)
            if (_storage?.ParameterKeyWhitelist != null)
            {
                var limitedWhitelist = _storage.ParameterKeyWhitelist.Take(50);
                foreach (var k in limitedWhitelist) keys.Add(k);
            }

            // FEATURE 25: Add learned parameters (limited to 20)
            if (_storage?.LearnedParameterKeys != null)
            {
                var limitedLearned = _storage.LearnedParameterKeys.Take(20);
                foreach (var k in limitedLearned) keys.Add(k);
            }

            // FEATURE 25: Merge disk-learned keys (also limited to 20)
            var diskKeys = _keyStore.LoadLearnedKeys();
            foreach (var k in diskKeys.Take(20)) keys.Add(k);

            // Warn if whitelist is too large
            if (keys.Count > 30 && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[PARAM_SNAPSHOT] ⚠️ Large parameter whitelist detected: {keys.Count} parameters. This may increase memory usage.");
            }

            return keys;
        }

        /// <summary>
        /// FEATURE 23: Must-capture keys that bypass parameter limits.
        /// </summary>
        public HashSet<string> GetMustCaptureKeys()
        {
            return new HashSet<string>(MUST_CAPTURE_KEYS, StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Determine if parameter should be captured based on policy.
        /// </summary>
        public bool ShouldCapture(string parameterName, Element element)
        {
            // Always capture must-capture keys
            if (MUST_CAPTURE_KEYS.Contains(parameterName))
                return true;

            // Capture if in whitelist
            return GetWhitelist().Contains(parameterName);
        }

        /// <summary>
        /// FEATURE 26: Map parameter names for category-specific handling.
        /// Example: "System Type" → "Service Type" for Cable Trays.
        /// </summary>
        public string MapParameterName(string requestedName, Element element)
        {
            // FEATURE 26: Cable Tray mapping
            bool isCableTray = element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_CableTray ||
                              element.Category?.Name?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                              element is CableTray;

            if (isCableTray && requestedName.Equals("System Type", StringComparison.OrdinalIgnoreCase))
            {
                return "Service Type";
            }

            return requestedName;
        }

        /// <summary>
        /// FEATURE 27: Emergency parameter limit to prevent memory bloat.
        /// </summary>
        public int GetMaxParameterLimit()
        {
            return 30; // Emergency brake
        }
    }
}