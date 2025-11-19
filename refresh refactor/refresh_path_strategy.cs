using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using SleevePlacementPath = JSE_RevitAddin_MEP_OPENINGS.Services.SleevePlacementPath;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Strategy interface for handling different refresh paths.
    /// Based on PERSISTENCE_TIMING_FLOW_DIAGRAM.md:
    /// - PATH 1: Replay mode (Adopt OFF, filter exists)
    /// - PATH 2: Fresh mode (first time adding filter)
    /// - PATH 3: Non-fresh mode (filter exists, refresh again)
    /// </summary>
    public interface IRefreshPathStrategy
    {
        /// <summary>
        /// Determines if structural updates should be enabled
        /// </summary>
        bool AllowStructuralUpdates { get; }
        
        /// <summary>
        /// Determines if 3-point validation should be enabled
        /// </summary>
        bool EnableThreePointValidation { get; }
        
        /// <summary>
        /// Determines if flag reset should be performed
        /// </summary>
        bool ShouldResetFlags { get; }
        
        /// <summary>
        /// Determines if flag sync should be performed
        /// </summary>
        bool ShouldSyncFlags { get; }
        
        /// <summary>
        /// Determines if GUID checking should be performed
        /// </summary>
        bool ShouldCheckGuids { get; }
        
        /// <summary>
        /// Gets the path name for logging
        /// </summary>
        string PathName { get; }
        
        /// <summary>
        /// Processes zones after validation (zone splitting for PATH 3)
        /// </summary>
        List<ClashZone> ProcessZonesAfterValidation(
            RefreshContext context,
            List<ClashZone> validatedZones,
            List<ClashZone> nonValidatedZones);
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves (DB first, then XML)
        /// </summary>
        void ResetInstanceIdsForDeletedSleeves(
            RefreshContext context,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory);
    }
    
    /// <summary>
    /// PATH 1: Replay Mode
    /// - "Adopt to Document" UNCHECKED
    /// - Filter already processed
    /// - Simple: Only flags/IDs, no structural updates
    /// </summary>
    public class Path1Strategy : IRefreshPathStrategy
    {
        public bool AllowStructuralUpdates => false;
        public bool EnableThreePointValidation => false;
        public bool ShouldResetFlags => true;
        public bool ShouldSyncFlags => false;
        public bool ShouldCheckGuids => false;
        public string PathName => "PATH 1 (Replay Mode)";
        
        public List<ClashZone> ProcessZonesAfterValidation(
            RefreshContext context,
            List<ClashZone> validatedZones,
            List<ClashZone> nonValidatedZones)
        {
            // PATH 1: Keep all zones (no validation, no splitting)
            return validatedZones ?? new List<ClashZone>();
        }
        
        public void ResetInstanceIdsForDeletedSleeves(
            RefreshContext context,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory)
        {
            // PATH 1: Reset instance IDs (DB first, then XML)
            if (categories == null || categories.Count == 0)
                return;
            
            var flagManager = new FlagManager(context.Document);
            flagManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, context.RefreshLogName);
        }
    }
    
    /// <summary>
    /// PATH 2: Fresh Mode
    /// - First time adding filter
    /// - "Adopt to Document" CHECKED
    /// - Simple: Detect → dump → place → done
    /// </summary>
    public class Path2Strategy : IRefreshPathStrategy
    {
        public bool AllowStructuralUpdates => true;
        public bool EnableThreePointValidation => false;
        public bool ShouldResetFlags => false;
        public bool ShouldSyncFlags => false;
        public bool ShouldCheckGuids => false;
        public string PathName => "PATH 2 (Fresh Mode)";
        
        public List<ClashZone> ProcessZonesAfterValidation(
            RefreshContext context,
            List<ClashZone> validatedZones,
            List<ClashZone> nonValidatedZones)
        {
            // PATH 2: No validation, return all zones as-is
            return validatedZones ?? new List<ClashZone>();
        }
        
        public void ResetInstanceIdsForDeletedSleeves(
            RefreshContext context,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory)
        {
            // PATH 2: No instance ID reset needed (fresh mode, no existing sleeves)
        }
    }
    
    /// <summary>
    /// PATH 3: Non-Fresh Mode
    /// - Filter exists, refresh again
    /// - "Adopt to Document" CHECKED
    /// - Complex: Zone splitting after validation
    /// </summary>
    public class Path3Strategy : IRefreshPathStrategy
    {
        public bool AllowStructuralUpdates => true;
        public bool EnableThreePointValidation => true;
        public bool ShouldResetFlags => true;
        public bool ShouldSyncFlags => true;
        public bool ShouldCheckGuids => true;
        public string PathName => "PATH 3 (Non-Fresh Mode)";
        
        public List<ClashZone> ProcessZonesAfterValidation(
            RefreshContext context,
            List<ClashZone> validatedZones,
            List<ClashZone> nonValidatedZones)
        {
            // PATH 3: Zone splitting
            // - Validated zones → PATH 1 behavior (simple, no timing issues)
            // - Non-validated zones → PATH 3 full logic (complex timing issues)
            
            var result = new List<ClashZone>();
            
            // Add validated zones (will use PATH 1 behavior - no structural updates)
            if (validatedZones != null && validatedZones.Count > 0)
            {
                result.AddRange(validatedZones);
                
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[PATH-3] Zone splitting: {validatedZones.Count} validated zones → PATH 1 behavior (simple)");
                }
            }
            
            // Add non-validated zones (will use PATH 3 full logic - structural updates)
            if (nonValidatedZones != null && nonValidatedZones.Count > 0)
            {
                result.AddRange(nonValidatedZones);
                
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info($"[PATH-3] Zone splitting: {nonValidatedZones.Count} non-validated zones → PATH 3 full logic (complex)");
                }
            }
            
            return result;
        }
        
        public void ResetInstanceIdsForDeletedSleeves(
            RefreshContext context,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory)
        {
            // PATH 3: Reset instance IDs (DB first, then XML)
            if (categories == null || categories.Count == 0)
                return;
            
            var flagManager = new FlagManager(context.Document);
            flagManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, context.RefreshLogName);
        }
    }
    
    /// <summary>
    /// Determines which refresh path strategy to use based on context
    /// </summary>
    public static class RefreshPathDeterminer
    {
        /// <summary>
        /// Determines the refresh path strategy based on:
        /// 1. enableThreePointValidation (Adopt to Document checkbox)
        /// 2. FileCombo existence (checked FIRST before determining path)
        /// 3. Whether filter exists (has existing clash zones)
        /// </summary>
        public static IRefreshPathStrategy DeterminePath(
            RefreshContext context,
            bool enableThreePointValidation)
        {
            // ✅ STEP 1: Check FileCombo existence and update IsFilterComboNew flag BEFORE determining path
            // This ensures the flag is set correctly based on actual FileCombo data in database
            UpdateIsFilterComboNewFlagBasedOnFileCombos(context);
            
            // PATH 1: "Adopt to Document" UNCHECKED
            if (!enableThreePointValidation)
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[PATH-DETERMINER] PATH 1 selected: Adopt to Document UNCHECKED");
                }
                return new Path1Strategy();
            }
            
            // PATH 2 or PATH 3: "Adopt to Document" CHECKED
            // Check if filter exists (has existing clash zones)
            bool filterExists = HasExistingClashZones(context);
            
            if (!filterExists)
            {
                // PATH 2: First time adding filter (fresh)
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[PATH-DETERMINER] PATH 2 selected: Fresh mode (first time adding filter)");
                }
                return new Path2Strategy();
            }
            else
            {
                // PATH 3: Filter exists, refresh again (non-fresh)
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Info("[PATH-DETERMINER] PATH 3 selected: Non-fresh mode (filter exists, refresh again)");
                }
                return new Path3Strategy();
            }
        }
        
        /// <summary>
        /// ✅ FILE COMBO CHECK: Updates IsFilterComboNew flag based on FileCombo existence in database
        /// Checks if FileCombos exist for current UI state (filter + linked files + host files + category)
        /// Sets IsFilterComboNew = true if FileCombos don't exist (fresh combo)
        /// Sets IsFilterComboNew = false if FileCombos exist (used combo)
        /// </summary>
        private static void UpdateIsFilterComboNewFlagBasedOnFileCombos(RefreshContext context)
        {
            try
            {
                if (context == null || context.Document == null)
                    return;
                
                if (context.SelectedFilterNames == null || context.SelectedFilterNames.Count == 0)
                    return;
                
                if (context.SelectedReferenceFiles == null || context.SelectedReferenceFiles.Count == 0)
                    return;
                
                if (context.SelectedHostFiles == null || context.SelectedHostFiles.Count == 0)
                    return;
                
                if (context.SelectedMepCategories == null || context.SelectedMepCategories.Count == 0)
                    return;
                
                using (var dbContext = new Data.SleeveDbContext(context.Document))
                {
                    var filterRepository = new Data.Repositories.FilterRepository(dbContext, msg =>
                    {
                        if (!context.IsDeploymentMode)
                            DebugLogger.Info($"[FILE-COMBO-CHECK] {msg}");
                    });
                    
                    // Build all file combo combinations from UI state
                    var allFileComboKeys = new List<(string LinkedFileKey, string HostFileKey)>();
                    foreach (var refFile in context.SelectedReferenceFiles)
                    {
                        foreach (var hostFile in context.SelectedHostFiles)
                        {
                            var linkedKey = NormalizeDocumentKey(refFile);
                            var hostKey = NormalizeDocumentKey(hostFile);
                            allFileComboKeys.Add((linkedKey, hostKey));
                        }
                    }
                    
                    // Check each filter+category combo
                    foreach (var filterName in context.SelectedFilterNames)
                    {
                        foreach (var category in context.SelectedMepCategories)
                        {
                            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                                continue;
                            
                            // Get filter ID
                            int filterId = filterRepository.GetFilterId(filterName, category);
                            if (filterId <= 0)
                            {
                                // Filter doesn't exist - ensure it's created with IsFilterComboNew = true
                                filterId = filterRepository.EnsureFilter(filterName, category);
                                if (filterId > 0 && !context.IsDeploymentMode)
                                {
                                    DebugLogger.Info($"[FILE-COMBO-CHECK] ✅ Created new filter '{filterName}' (Category='{category}') with IsFilterComboNew=true");
                                }
                                continue;
                            }
                            
                            // ✅ FIX: Check and update IsFilterComboNew flag PER FILE COMBO (not per filter+category)
                            // Each file combo (linked file + host file) has its own flag
                            using (var cmd = dbContext.Connection.CreateCommand())
                            {
                                foreach (var (linkedKey, hostKey) in allFileComboKeys)
                                {
                                    // Check if this specific file combo exists
                                    cmd.CommandText = @"
                                        SELECT ComboId, IsFilterComboNew FROM FileCombos 
                                        WHERE FilterId = @FilterId AND LinkedFileKey = @LinkedFileKey AND HostFileKey = @HostFileKey";
                                    cmd.Parameters.Clear();
                                    cmd.Parameters.AddWithValue("@FilterId", filterId);
                                    cmd.Parameters.AddWithValue("@LinkedFileKey", linkedKey);
                                    cmd.Parameters.AddWithValue("@HostFileKey", hostKey);
                                    
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            // File combo exists - check if flag needs updating
                                            int comboId = reader.GetInt32(0);
                                            int isNew = reader.GetInt32(1);
                                            
                                            if (isNew == 1)
                                            {
                                                // File combo exists but flag is still true - should be false (already used)
                                                UpdateFileComboFlag(dbContext, comboId, false);
                                                if (!context.IsDeploymentMode)
                                                {
                                                    DebugLogger.Info($"[FILE-COMBO-CHECK] ✅ FileCombo exists (ComboId={comboId}, Linked='{linkedKey}', Host='{hostKey}') → IsFilterComboNew=false (used combo)");
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // File combo doesn't exist - will be created with IsFilterComboNew=1 when clash zones are saved
                                            if (!context.IsDeploymentMode)
                                            {
                                                DebugLogger.Info($"[FILE-COMBO-CHECK] ⚠️ FileCombo missing (Linked='{linkedKey}', Host='{hostKey}') → will be created with IsFilterComboNew=true (fresh combo)");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!context.IsDeploymentMode)
                {
                    DebugLogger.Warning($"[FILE-COMBO-CHECK] ❌ Error checking FileCombo existence: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// ✅ FIX: Updates IsFilterComboNew flag in FileCombos table for a specific file combo
        /// </summary>
        private static void UpdateFileComboFlag(Data.SleeveDbContext dbContext, int comboId, bool isNew)
        {
            try
            {
                using (var cmd = dbContext.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE FileCombos
                        SET IsFilterComboNew = @IsFilterComboNew
                        WHERE ComboId = @ComboId";
                    cmd.Parameters.AddWithValue("@IsFilterComboNew", isNew ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ComboId", comboId);
                    
                    var affected = cmd.ExecuteNonQuery();
                    if (affected > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[FILE-COMBO-CHECK] ✅ Updated IsFilterComboNew={isNew} for FileCombo ComboId={comboId}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[FILE-COMBO-CHECK] ❌ Error updating IsFilterComboNew flag for ComboId={comboId}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Normalizes document key (same logic as ClashZoneRepository.NormalizeDocumentKey)
        /// </summary>
        private static string NormalizeDocumentKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Unknown";
            }

            var trimmed = value.Trim();

            // Strip ": <number> : location Shared" style suffixes
            var locationMatch = System.Text.RegularExpressions.Regex.Match(
                trimmed,
                @":\s*\d+\s*:\s*location\s+Shared",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (locationMatch.Success)
            {
                trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
            }

            // Remove trailing "(xx elements)" or similar
            var parenIndex = trimmed.IndexOf('(');
            if (parenIndex >= 0)
            {
                trimmed = trimmed.Substring(0, parenIndex).Trim();
            }

            // If the string is a full path, reduce to file name
            trimmed = System.IO.Path.GetFileName(trimmed);

            // Drop extension
            var extIndex = trimmed.LastIndexOf('.');
            if (extIndex >= 0)
            {
                trimmed = trimmed.Substring(0, extIndex);
            }

            return trimmed;
        }
        
        /// <summary>
        /// Checks if filter has existing clash zones (determines if it's fresh or non-fresh)
        /// </summary>
        private static bool HasExistingClashZones(RefreshContext context)
        {
            // Check if we have existing clash zones in context
            if (context.ExistingClashZones != null && context.ExistingClashZones.Count > 0)
            {
                return true;
            }
            
            // Check XML cache for existing zones
            if (context.XmlCache != null)
            {
                // Check Filter XML
                if (context.XmlCache.FilterXml != null)
                {
                    foreach (var filterXml in context.XmlCache.FilterXml.Values)
                    {
                        // ✅ FIX: filterXml IS a ClashZoneStorage, so access Filters directly
                        if (filterXml?.Filters != null)
                        {
                            foreach (var filterGroup in filterXml.Filters)
                            {
                                if (filterGroup?.FileCombos != null)
                                {
                                    foreach (var fileCombo in filterGroup.FileCombos)
                                    {
                                        if (fileCombo?.ClashZones != null && fileCombo.ClashZones.Count > 0)
                                        {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                        
                        // ✅ FALLBACK: Also check flat structure (backward compatibility)
                        if (filterXml?.ClashZones != null && filterXml.ClashZones.Count > 0)
                        {
                            return true;
                        }
                    }
                }
                
                // Check Global XML
                if (context.XmlCache.GlobalXml != null)
                {
                    foreach (var globalXml in context.XmlCache.GlobalXml.Values)
                    {
                        // ✅ FIX: CategoryGlobalIndex has Filters and Entries, not Categories
                        // Check hierarchical structure first
                        if (globalXml?.Filters != null)
                        {
                            foreach (var filterGroup in globalXml.Filters)
                            {
                                if (filterGroup?.FileCombos != null)
                                {
                                    foreach (var fileCombo in filterGroup.FileCombos)
                                    {
                                        if (fileCombo?.Entries != null && fileCombo.Entries.Count > 0)
                                        {
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                        
                        // ✅ FALLBACK: Also check flat structure (backward compatibility)
                        if (globalXml?.Entries != null && globalXml.Entries.Count > 0)
                        {
                            return true;
                        }
                    }
                }
            }
            
            return false;
        }
        
        /// <summary>
        /// ✅ PLACEMENT PATH: Determines placement path based on FileCombo existence in database
        /// FileCombo existence is checked FIRST (before IsFilterComboNew flag) because FileCombos are populated when clash zones are saved
        /// PATH 1 (Replay): FileCombos exist (clash zones have been saved before)
        /// PATH 2 (Sizing/Detection): No FileCombos exist (fresh filter+category combo, no clash zones saved yet)
        /// 
        /// ✅ NEW: Also checks if OpeningSettings have changed - if changed, routes to PATH 2 even if IsFilterComboNew=0
        /// </summary>
        public static SleevePlacementPath DeterminePlacementPath(
            Document document,
            string filterName,
            string category,
            string logPrefix = "",
            Dictionary<string, double> currentClearanceSettings = null)
        {
            try
            {
                // ✅ FIX: Check IsFilterComboNew flag PER FILE COMBO (not per filter+category)
                // If ANY file combo has IsFilterComboNew=1 → PATH 2 (Sizing/Detection)
                // If ALL file combos have IsFilterComboNew=0 → PATH 1 (Replay)
                bool hasNewFileCombos = HasNewFileCombos(document, filterName, category);
                
                if (hasNewFileCombos)
                {
                    // PATH 2: At least one file combo is new - needs full detection/sizing
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"{logPrefix}[PLACEMENT-PATH] PATH 2 (Sizing/Detection) selected: New file combos found for filter '{filterName}' + category '{category}' (IsFilterComboNew=1)");
                    }
                    return SleevePlacementPath.Sizing;
                }
                else
                {
                    // ✅ PATH 1: All file combos are used - check if conditions changed
                    // If OpeningSettings changed → Route to PATH 2 (Sizing) to recalculate sizes
                    if (currentClearanceSettings != null)
                    {
                        bool conditionsChanged = CheckConditionsChanged(document, filterName, category, currentClearanceSettings);
                        if (conditionsChanged)
                        {
                            // Conditions changed → Route to PATH 2 (Sizing)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"{logPrefix}[PLACEMENT-PATH] PATH 2 (Sizing) selected: OpeningSettings changed for filter '{filterName}' + category '{category}' (even though IsFilterComboNew=0)");
                            }
                            return SleevePlacementPath.Sizing;
                        }
                    }
                    
                    // PATH 1: All file combos are used and conditions unchanged - replay with saved data
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"{logPrefix}[PLACEMENT-PATH] PATH 1 (Replay) selected: All file combos used for filter '{filterName}' + category '{category}' (IsFilterComboNew=0, conditions unchanged)");
                    }
                    return SleevePlacementPath.Replay;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{logPrefix}[PLACEMENT-PATH] Error determining placement path: {ex.Message} - Defaulting to PATH 2 (Sizing/Detection)");
                return SleevePlacementPath.Sizing; // Safe default: assume fresh combo
            }
        }
        
        /// <summary>
        /// ✅ CONDITION CHANGE DETECTION: Checks if OpeningSettings have changed in UI
        /// Compares current UI OpeningSettings with saved OpeningSettings in database
        /// Returns true if conditions changed, false if unchanged
        /// </summary>
        private static bool CheckConditionsChanged(
            Document document,
            string filterName,
            string category,
            Dictionary<string, double> currentClearanceSettings)
        {
            try
            {
                // Load saved OpeningSettings from database
                using (var dbContext = new Data.SleeveDbContext(document))
                {
                    var filterRepository = new Data.Repositories.FilterRepository(dbContext, _ => { });
                    var (_, savedOpeningSettings, _, _, _) = filterRepository.LoadFilterUIState(filterName, category);
                    
                    if (savedOpeningSettings == null)
                    {
                        // No saved settings → treat as changed (first time)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CONDITION-CHECK] No saved OpeningSettings found for filter '{filterName}' + category '{category}' - treating as changed");
                        }
                        return true;
                    }
                    
                    // Compare clearance settings
                    if (savedOpeningSettings.ClearanceSettings == null)
                    {
                        // Saved settings have no clearance → treat as changed if current has clearance
                        if (currentClearanceSettings != null && currentClearanceSettings.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CONDITION-CHECK] Saved OpeningSettings has no ClearanceSettings, but current UI has {currentClearanceSettings.Count} settings - conditions changed");
                            }
                            return true;
                        }
                        return false; // Both have no clearance settings
                    }
                    
                    // Compare each clearance setting
                    var savedClearance = savedOpeningSettings.ClearanceSettings;
                    bool hasChanges = false;
                    
                    // Check if any current setting differs from saved
                    foreach (var currentSetting in currentClearanceSettings)
                    {
                        if (!savedClearance.ContainsKey(currentSetting.Key))
                        {
                            // New setting added
                            hasChanges = true;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CONDITION-CHECK] New clearance setting '{currentSetting.Key}' = {currentSetting.Value} (not in saved settings)");
                            }
                            break;
                        }
                        
                        // Compare values (with tolerance for floating point)
                        double savedValue = savedClearance[currentSetting.Key];
                        if (Math.Abs(savedValue - currentSetting.Value) > 0.001)
                        {
                            // Value changed
                            hasChanges = true;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CONDITION-CHECK] Clearance setting '{currentSetting.Key}' changed: {savedValue} → {currentSetting.Value}");
                            }
                            break;
                        }
                    }
                    
                    // Check if any saved setting was removed
                    if (!hasChanges)
                    {
                        foreach (var savedSetting in savedClearance)
                        {
                            if (!currentClearanceSettings.ContainsKey(savedSetting.Key))
                            {
                                // Setting removed
                                hasChanges = true;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CONDITION-CHECK] Clearance setting '{savedSetting.Key}' removed from UI (was {savedSetting.Value})");
                                }
                                break;
                            }
                        }
                    }
                    
                    if (hasChanges && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CONDITION-CHECK] ✅ Conditions changed for filter '{filterName}' + category '{category}' - routing to PATH 2 (Sizing)");
                    }
                    else if (!hasChanges && !DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CONDITION-CHECK] ✅ Conditions unchanged for filter '{filterName}' + category '{category}' - using PATH 1 (Replay)");
                    }
                    
                    return hasChanges;
                }
            }
            catch (Exception ex)
            {
                // On error, assume conditions unchanged (safe default for PATH 1)
                DebugLogger.Warning($"[CONDITION-CHECK] Error checking condition changes: {ex.Message} - Assuming unchanged (PATH 1)");
                return false;
            }
        }
        
        /// <summary>
        /// ✅ FIX: Checks if ANY FileCombo has IsFilterComboNew=1 for filter+category combo
        /// If any file combo is new → PATH 2 (Sizing/Detection)
        /// If all file combos are used (IsFilterComboNew=0) → PATH 1 (Replay)
        /// </summary>
        private static bool HasNewFileCombos(Document document, string filterName, string category)
        {
            try
            {
                using (var dbContext = new Data.SleeveDbContext(document))
                {
                    var filterRepository = new Data.Repositories.FilterRepository(dbContext, _ => { });
                    int filterId = filterRepository.GetFilterId(filterName, category);
                    
                    if (filterId <= 0)
                    {
                        // Filter doesn't exist → all file combos are new
                        return true;
                    }
                    
                    using (var cmd = dbContext.Connection.CreateCommand())
                    {
                        // Check if ANY file combo has IsFilterComboNew=1 for this filter
                        cmd.CommandText = @"
                            SELECT COUNT(*) FROM FileCombos 
                            WHERE FilterId = @FilterId AND IsFilterComboNew = 1";
                        cmd.Parameters.AddWithValue("@FilterId", filterId);
                        
                        var count = cmd.ExecuteScalar();
                        if (count != null && Convert.ToInt32(count) > 0)
                        {
                            // At least one file combo is new
                            return true;
                        }
                    }
                }
                
                // All file combos are used (IsFilterComboNew=0) or no file combos exist
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[PLACEMENT-PATH] Error checking IsFilterComboNew flag: {ex.Message} - Defaulting to PATH 2 (Sizing)");
                return true; // Safe default: assume new file combos
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL: Checks if FileCombos exist in database for filter+category combo
        /// FileCombos are created when clash zones are saved, so this is the source of truth
        /// Returns true if FileCombos exist, false if none found
        /// </summary>
        private static bool CheckFileComboExistence(Document document, string filterName, string category)
        {
            try
            {
                // Get database context (disposes automatically)
                using (var dbContext = new Data.SleeveDbContext(document))
                {
                    using (var cmd = dbContext.Connection.CreateCommand())
                    {
                        // Check if any FileCombos exist for this filter+category combo
                        cmd.CommandText = @"
                            SELECT COUNT(*) FROM FileCombos fc
                            INNER JOIN Filters f ON fc.FilterId = f.FilterId
                            WHERE f.FilterName = @FilterName AND f.Category = @Category";
                        cmd.Parameters.AddWithValue("@FilterName", filterName);
                        cmd.Parameters.AddWithValue("@Category", category);
                        
                        var count = cmd.ExecuteScalar();
                        if (count != null)
                        {
                            int fileComboCount = Convert.ToInt32(count);
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[PLACEMENT-PATH] Found {fileComboCount} FileCombo(s) for filter '{filterName}' + category '{category}'");
                            }
                            return fileComboCount > 0;
                        }
                    }
                }
                
                // If no FileCombos found, this is a fresh combo
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[PLACEMENT-PATH] No FileCombos found for filter '{filterName}' + category '{category}' - fresh combo");
                }
                return false; // No FileCombos = fresh combo
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[PLACEMENT-PATH] Error checking FileCombo existence: {ex.Message} - assuming no FileCombos (fresh combo)");
                return false; // Safe default: assume fresh combo (no FileCombos)
            }
        }
    }
}

