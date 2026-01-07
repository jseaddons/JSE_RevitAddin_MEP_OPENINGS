using System;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Mark prefix configuration for MEPMARK parameter
    /// Stores project-level and category-specific prefixes
    /// Used for UI state transfer to mark parameter command
    /// </summary>
    public class MarkPrefixSettings
    {
        /// <summary>
        /// Project-level prefix (e.g., "JSE", "PROJ") - prepended to all marks
        /// </summary>
        public string ProjectPrefix { get; set; } = "";
        
        /// <summary>
        /// Discipline prefix for Ducts (default: "DCT")
        /// </summary>
        public string DuctPrefix { get; set; } = "DCT";
        
        /// <summary>
        /// Discipline prefix for Pipes (default: "PLU")
        /// </summary>
        public string PipePrefix { get; set; } = "PLU";
        
        /// <summary>
        /// Discipline prefix for Cable Trays (default: "ELE")
        /// </summary>
        public string CableTrayPrefix { get; set; } = "ELE";
        
        /// <summary>
        /// Discipline prefix for Duct Accessories/Dampers (default: "DMP")
        /// </summary>
        public string DamperPrefix { get; set; } = "DMP";
        
        /// <summary>
        /// If true, re-apply marks to all sleeves (overwrite existing marks with new prefix)
        /// If false, only mark sleeves that don't have a mark yet (default behavior)
        /// </summary>
        public bool RemarkAll { get; set; } = false;
        
        /// <summary>
        /// ✅ NEW: Number format for running numbers (default: "000" = 001, 002, 003...)
        /// Options: "00" (01, 02...), "000" (001, 002...), "0000" (0001, 0002...)
        /// </summary>
        public string NumberFormat { get; set; } = "000";
        
        /// <summary>
        /// ✅ NEW: Individual remark checkboxes for each prefix category
        /// </summary>
        public bool RemarkProjectPrefix { get; set; } = false;
        public bool RemarkDuctPrefix { get; set; } = true;
        public bool RemarkPipePrefix { get; set; } = true;
        public bool RemarkCableTrayPrefix { get; set; } = true;
        public bool RemarkDamperPrefix { get; set; } = true;
        
        /// <summary>
        /// ✅ BIM 360 OPTIMIZATION: If true, only mark sleeves visible in active view/sheet
        /// This enables per-sheet numbering and dramatically improves performance on BIM 360
        /// </summary>
        public bool ActiveViewOnly { get; set; } = false;

        /// <summary>
        /// ✅ NEW: Starting number for sequential numbering (default: 1)
        /// </summary>
        public int StartNumber { get; set; } = 1;
        
        /// <summary>
        /// ✅ NEW: If true, numbering will continue from the sequence used in another view/floor plan
        /// </summary>
        public bool UseContinueNumbering { get; set; } = false;
        
        /// <summary>
        /// ✅ NEW: Name of the view/floor plan to pull numbering sequence from
        /// </summary>
        public string? ContinueFromViewName { get; set; }


        
        /// <summary>
        /// ✅ NEW: System Type overrides for Ducts (System Type → Prefix mapping)
        /// Tier 2: Overrides discipline prefix when System Type matches
        /// Example: "Supply Air" → "V" (overrides "DCT")
        /// </summary>
        public Dictionary<string, string> DuctSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        
        /// <summary>
        /// ✅ NEW: System Type overrides for Pipes (System Type → Prefix mapping)
        /// Tier 2: Overrides discipline prefix when System Type matches
        /// Example: "Domestic Cold Water" → "CW" (overrides "PLU")
        /// </summary>
        public Dictionary<string, string> PipeSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        
        /// <summary>
        /// ✅ NEW: System Type overrides for Duct Accessories (System Type → Prefix mapping)
        /// Tier 2: Overrides discipline prefix when System Type matches
        /// Example: "Fire Damper" → "FD" (overrides "DMP")
        /// </summary>
        public Dictionary<string, string> DuctAccessoriesSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        
        /// <summary>
        /// ✅ NEW: Service Type overrides for Cable Trays (Service Type → Prefix mapping)
        /// Tier 2: Overrides discipline prefix when Service Type matches
        /// Example: "Power" → "P" (overrides "ELE")
        /// </summary>
        public Dictionary<string, string> CableTrayServiceTypeOverrides { get; set; } = new Dictionary<string, string>();
        
        /// <summary>
        /// Get discipline prefix for a specific category
        /// </summary>
        public string GetDisciplinePrefix(string category)
        {
            return category switch
            {
                "Ducts" => DuctPrefix,
                "Pipes" => PipePrefix,
                "Cable Trays" => CableTrayPrefix,
                "Duct Accessories" => DamperPrefix,
                _ => "OPN" // Generic fallback
            };
        }
        
        /// <summary>
        /// Get remark flag for a specific category
        /// Returns true if remark is enabled for this category, false otherwise
        /// ✅ CRITICAL FIX: 
        /// - If RemarkAll is true, always return true (remark all categories)
        /// - If RemarkProjectPrefix is true, return true for ALL categories (project-wide remark)
        /// - Otherwise, return the category-specific remark flag
        /// </summary>
        public bool GetRemarkFlag(string category)
        {
            // ✅ FIX: If RemarkAll is true, remark all categories regardless of individual flags
            if (RemarkAll)
                return true;
            
            // ✅ FIX: If RemarkProjectPrefix is true, remark all categories (project-wide remark)
            // This handles the "Project Prefix" checkbox which should remark all categories
            if (RemarkProjectPrefix)
                return true;
            
            // Otherwise, return category-specific remark flag
            return category switch
            {
                "Ducts" => RemarkDuctPrefix,
                "Pipes" => RemarkPipePrefix,
                "Cable Trays" => RemarkCableTrayPrefix,
                "Duct Accessories" => RemarkDamperPrefix,
                _ => false // Unknown categories: don't remark unless RemarkAll or RemarkProjectPrefix is true
            };
        }
        
        /// <summary>
        /// ✅ CRITICAL: Two-tier prefix resolution - System Type override takes PRECEDENCE over discipline prefix
        /// Uses case-insensitive matching to ensure system types are found correctly
        /// </summary>
        public string GetPrefixForElement(string category, string? systemType = null, string? serviceType = null)
        {
            // ✅ TIER 2 (HIGHEST PRIORITY): Check System Type override first (if applicable)
            // This MUST take precedence over discipline prefix
            NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] Resolving prefix for Category: {category}, SystemType: '{systemType}', ServiceType: '{serviceType}'");

            if (category == "Ducts" && !string.IsNullOrEmpty(systemType))
            {
                // ✅ CRITICAL FIX: Use case-insensitive matching for system type lookup
                var matchingOverride = DuctSystemTypeOverrides.FirstOrDefault(kvp => 
                    string.Equals(kvp.Key, systemType, StringComparison.OrdinalIgnoreCase));
                
                if (!string.IsNullOrEmpty(matchingOverride.Key))
                {
                    NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] ✅ SYSTEM TYPE OVERRIDE: Duct System Type '{systemType}' → Prefix '{matchingOverride.Value}'");
                    return matchingOverride.Value; // System Type prefix takes precedence
                }
            }
            
            if (category == "Pipes" && !string.IsNullOrEmpty(systemType))
            {
                // ✅ CRITICAL FIX: Use case-insensitive matching for system type lookup
                var matchingOverride = PipeSystemTypeOverrides.FirstOrDefault(kvp => 
                    string.Equals(kvp.Key, systemType, StringComparison.OrdinalIgnoreCase));
                
                if (!string.IsNullOrEmpty(matchingOverride.Key))
                {
                    NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] ✅ SYSTEM TYPE OVERRIDE: Pipe System Type '{systemType}' → Prefix '{matchingOverride.Value}'");
                    return matchingOverride.Value; // System Type prefix takes precedence
                }
            }
            
            if (category == "Duct Accessories" && !string.IsNullOrEmpty(systemType))
            {
                // ✅ CRITICAL FIX: Use case-insensitive matching for system type lookup
                var matchingOverride = DuctAccessoriesSystemTypeOverrides.FirstOrDefault(kvp => 
                    string.Equals(kvp.Key, systemType, StringComparison.OrdinalIgnoreCase));
                
                if (!string.IsNullOrEmpty(matchingOverride.Key))
                {
                    NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] ✅ SYSTEM TYPE OVERRIDE: Duct Accessories System Type '{systemType}' → Prefix '{matchingOverride.Value}'");
                    return matchingOverride.Value; // System Type prefix takes precedence
                }
            }
            
            if (category == "Cable Trays" && !string.IsNullOrEmpty(serviceType))
            {
                // ✅ CRITICAL FIX: Use case-insensitive matching for service type lookup
                var matchingOverride = CableTrayServiceTypeOverrides.FirstOrDefault(kvp => 
                    string.Equals(kvp.Key, serviceType, StringComparison.OrdinalIgnoreCase));
                
                if (!string.IsNullOrEmpty(matchingOverride.Key))
                {
                    NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] ✅ SERVICE TYPE OVERRIDE: Cable Tray Service Type '{serviceType}' → Prefix '{matchingOverride.Value}'");
                    return matchingOverride.Value; // Service Type prefix takes precedence
                }
            }
            
            // ✅ TIER 1 (FALLBACK): No system/service type override found - use discipline prefix
            var disciplinePrefix = GetDisciplinePrefix(category);
            NumberingDebugLogger.LogInfo($"[MarkPrefixSettings] No override found for System Type '{systemType}' in category '{category}' - using discipline prefix '{disciplinePrefix}'");
            return disciplinePrefix;
        }
    }
}
