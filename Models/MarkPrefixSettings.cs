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
        public bool RemarkDuctPrefix { get; set; } = false;
        public bool RemarkPipePrefix { get; set; } = false;
        public bool RemarkCableTrayPrefix { get; set; } = false;
        public bool RemarkDamperPrefix { get; set; } = false;
        
        /// <summary>
        /// ✅ NEW: System Type overrides for Ducts (System Type → Prefix mapping)
        /// Tier 2: Overrides discipline prefix when System Type matches
        /// Example: "Supply Air" → "V" (overrides "DCT")
        /// </summary>
        public Dictionary<string, string> DuctSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        
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
        /// ✅ CRITICAL FIX: If RemarkAll is true, always return true (remark all categories)
        /// Otherwise, return the category-specific remark flag
        /// </summary>
        public bool GetRemarkFlag(string category)
        {
            // ✅ CRITICAL FIX: If RemarkAll is true, remark all categories regardless of individual flags
            // This handles the "Project Prefix" checkbox which should remark all categories
            if (RemarkAll || RemarkProjectPrefix)
                return true;
            
            // Otherwise, return category-specific remark flag
            return category switch
            {
                "Ducts" => RemarkDuctPrefix,
                "Pipes" => RemarkPipePrefix,
                "Cable Trays" => RemarkCableTrayPrefix,
                "Duct Accessories" => RemarkDamperPrefix,
                _ => false // Unknown categories: don't remark unless RemarkAll is true
            };
        }
        
        /// <summary>
        /// ✅ NEW: Two-tier prefix resolution - checks System Type override first, then falls back to discipline prefix
        /// </summary>
        public string GetPrefixForElement(string category, string systemType = null, string serviceType = null)
        {
            // Tier 2: Check System Type override first (if applicable)
            if (category == "Ducts" && !string.IsNullOrEmpty(systemType) && DuctSystemTypeOverrides.ContainsKey(systemType))
            {
                return DuctSystemTypeOverrides[systemType]; // Override: System Type prefix
            }
            
            if (category == "Cable Trays" && !string.IsNullOrEmpty(serviceType) && CableTrayServiceTypeOverrides.ContainsKey(serviceType))
            {
                return CableTrayServiceTypeOverrides[serviceType]; // Override: Service Type prefix
            }
            
            // Tier 1: Fallback to discipline prefix
            return GetDisciplinePrefix(category);
        }
    }
}
