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
    }
}
