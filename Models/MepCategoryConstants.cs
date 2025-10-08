using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Centralized MEP category names to avoid confusion (singular vs plural)
    /// Based on Revit API BuiltInCategory standard names
    /// </summary>
    public static class MepCategoryConstants
    {
        // ✅ Standard category names (match Revit API BuiltInCategory)
        public const string DUCTS = "Ducts";                    // OST_DuctCurves
        public const string PIPES = "Pipes";                    // OST_PipeCurves
        public const string CABLE_TRAYS = "Cable Trays";        // OST_CableTray
        public const string DUCT_ACCESSORIES = "Duct Accessories"; // OST_DuctAccessory
        
        // XML file suffixes (lowercase with underscores)
        public const string DUCTS_XML_SUFFIX = "ducts";
        public const string PIPES_XML_SUFFIX = "pipes";
        public const string CABLE_TRAYS_XML_SUFFIX = "cable_trays";
        public const string DUCT_ACCESSORIES_XML_SUFFIX = "duct_accessories";
        
        /// <summary>
        /// Get XML file suffix for a category name
        /// </summary>
        public static string GetXmlSuffix(string category)
        {
            return category switch
            {
                DUCTS => DUCTS_XML_SUFFIX,
                PIPES => PIPES_XML_SUFFIX,
                CABLE_TRAYS => CABLE_TRAYS_XML_SUFFIX,
                DUCT_ACCESSORIES => DUCT_ACCESSORIES_XML_SUFFIX,
                _ => category?.ToLower().Replace(" ", "_") ?? ""
            };
        }
        
        /// <summary>
        /// Get Revit BuiltInCategory for a category name
        /// </summary>
        public static BuiltInCategory GetBuiltInCategory(string category)
        {
            return category switch
            {
                DUCTS => BuiltInCategory.OST_DuctCurves,
                PIPES => BuiltInCategory.OST_PipeCurves,
                CABLE_TRAYS => BuiltInCategory.OST_CableTray,
                DUCT_ACCESSORIES => BuiltInCategory.OST_DuctAccessory,
                _ => BuiltInCategory.INVALID
            };
        }
        
        /// <summary>
        /// Normalize category name to standard (handles "Duct" → "Ducts", etc.)
        /// </summary>
        public static string Normalize(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return DUCTS; // Default
            
            var normalized = category.Trim();
            
            // Handle common variations
            if (normalized.Equals("Duct", System.StringComparison.OrdinalIgnoreCase))
                return DUCTS;
            
            if (normalized.Equals("Pipe", System.StringComparison.OrdinalIgnoreCase))
                return PIPES;
            
            if (normalized.Equals("Cable Tray", System.StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("CableTray", System.StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("Cabletray", System.StringComparison.OrdinalIgnoreCase))
                return CABLE_TRAYS;
            
            if (normalized.Equals("Duct Accessory", System.StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("DuctAccessory", System.StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("Damper", System.StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("Dampers", System.StringComparison.OrdinalIgnoreCase))
                return DUCT_ACCESSORIES;
            
            // Return as-is if no match
            return normalized;
        }
    }
}

