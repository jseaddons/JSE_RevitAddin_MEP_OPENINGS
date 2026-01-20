using System;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the conditions/settings for opening placement for a specific filter
    /// Saved as FilterName_CONDITIONS.xml alongside the clash zone data
    /// This separates configuration (CONDITIONS) from data (CLASH ZONES)
    /// </summary>
    [XmlRoot("OpeningConditions")]
    public class OpeningConditions
    {
        /// <summary>
        /// The filter name this conditions file belongs to
        /// </summary>
        public string FilterName { get; set; } = string.Empty;
        
        /// <summary>
        /// The category this conditions file is for (Ducts, Pipes, Cable Trays, etc.)
        /// </summary>
        public string Category { get; set; } = string.Empty;
        
        /// <summary>
        /// When these conditions were created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// When these conditions were last modified
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Clearance settings for different duct types
        /// </summary>
        public ClearanceSettings ClearanceSettings { get; set; } = new ClearanceSettings();
        
        /// <summary>
        /// Opening type preferences (Circular vs Rectangular)
        /// </summary>
        public OpeningTypePreferences OpeningTypePreferences { get; set; } = new OpeningTypePreferences();
        
        /// <summary>
        /// Level constraints for placement (future expansion)
        /// </summary>
        public LevelConstraints LevelConstraints { get; set; } = new LevelConstraints();
        
        /// <summary>
        /// Creation mode (future expansion)
        /// </summary>
        public string CreationMode { get; set; } = "Opening"; // Opening, Recess, Auto

        /// <summary>
        /// Settings for sleeve sizing (rounding, thresholds)
        /// </summary>
        public SizingSettings SizingSettings { get; set; } = new SizingSettings();
    }
    
    /// <summary>
    /// Clearance settings for sleeve placement
    /// </summary>
    public class ClearanceSettings
    {
        // Duct clearances
        /// <summary>
        /// Clearance for rectangular ducts (normal/non-insulated)
        /// </summary>
        public double RectangularNormal { get; set; } = 50.0;
        
        /// <summary>
        /// Clearance for rectangular ducts (insulated)
        /// </summary>
        public double RectangularInsulated { get; set; } = 25.0;
        
        /// <summary>
        /// Clearance for round ducts (normal/non-insulated)
        /// </summary>
        public double RoundNormal { get; set; } = 50.0;
        
        /// <summary>
        /// Clearance for round ducts (insulated)
        /// </summary>
        public double RoundInsulated { get; set; } = 50.0;
        
        // Pipe clearances
        /// <summary>
        /// Clearance for pipes (normal/non-insulated)
        /// </summary>
        public double PipesNormal { get; set; } = 50.0;
        
        /// <summary>
        /// Clearance for pipes (insulated)
        /// </summary>
        public double PipesInsulated { get; set; } = 25.0;
        
        // Cable Tray clearances
        /// <summary>
        /// Clearance for cable tray top side (normal/non-insulated)
        /// </summary>
        public double CableTrayTop { get; set; } = 100.0;
        
        /// <summary>
        /// Clearance for cable tray top side (insulated)
        /// </summary>
        public double CableTrayTopInsulated { get; set; } = 100.0;
        
        /// <summary>
        /// Clearance for cable tray other sides (left, right, bottom) (normal/non-insulated)
        /// </summary>
        public double CableTrayOther { get; set; } = 50.0;
        
        /// <summary>
        /// Clearance for cable tray other sides (left, right, bottom) (insulated)
        /// </summary>
        public double CableTrayOtherInsulated { get; set; } = 50.0;
        
        // Duct Accessory (Damper) clearances
        /// <summary>
        /// Clearance for duct accessories MEP side (for MSFD dampers - connector side) (normal/non-insulated)
        /// </summary>
        public double DuctAccessoryMepNormal { get; set; } = 100.0;
        
        /// <summary>
        /// Clearance for duct accessories MEP side (for MSFD dampers - connector side) (insulated)
        /// </summary>
        public double DuctAccessoryMepInsulated { get; set; } = 100.0;
        
        /// <summary>
        /// Clearance for duct accessories other sides (for Standard and MSFD non-connector sides) (normal/non-insulated)
        /// </summary>
        public double DuctAccessoryOtherNormal { get; set; } = 50.0;
        
        /// <summary>
        /// Clearance for duct accessories other sides (for Standard and MSFD non-connector sides) (insulated)
        /// </summary>
        public double DuctAccessoryOtherInsulated { get; set; } = 50.0;
    }
    
    /// <summary>
    /// Opening type preferences for different MEP elements
    /// </summary>
    public class OpeningTypePreferences
    {
        /// <summary>
        /// Opening type for round ducts (Circular or Rectangular)
        /// </summary>
        public string RoundDucts { get; set; } = "Circular";
        
        /// <summary>
        /// Opening type for pipes (Circular or Rectangular)
        /// </summary>
        public string Pipes { get; set; } = "Circular";
    }
    
    /// <summary>
    /// Level constraints for sleeve placement (future expansion)
    /// </summary>
    public class LevelConstraints
    {
        /// <summary>
        /// Level constraint for horizontal openings (walls)
        /// </summary>
        public string HorizontalLevel { get; set; } = "Host Level";
        
        /// <summary>
        /// Level constraint for vertical openings (floors/ceilings)
        /// </summary>
        public string VerticalLevel { get; set; } = "Host Level";
    }
}


