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
    }
    
    /// <summary>
    /// Clearance settings for sleeve placement
    /// </summary>
    public class ClearanceSettings
    {
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


