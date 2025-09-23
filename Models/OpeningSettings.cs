using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents opening-specific settings and parameters
    /// </summary>
    [XmlRoot("OpeningSettings")]
    public class OpeningSettings
    {
        /// <summary>
        /// Selected MEP type (Pipe, Duct, Cable Tray, etc.)
        /// </summary>
        public string SelectedMepType { get; set; } = "Pipe";
        
        /// <summary>
        /// Opening type (Rectangular, Circular)
        /// </summary>
        public string OpeningType { get; set; } = "Rectangular";
        
        /// <summary>
        /// Clearance settings for different MEP types
        /// </summary>
        [XmlIgnore]
        public Dictionary<string, double> ClearanceSettings { get; set; } = new Dictionary<string, double>();

        /// <summary>
        /// Serializable clearance settings for XML serialization
        /// </summary>
        [XmlArray("ClearanceSettings")]
        [XmlArrayItem("ClearanceSetting")]
        public ClearanceSetting[] SerializableClearanceSettings
        {
            get
            {
                var settings = new List<ClearanceSetting>();
                foreach (var kvp in ClearanceSettings)
                {
                    settings.Add(new ClearanceSetting { Key = kvp.Key, Value = kvp.Value });
                }
                return settings.ToArray();
            }
            set
            {
                ClearanceSettings.Clear();
                if (value != null)
                {
                    foreach (var setting in value)
                    {
                        ClearanceSettings[setting.Key] = setting.Value;
                    }
                }
            }
        }
        
        /// <summary>
        /// Default clearance value
        /// </summary>
        public double DefaultClearance { get; set; } = 25.0; // mm
        
        /// <summary>
        /// Whether to create openings automatically
        /// </summary>
        public bool AutoCreateOpenings { get; set; } = true;
        
        /// <summary>
        /// Whether to update existing openings
        /// </summary>
        public bool UpdateExistingOpenings { get; set; } = true;
        
        /// <summary>
        /// Whether to delete orphaned openings
        /// </summary>
        public bool DeleteOrphanedOpenings { get; set; } = false;
        
        /// <summary>
        /// Opening family name to use
        /// </summary>
        public string OpeningFamilyName { get; set; } = "Rectangular Opening";
        
        /// <summary>
        /// When these settings were created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// When these settings were last modified
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// Serializable wrapper for clearance key-value pairs
    /// </summary>
    public class ClearanceSetting
    {
        [XmlAttribute("Key")]
        public string Key { get; set; } = string.Empty;

        [XmlAttribute("Value")]
        public double Value { get; set; }
    }
}
