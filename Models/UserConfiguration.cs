using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a user's configuration settings for the MEP Openings application
    /// </summary>
    public class UserConfiguration
    {
        /// <summary>
        /// List of selected reference files (MEP files)
        /// </summary>
        public List<string> SelectedReferenceFiles { get; set; } = new List<string>();
        
        /// <summary>
        /// List of selected host files (Architectural/Structural files)
        /// </summary>
        public List<string> SelectedHostFiles { get; set; } = new List<string>();
        
        /// <summary>
        /// List of selected MEP categories
        /// </summary>
        public List<string> SelectedMepCategories { get; set; } = new List<string>();
        
        /// <summary>
        /// List of selected host categories
        /// </summary>
        public List<string> SelectedHostCategories { get; set; } = new List<string>();
        
        /// <summary>
        /// Opening-specific settings
        /// </summary>
        public OpeningSettings OpeningSettings { get; set; } = new OpeningSettings();
        
        /// <summary>
        /// When this configuration was created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// When this configuration was last modified
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;
        
        /// <summary>
        /// User who created this configuration
        /// </summary>
        public string CreatedBy { get; set; } = string.Empty;

        /// <summary>
        /// Advanced settings for ConVoid-like features.
        /// </summary>
        public SettingsModel AdvancedSettings { get; set; } = new SettingsModel();
        
        /// <summary>
        /// Storage for clash zones detected in this configuration
        /// </summary>
        public ClashZoneStorage ClashZoneStorage { get; set; } = new ClashZoneStorage();
    }
}
