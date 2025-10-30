using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// MEP categories for opening creation
    /// </summary>
    public enum MepCategory
    {
        Ducts,
        DuctAccessories,
        Pipes,
        CableTrays
    }

    /// <summary>
    /// Opening types for different MEP elements
    /// </summary>
    public enum OpeningType
    {
        RectangularSleeves,
        RectangularClusters,
        CircularSleeves
    }

    /// <summary>
    /// Represents a filter for opening creation based on MEP category and opening type
    /// </summary>
    [XmlRoot("OpeningFilter")]
    public class OpeningFilter
    {
        /// <summary>
        /// User-defined name for this filter (e.g., "Fire Fighting", "Data Devices")
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// MEP category this filter applies to
        /// </summary>
        public MepCategory Category { get; set; }

        /// <summary>
        /// Type of opening to create
        /// </summary>
        public OpeningType OpeningType { get; set; }

        /// <summary>
        /// Whether this filter is enabled
        /// </summary>
        public bool IsEnabled { get; set; } = true;

        /// <summary>
        /// Additional parameters specific to this filter
        /// </summary>
        [XmlIgnore]
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// When this filter was created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        /// <summary>
        /// When this filter was last modified
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;

        /// <summary>
        /// Snapshot of opening-related settings (e.g., clearances) captured when saving the filter
        /// </summary>
        public OpeningSettings? OpeningSettings { get; set; }

        /// <summary>
        /// Human-readable MEP category name as selected in UI (e.g., "Ducts", "Pipes")
        /// </summary>
        public string SelectedMepCategoryName { get; set; } = string.Empty;

        /// <summary>
        /// Selected MEP category names for UI restoration (supports multiple)
        /// </summary>
        [XmlArray("SelectedMepCategoryNames")]
        [XmlArrayItem("CategoryName")]
        public List<string> SelectedMepCategoryNames { get; set; } = new List<string>();

        /// <summary>
        /// Names of selected reference/linked files associated with this filter
        /// </summary>
        [XmlArray("SelectedReferenceFiles")]
        [XmlArrayItem("ReferenceFile")]
        public List<string> SelectedReferenceFiles { get; set; } = new List<string>();

        /// <summary>
        /// Names of selected host/linked files associated with this filter
        /// </summary>
        [XmlArray("SelectedHostFiles")]
        [XmlArrayItem("HostFile")]
        public List<string> SelectedHostFiles { get; set; } = new List<string>();

        /// <summary>
        /// Selected host element types (e.g., "Wall", "Floor", "Structural Framing") for UI restoration and filtering
        /// </summary>
        [XmlArray("SelectedHostElementTypes")]
        [XmlArrayItem("HostType")]
        public List<string> SelectedHostElementTypes { get; set; } = new List<string>();

        /// <summary>
        /// Clash zone storage for this filter
        /// </summary>
        public ClashZoneStorage? ClashZoneStorage { get; set; }


        /// <summary>
        /// Create a default filter for a given category
        /// </summary>
        public static OpeningFilter CreateDefault(MepCategory category, string name)
        {
            return new OpeningFilter
            {
                Name = name,
                Category = category,
                OpeningType = GetDefaultOpeningType(category),
                IsEnabled = true
            };
        }

        /// <summary>
        /// Get the default opening type for a category
        /// </summary>
        private static OpeningType GetDefaultOpeningType(MepCategory category)
        {
            return category switch
            {
                MepCategory.Ducts => OpeningType.RectangularSleeves,
                MepCategory.DuctAccessories => OpeningType.RectangularSleeves,
                MepCategory.Pipes => OpeningType.RectangularSleeves,
                MepCategory.CableTrays => OpeningType.RectangularSleeves,
                _ => OpeningType.RectangularSleeves
            };
        }

        /// <summary>
        /// Get a human-readable description of this filter
        /// </summary>
        public string GetDescription()
        {
            return $"{Name} - {Category} ({OpeningType})";
        }

        /// <summary>
        /// Clone this filter
        /// </summary>
        public OpeningFilter Clone()
        {
            return new OpeningFilter
            {
                Name = this.Name,
                Category = this.Category,
                OpeningType = this.OpeningType,
                IsEnabled = this.IsEnabled,
                Parameters = new Dictionary<string, object>(this.Parameters),
                CreatedDate = this.CreatedDate,
                LastModified = DateTime.Now
            };
        }
    }
}

