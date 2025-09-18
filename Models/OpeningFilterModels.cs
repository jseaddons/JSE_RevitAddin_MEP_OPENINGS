using System;
using System.Collections.Generic;

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
