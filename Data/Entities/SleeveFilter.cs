using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Entities
{
    /// <summary>
    /// SQLite entity for Filters table
    /// Maps to existing OpeningFilter model
    /// </summary>
    public class SleeveFilter
    {
        public int FilterId { get; set; }
        public string FilterName { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }
}

