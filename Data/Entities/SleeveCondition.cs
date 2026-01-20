using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Entities
{
    /// <summary>
    /// SQLite entity for Conditions table
    /// Maps to existing OpeningConditions model
    /// </summary>
    public class SleeveCondition
    {
        public int ConditionId { get; set; }
        public int FilterId { get; set; }
        public string Category { get; set; } = string.Empty;
        public double? RectNormal { get; set; }
        public double? RectInsulated { get; set; }
        public double? RoundNormal { get; set; }
        public double? RoundInsulated { get; set; }
        public double? PipesNormal { get; set; }
        public double? PipesInsulated { get; set; }
        public double? CableTrayTop { get; set; }
        public double? CableTrayOther { get; set; }
        public string? OpeningPrefs { get; set; } // JSON blob
        public DateTime UpdatedAt { get; set; }
    }
}

