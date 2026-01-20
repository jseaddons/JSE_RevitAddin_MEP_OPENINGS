using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Settings for sleeve sizing calculations
    /// </summary>
    public class SizingSettings
    {
        public bool EnableRounding { get; set; } = true;
        public double RoundingIncrementMm { get; set; } = 1.0;
        public double CircularToRectangularThresholdMm { get; set; } = 0.0;
    }
}
