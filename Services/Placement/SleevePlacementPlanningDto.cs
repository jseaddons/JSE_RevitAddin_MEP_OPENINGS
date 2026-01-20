using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Immutable data transfer object carrying pre-computed sleeve planning values
    /// for a single clash zone. Generated prior to Revit API element placement.
    /// </summary>
    public class SleevePlacementPlanningDto
    {
        public Guid ClashZoneId { get; }
        public string HostType { get; }
        public double RawMepSizeFt { get; }            // Unrounded base size (feet)
        public double InsulationThicknessFt { get; }   // Insulation thickness if present (feet)
        public double TargetWidthFt { get; }           // Planned sleeve width (feet)
        public double TargetHeightFt { get; }          // Planned sleeve height (feet)
        public double ClearanceFt { get; }             // Calculated clearance (feet)
        public double RequiredDepthFt { get; }         // Host thickness + clearance (feet)
        public double RotationAngleDeg { get; }        // Suggested rotation angle (degrees)
        public ClearanceRiskClassification ClearanceRisk { get; } // Risk category for visual diagnostics
        public bool ShouldSkip { get; }                // Indicates this zone should be skipped during placement
        public string SkipReason { get; }              // Reason for skip
        public string LogSummary { get; }              // Pre-built log line to batch-write later

        public SleevePlacementPlanningDto(
            Guid clashZoneId,
            string hostType,
            double rawMepSizeFt,
            double insulationThicknessFt,
            double targetWidthFt,
            double targetHeightFt,
            double clearanceFt,
            double requiredDepthFt,
            double rotationAngleDeg,
            ClearanceRiskClassification clearanceRisk,
            bool shouldSkip,
            string skipReason,
            string logSummary)
        {
            ClashZoneId = clashZoneId;
            HostType = hostType;
            RawMepSizeFt = rawMepSizeFt;
            InsulationThicknessFt = insulationThicknessFt;
            TargetWidthFt = targetWidthFt;
            TargetHeightFt = targetHeightFt;
            ClearanceFt = clearanceFt;
            RequiredDepthFt = requiredDepthFt;
            RotationAngleDeg = rotationAngleDeg;
            ClearanceRisk = clearanceRisk;
            ShouldSkip = shouldSkip;
            SkipReason = skipReason;
            LogSummary = logSummary;
        }
    }

    /// <summary>
    /// Classification of clearance risk level for a planned sleeve.
    /// Purely diagnostic – does not affect placement logic directly.
    /// </summary>
    public enum ClearanceRiskClassification
    {
        Unknown = 0,
        Low = 1,
        Medium = 2,
        High = 3,
        Critical = 4
    }
}
