using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// DTO containing pre-calculated placement data for a sleeve.
    /// Used to decouple the planning phase from the execution phase.
    /// </summary>
    public class SleevePlacementPlanningDto
    {
        public Guid ClashZoneId { get; }
        public string HostType { get; }
        public double TargetWidthFt { get; }
        public double TargetHeightFt { get; }
        public double TargetDiameterFt { get; }
        public double TargetDepthFt { get; } // ✅ Added to match usage in Orchestrator
        public double RequiredDepthFt { get; } // Keeping both for compatibility if needed
        public double RotationRadians { get; }
        public XYZ PlacementPoint { get; }
        public bool IsCircular { get; }
        public string SleeveFamilyName { get; }
        
        // Metadata for logging and risk assessment
        public double RawSleeveSizeFt { get; }
        public double InsulationThicknessFt { get; }
        public double ClearanceFt { get; }
        public ClearanceRiskClassification Risk { get; }
        public bool ShouldSkip { get; }
        public string SkipReason { get; }
        public string LogTrace { get; }

        public SleevePlacementPlanningDto(
            Guid clashZoneId,
            string hostType,
            double rawSleeveSizeFt,
            double insulationThicknessFt,
            double targetWidthFt,
            double targetHeightFt,
            double clearanceFt,
            double requiredDepthFt,
            double rotationRadians,
            ClearanceRiskClassification risk,
            bool shouldSkip,
            string skipReason,
            string logTrace,
            string familyName,
            XYZ placementPoint = null,
            bool isCircular = false)
        {
            ClashZoneId = clashZoneId;
            HostType = hostType;
            RawSleeveSizeFt = rawSleeveSizeFt;
            InsulationThicknessFt = insulationThicknessFt;
            TargetWidthFt = targetWidthFt;
            TargetHeightFt = targetHeightFt;
            TargetDiameterFt = isCircular ? Math.Max(targetWidthFt, targetHeightFt) : 0;
            ClearanceFt = clearanceFt;
            RequiredDepthFt = requiredDepthFt;
            TargetDepthFt = requiredDepthFt; // ✅ Initialize with same value
            RotationRadians = rotationRadians;
            Risk = risk;
            ShouldSkip = shouldSkip;
            SkipReason = skipReason;
            LogTrace = logTrace;
            SleeveFamilyName = familyName;
            PlacementPoint = placementPoint;
            IsCircular = isCircular;
        }
    }

    public enum ClearanceRiskClassification
    {
        Unknown,
        None,
        Low,
        Medium,
        High,
        Critical
    }
}
