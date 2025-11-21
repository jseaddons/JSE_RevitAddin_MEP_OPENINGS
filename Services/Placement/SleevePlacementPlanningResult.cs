using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Aggregate result returned by planner containing DTO list and summary metrics.
    /// </summary>
    public class SleevePlacementPlanningResult
    {
        public IReadOnlyList<SleevePlacementPlanningDto> Items { get; }
        public int TotalCount { get; }
        public int SkippedCount { get; }
        public int HighRiskCount { get; }
        public int CriticalRiskCount { get; }
        public double PlanningDurationMs { get; }

        public SleevePlacementPlanningResult(
            IReadOnlyList<SleevePlacementPlanningDto> items,
            int skipped,
            int highRisk,
            int criticalRisk,
            double durationMs)
        {
            Items = items;
            TotalCount = items?.Count ?? 0;
            SkippedCount = skipped;
            HighRiskCount = highRisk;
            CriticalRiskCount = criticalRisk;
            PlanningDurationMs = durationMs;
        }
    }
}
