namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Aggregated result from the orchestrator's pipeline execution.
    /// Summarizes final placement counts and overall success for fail-safe reporting.
    /// </summary>
    public sealed class OrchestratorResult
    {
        public bool Success { get; }
        public int PlacedCount { get; }
        public int ErrorCount { get; }
        public string CorrelationId { get; }

        /// <summary>
        /// Number of cluster sleeves created (replaces multiple individual sleeves)
        /// </summary>
        public int ClusterCount { get; }

        /// <summary>
        /// Number of individual zones that were grouped into clusters
        /// </summary>
        public int ClusteredZoneCount { get; }

        public OrchestratorResult(bool success, int placedCount, int errorCount, string correlationId, int clusterCount = 0, int clusteredZoneCount = 0)
        {
            Success = success;
            PlacedCount = placedCount;
            ErrorCount = errorCount;
            CorrelationId = correlationId;
            ClusterCount = clusterCount;
            ClusteredZoneCount = clusteredZoneCount;
        }
    }
}
