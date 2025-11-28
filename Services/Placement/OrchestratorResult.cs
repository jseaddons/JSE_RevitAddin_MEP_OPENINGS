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

        public OrchestratorResult(bool success, int placedCount, int errorCount, string correlationId)
        {
            Success = success;
            PlacedCount = placedCount;
            ErrorCount = errorCount;
            CorrelationId = correlationId;
        }
    }
}
