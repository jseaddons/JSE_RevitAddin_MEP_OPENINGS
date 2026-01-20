using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Result object for orchestration operations
    /// </summary>
    public class OrchestrationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public List<string> ProcessedDisciplines { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();
        public int TotalCommandsExecuted { get; set; }
        public bool MarkingCompleted { get; set; }
    }
}

