namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Result object for command execution with orchestrator
    /// </summary>
    public class CommandExecutionResult
    {
        /// <summary>
        /// Whether the command executed successfully
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// Error message if execution failed
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}

