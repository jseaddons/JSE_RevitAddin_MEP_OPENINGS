using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs
{
    /// <summary>
    /// Event arguments for operation completion events
    /// </summary>
    public class OperationCompletedEventArgs : System.EventArgs
    {
        public string OperationName { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public TimeSpan Duration { get; set; }
        public DateTime Timestamp { get; set; }

        public OperationCompletedEventArgs()
        {
            Timestamp = DateTime.Now;
        }
    }
}
