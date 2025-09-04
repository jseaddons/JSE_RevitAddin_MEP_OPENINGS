using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs
{
    /// <summary>
    /// Event arguments for progress update events
    /// </summary>
    public class ProgressUpdateEventArgs : System.EventArgs
    {
        public string OperationName { get; set; } = string.Empty;
        public int Current { get; set; }
        public int Total { get; set; }
        public int Percentage { get; set; }
        public string StepDescription { get; set; } = string.Empty;
        public DateTime Timestamp { get; set; }

        public ProgressUpdateEventArgs()
        {
            Timestamp = DateTime.Now;
        }
    }
}
