using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Models.EventArgs
{
    /// <summary>
    /// Event arguments for status update events
    /// </summary>
    public class StatusUpdateEventArgs : System.EventArgs
    {
        public string Message { get; }
        public StatusType Type { get; }
        public DateTime Timestamp { get; }
        public string Operation { get; }

        public StatusUpdateEventArgs(string message, StatusType type, DateTime timestamp, string operation = "")
        {
            Message = message ?? throw new ArgumentNullException(nameof(message));
            Type = type;
            Timestamp = timestamp;
            Operation = operation ?? string.Empty;
        }
    }
}
