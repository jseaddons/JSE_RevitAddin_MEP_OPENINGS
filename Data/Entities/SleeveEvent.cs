using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Entities
{
    /// <summary>
    /// SQLite entity for SleeveEvents table
    /// Audit trail for sleeve placement events
    /// </summary>
    public class SleeveEvent
    {
        public int EventId { get; set; }
        public int ClashZoneId { get; set; }
        public string EventType { get; set; } = string.Empty; // "Placed", "Deleted", "Clustered"
        public string? Payload { get; set; } // JSON for diagnostics
        public DateTime CreatedAt { get; set; }
    }
}

