using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Entities
{
    /// <summary>
    /// Represents a persisted sleeve parameter snapshot record in SQLite.
    /// Keyed by either SleeveInstanceId (individual) or ClusterInstanceId (cluster).
    /// </summary>
    public class SleeveSnapshot
    {
        public int SnapshotId { get; set; }
        public int? SleeveInstanceId { get; set; }
        public int? ClusterInstanceId { get; set; }
        public string SourceType { get; set; } = "Individual"; // Individual | Cluster
        public int? FilterId { get; set; }
        public int? ComboId { get; set; }
        public string MepElementIdsJson { get; set; } = "[]";
        public string HostElementIdsJson { get; set; } = "[]";
        public string MepParametersJson { get; set; } = "{}";
        public string HostParametersJson { get; set; } = "{}";
        public string SourceDocKeysJson { get; set; } = "[]";
        public string HostDocKeysJson { get; set; } = "[]";
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}

