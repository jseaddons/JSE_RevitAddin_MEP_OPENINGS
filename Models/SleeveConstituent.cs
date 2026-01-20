using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a constituent sleeve that is part of a combined sleeve.
    /// A constituent can be either an individual sleeve (ClashZone) or a cluster sleeve.
    /// </summary>
    public class SleeveConstituent
    {
        // ============================================================================
        // PRIMARY IDENTIFIERS
        // ============================================================================
        
        /// <summary>
        /// Database primary key (auto-increment)
        /// </summary>
        public int ConstituentId { get; set; }
        
        /// <summary>
        /// Foreign key to CombinedSleeves table
        /// </summary>
        public int CombinedSleeveId { get; set; }
        
        // ============================================================================
        // CONSTITUENT TYPE
        // ============================================================================
        
        /// <summary>
        /// Type of constituent: Individual or Cluster
        /// </summary>
        public ConstituentType Type { get; set; }
        
        /// <summary>
        /// MEP category of this constituent: 'Pipes', 'Ducts', 'Cable Trays', 'Conduits'
        /// </summary>
        public string Category { get; set; }
        
        // ============================================================================
        // INDIVIDUAL SLEEVE REFERENCES (ConstituentType = Individual)
        // ============================================================================
        
        /// <summary>
        /// Foreign key to ClashZones table (for individual sleeves)
        /// Null if this is a cluster constituent
        /// </summary>
        public int? ClashZoneId { get; set; }
        
        /// <summary>
        /// Deterministic GUID for individual sleeve lookup
        /// Null if this is a cluster constituent
        /// </summary>
        public Guid? ClashZoneGuid { get; set; }
        
        // ============================================================================
        // CLUSTER SLEEVE REFERENCES (ConstituentType = Cluster)
        // ============================================================================
        
        /// <summary>
        /// Foreign key to ClusterSleeves table (for cluster sleeves)
        /// Null if this is an individual constituent
        /// </summary>
        public int? ClusterSleeveId { get; set; }
        
        /// <summary>
        /// Revit element ID for cluster sleeve (ElementId.IntegerValue)
        /// Null if this is an individual constituent
        /// </summary>
        public int? ClusterInstanceId { get; set; }
        
        // ============================================================================
        // TIMESTAMPS
        // ============================================================================
        
        /// <summary>
        /// Timestamp when this constituent was added to the combined sleeve
        /// </summary>
        public DateTime CreatedAt { get; set; }
    }
    
    /// <summary>
    /// Enum representing the type of sleeve constituent
    /// </summary>
    public enum ConstituentType
    {
        /// <summary>
        /// Individual sleeve (single clash zone)
        /// </summary>
        Individual = 0,
        
        /// <summary>
        /// Cluster sleeve (multiple clash zones grouped together)
        /// </summary>
        Cluster = 1
    }
}
