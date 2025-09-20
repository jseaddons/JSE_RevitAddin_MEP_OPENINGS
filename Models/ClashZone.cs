using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a clash zone detected between MEP and structural elements
    /// </summary>
    public class ClashZone
    {
        /// <summary>
        /// Unique identifier for this clash zone
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// The MEP element involved in the clash
        /// </summary>
        public ElementId MepElementId { get; set; }
        
        /// <summary>
        /// The structural element involved in the clash
        /// </summary>
        public ElementId StructuralElementId { get; set; }
        
        /// <summary>
        /// The intersection point where the clash occurs
        /// </summary>
        public XYZ IntersectionPoint { get; set; }
        
        /// <summary>
        /// The bounding box of the clash zone
        /// </summary>
        public BoundingBoxXYZ ClashBoundingBox { get; set; }
        
        /// <summary>
        /// The diameter/size of the MEP element at this clash point
        /// </summary>
        public double MepElementSize { get; set; }
        
        /// <summary>
        /// The clearance required for this clash zone
        /// </summary>
        public double RequiredClearance { get; set; }
        
        /// <summary>
        /// Whether this clash zone has been resolved (sleeve placed)
        /// </summary>
        public bool IsResolved { get; set; } = false;
        
        /// <summary>
        /// The sleeve element ID if resolved
        /// </summary>
        public ElementId? ResolvedSleeveId { get; set; }
        
        /// <summary>
        /// When this clash zone was first detected
        /// </summary>
        public DateTime DetectedAt { get; set; } = DateTime.Now;
        
        /// <summary>
        /// When this clash zone was last updated
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Hash of the MEP element geometry for change detection
        /// </summary>
        public string MepElementGeometryHash { get; set; } = string.Empty;
        
        /// <summary>
        /// Hash of the structural element geometry for change detection
        /// </summary>
        public string StructuralElementGeometryHash { get; set; } = string.Empty;
        
        /// <summary>
        /// The document path where this clash was detected
        /// </summary>
        public string DocumentPath { get; set; } = string.Empty;
        
        /// <summary>
        /// Additional metadata about the clash
        /// </summary>
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();
    }
    
    /// <summary>
    /// Container for storing clash zones in a profile
    /// </summary>
    public class ClashZoneStorage
    {
        /// <summary>
        /// List of all detected clash zones
        /// </summary>
        public List<ClashZone> ClashZones { get; set; } = new List<ClashZone>();
        
        /// <summary>
        /// When this clash zone storage was created
        /// </summary>
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        
        /// <summary>
        /// When this clash zone storage was last updated
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        
        /// <summary>
        /// The document path this storage is associated with
        /// </summary>
        public string DocumentPath { get; set; } = string.Empty;
        
        /// <summary>
        /// Hash of the document for change detection
        /// </summary>
        public string DocumentHash { get; set; } = string.Empty;
        
        /// <summary>
        /// Version of the clash detection algorithm used
        /// </summary>
        public string AlgorithmVersion { get; set; } = "1.0";
        
        /// <summary>
        /// Settings used for clash detection
        /// </summary>
        public Dictionary<string, object> DetectionSettings { get; set; } = new Dictionary<string, object>();
    }
}
