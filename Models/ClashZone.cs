using System;
using System.Collections.Generic;
using System.Xml.Serialization;
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
        [XmlIgnore]
        public ElementId MepElementId { get; set; }
        
        /// <summary>
        /// XML serializable MEP element ID
        /// </summary>
        public int MepElementIdValue
        {
            get => MepElementId?.IntegerValue ?? -1;
            set => MepElementId = value > 0 ? new ElementId(value) : ElementId.InvalidElementId;
        }
        
        /// <summary>
        /// The MEP element unique ID for robust tracking across sessions
        /// Pre-calculated during refresh to avoid linked file access during placement
        /// </summary>
        public string MepElementUniqueId { get; set; } = string.Empty;
        
        /// <summary>
        /// The structural element involved in the clash
        /// </summary>
        [XmlIgnore]
        public ElementId StructuralElementId { get; set; }
        
        /// <summary>
        /// XML serializable structural element ID
        /// </summary>
        public int StructuralElementIdValue
        {
            get => StructuralElementId?.IntegerValue ?? -1;
            set => StructuralElementId = value > 0 ? new ElementId(value) : ElementId.InvalidElementId;
        }
        
        /// <summary>
        /// The intersection point where the clash occurs
        /// </summary>
        [XmlIgnore]
        public XYZ IntersectionPoint { get; set; }
        
        /// <summary>
        /// XML serializable intersection point X coordinate
        /// </summary>
        public double IntersectionPointX
        {
            get => IntersectionPoint?.X ?? 0.0;
            set { 
                if (IntersectionPoint == null) 
                    IntersectionPoint = new XYZ(value, 0, 0); 
                else 
                    IntersectionPoint = new XYZ(value, IntersectionPoint.Y, IntersectionPoint.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable intersection point Y coordinate
        /// </summary>
        public double IntersectionPointY
        {
            get => IntersectionPoint?.Y ?? 0.0;
            set { 
                if (IntersectionPoint == null) 
                    IntersectionPoint = new XYZ(0, value, 0); 
                else 
                    IntersectionPoint = new XYZ(IntersectionPoint.X, value, IntersectionPoint.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable intersection point Z coordinate
        /// </summary>
        public double IntersectionPointZ
        {
            get => IntersectionPoint?.Z ?? 0.0;
            set { 
                if (IntersectionPoint == null) 
                    IntersectionPoint = new XYZ(0, 0, value); 
                else 
                    IntersectionPoint = new XYZ(IntersectionPoint.X, IntersectionPoint.Y, value); 
            }
        }
        
        /// <summary>
        /// The bounding box of the clash zone
        /// </summary>
        [XmlIgnore]
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
        /// Whether this clash zone has been resolved (individual sleeve placed)
        /// </summary>
        public bool IsResolved { get; set; } = false;
        
        /// <summary>
        /// Whether this clash zone has been resolved by cluster sleeve
        /// </summary>
        public bool IsClusterResolved { get; set; } = false;
        
        /// <summary>
        /// Whether this clash zone is part of a cluster (used by clustering algorithm)
        /// </summary>
        public bool IsClustered { get; set; } = false;
        
        /// <summary>
        /// The individual sleeve element ID if resolved
        /// </summary>
        [XmlIgnore]
        public ElementId? ResolvedSleeveId { get; set; }
        
        /// <summary>
        /// The cluster sleeve element ID if cluster resolved
        /// </summary>
        [XmlIgnore]
        public ElementId? ClusterSleeveId { get; set; }
        
        /// <summary>
        /// The cluster sleeve element ID as integer (for XML serialization)
        /// </summary>
        public int ClusterSleeveInstanceId { get; set; } = -1;

        /// <summary>
        /// The placed sleeve instance ID (integer value for serialization and tracking)
        /// </summary>
        public int SleeveInstanceId { get; set; } = -1;
        
        /// <summary>
        /// The family name of the placed sleeve (e.g., "RectangularOpeningOnWall")
        /// </summary>
        public string SleeveFamilyName { get; set; } = string.Empty;
        
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
        /// ⚠️ CRITICAL PROPERTY - DO NOT REMOVE ⚠️
        /// The category of the MEP element (e.g., "Ducts", "Pipes", "Cable Trays", "Duct Accessories")
        /// This is essential for category-specific processing and validation
        /// Each placement service validates this to ensure it only processes its own category
        /// </summary>
        public string MepElementCategory { get; set; } = string.Empty;
        
        /// <summary>
        /// ⚠️ CRITICAL PROPERTY - DO NOT REMOVE ⚠️
        /// The shape of the duct (e.g., "Round", "Rectangular") - extracted from duct family name
        /// Only applicable for Ducts category
        /// Determines which sleeve family to use (DuctOpeningOnWall vs DuctOpeningOnWallround)
        /// </summary>
        public string DuctShape { get; set; } = string.Empty;
        
        /// <summary>
        /// ⚠️ CRITICAL PROPERTY - DO NOT REMOVE ⚠️
        /// The insulation type of the MEP element (e.g., "Normal", "Insulated")
        /// Used to determine which clearance value to apply (normal vs insulated)
        /// </summary>
        public string InsulationType { get; set; } = "Normal";
        
        /// <summary>
        /// The formatted size of the MEP element (e.g., "600x300", "Ø200")
        /// Pre-calculated during refresh to avoid linked file access during placement
        /// </summary>
        public string MepElementFormattedSize { get; set; } = string.Empty;
        
        /// <summary>
        /// The system abbreviation of the MEP element (e.g., "SA", "RA", "EX")
        /// Pre-calculated during refresh to avoid linked file access during placement
        /// </summary>
        public string MepElementSystemAbbreviation { get; set; } = string.Empty;
        
        /// <summary>
        /// For fire dampers: connector side direction ("Left", "Right", "Top", "Bottom")
        /// Used to determine offset direction for MSFD dampers
        /// </summary>
        public string DamperConnectorSide { get; set; } = string.Empty;
        
        /// <summary>
        /// For fire dampers: whether this is an MSFD (multi-smoke fire damper) type
        /// MSFD dampers require offset placement toward connector side
        /// </summary>
        public bool IsMSFDDamper { get; set; } = false;
        
        /// <summary>
        /// The document path where this clash was detected
        /// </summary>
        public string DocumentPath { get; set; } = string.Empty;
        
        /// <summary>
        /// The document title where the structural element is located (for linked elements)
        /// </summary>
        public string StructuralElementDocumentTitle { get; set; } = string.Empty;
        
        /// <summary>
        /// The type of structural element (Wall, Structural Framing, Floor)
        /// </summary>
        public string StructuralElementType { get; set; } = string.Empty;
        
        /// <summary>
        /// The orientation of the host element (X or Y for walls/framing, blank for floors)
        /// Pre-calculated during refresh for efficient clustering and orientation logic
        /// </summary>
        public string HostOrientation { get; set; } = string.Empty;
        
        /// <summary>
        /// The thickness of the structural element (for depth calculation)
        /// </summary>
        public double StructuralElementThickness { get; set; } = 0.0;
        
        /// <summary>
        /// Pre-calculated structural element normal/direction for orientation calculation
        /// For walls: wall normal vector
        /// For floors: not needed (use MEP orientation)
        /// For framing: framing direction vector
        /// </summary>
        [XmlIgnore]
        public XYZ StructuralElementNormal { get; set; }
        
        /// <summary>
        /// XML serializable structural element normal X coordinate
        /// </summary>
        public double StructuralElementNormalX
        {
            get => StructuralElementNormal?.X ?? 0.0;
            set { 
                if (StructuralElementNormal == null) 
                    StructuralElementNormal = new XYZ(value, 0, 0); 
                else 
                    StructuralElementNormal = new XYZ(value, StructuralElementNormal.Y, StructuralElementNormal.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable structural element normal Y coordinate
        /// </summary>
        public double StructuralElementNormalY
        {
            get => StructuralElementNormal?.Y ?? 0.0;
            set { 
                if (StructuralElementNormal == null) 
                    StructuralElementNormal = new XYZ(0, value, 0); 
                else 
                    StructuralElementNormal = new XYZ(StructuralElementNormal.X, value, StructuralElementNormal.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable structural element normal Z coordinate
        /// </summary>
        public double StructuralElementNormalZ
        {
            get => StructuralElementNormal?.Z ?? 0.0;
            set { 
                if (StructuralElementNormal == null) 
                    StructuralElementNormal = new XYZ(0, 0, value); 
                else 
                    StructuralElementNormal = new XYZ(StructuralElementNormal.X, StructuralElementNormal.Y, value); 
            }
        }
        
        // NEW: Pre-calculated placement data (calculated during refresh, used during placement)
        
        /// <summary>
        /// Pre-calculated final placement point for the sleeve (no calculation needed during placement)
        /// </summary>
        [XmlIgnore]
        public XYZ SleevePlacementPoint { get; set; }
        
        /// <summary>
        /// XML serializable sleeve placement point X coordinate
        /// </summary>
        public double SleevePlacementPointX
        {
            get => SleevePlacementPoint?.X ?? 0.0;
            set { 
                if (SleevePlacementPoint == null) 
                    SleevePlacementPoint = new XYZ(value, 0, 0); 
                else 
                    SleevePlacementPoint = new XYZ(value, SleevePlacementPoint.Y, SleevePlacementPoint.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Y coordinate
        /// </summary>
        public double SleevePlacementPointY
        {
            get => SleevePlacementPoint?.Y ?? 0.0;
            set { 
                if (SleevePlacementPoint == null) 
                    SleevePlacementPoint = new XYZ(0, value, 0); 
                else 
                    SleevePlacementPoint = new XYZ(SleevePlacementPoint.X, value, SleevePlacementPoint.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Z coordinate
        /// </summary>
        public double SleevePlacementPointZ
        {
            get => SleevePlacementPoint?.Z ?? 0.0;
            set { 
                if (SleevePlacementPoint == null) 
                    SleevePlacementPoint = new XYZ(0, 0, value); 
                else 
                    SleevePlacementPoint = new XYZ(SleevePlacementPoint.X, SleevePlacementPoint.Y, value); 
            }
        }
        
        /// <summary>
        /// Pre-calculated MEP element width including clearance (no linked file access needed during placement)
        /// </summary>
        public double MepElementWidth { get; set; }
        
        /// <summary>
        /// Pre-calculated MEP element height including clearance (no linked file access needed during placement)
        /// </summary>
        public double MepElementHeight { get; set; }
        
        /// <summary>
        /// Pre-calculated MEP element orientation vector (no linked file access needed during placement)
        /// </summary>
        [XmlIgnore]
        public XYZ MepElementOrientation { get; set; }
        
        /// <summary>
        /// Pre-calculated MEP element orientation direction ("X" or "Y") for sleeve rotation
        /// </summary>
        public string MepElementOrientationDirection { get; set; }
        
        /// <summary>
        /// XML serializable MEP element orientation X component
        /// </summary>
        public double MepElementOrientationX
        {
            get => MepElementOrientation?.X ?? 0.0;
            set { 
                if (MepElementOrientation == null) 
                    MepElementOrientation = new XYZ(value, 0, 0); 
                else 
                    MepElementOrientation = new XYZ(value, MepElementOrientation.Y, MepElementOrientation.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable MEP element orientation Y component
        /// </summary>
        public double MepElementOrientationY
        {
            get => MepElementOrientation?.Y ?? 0.0;
            set { 
                if (MepElementOrientation == null) 
                    MepElementOrientation = new XYZ(0, value, 0); 
                else 
                    MepElementOrientation = new XYZ(MepElementOrientation.X, value, MepElementOrientation.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable MEP element orientation Z component
        /// </summary>
        public double MepElementOrientationZ
        {
            get => MepElementOrientation?.Z ?? 0.0;
            set { 
                if (MepElementOrientation == null) 
                    MepElementOrientation = new XYZ(0, 0, value); 
                else 
                    MepElementOrientation = new XYZ(MepElementOrientation.X, MepElementOrientation.Y, value); 
            }
        }
        
        /// <summary>
        /// Pipe opening type for family selection ("Circular" or "Rectangular", empty for non-pipes)
        /// </summary>
        public string PipeOpeningType { get; set; } = string.Empty;
        
        /// <summary>
        /// Pre-calculated MEP element level name (no linked file access needed during placement)
        /// </summary>
        public string MepElementLevelName { get; set; } = string.Empty;
        
        /// <summary>
        /// Pre-calculated MEP element level elevation (no linked file access needed during placement)
        /// </summary>
        public double MepElementLevelElevation { get; set; } = 0.0;
        
        /// <summary>
        /// Additional metadata about the clash
        /// </summary>
        [XmlIgnore]
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Snapshot of selected MEP parameter values at Refresh time (whitelisted keys)
        /// </summary>
        [XmlArray("MepParameterValues")]
        [XmlArrayItem("Param")]
        public List<SerializableKeyValue> MepParameterValues { get; set; } = new List<SerializableKeyValue>();

        /// <summary>
        /// Snapshot of selected Host parameter values at Refresh time (whitelisted keys)
        /// </summary>
        [XmlArray("HostParameterValues")]
        [XmlArrayItem("Param")]
        public List<SerializableKeyValue> HostParameterValues { get; set; } = new List<SerializableKeyValue>();

        /// <summary>
        /// Document key for MEP element (distinguishes active vs linked docs)
        /// </summary>
        public string SourceDocKey { get; set; } = string.Empty;

        /// <summary>
        /// Document key for Host element (distinguishes active vs linked docs)
        /// </summary>
        public string HostDocKey { get; set; } = string.Empty;

        /// <summary>
        /// Whether this clash is eligible for processing under the current UI selection (host types, etc.)
        /// This flag is set during refresh and used during placement to respect UI state without deleting zones.
        /// </summary>
        public bool IsEligibleByCurrentUi { get; set; } = true;
    }
    
    /// <summary>
    /// Container for storing clash zones in a profile
    /// </summary>
    [XmlRoot("ClashZoneStorage")]
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
        [XmlIgnore]
        public Dictionary<string, object> DetectionSettings { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// Whitelisted parameter keys to snapshot (persisted once per file)
        /// </summary>
        [XmlArray("ParameterKeyWhitelist")]
        [XmlArrayItem("Key")]
        public List<string> ParameterKeyWhitelist { get; set; } = new List<string>();

        /// <summary>
        /// Additional keys learned on-demand during mapping, merged into whitelist next run
        /// </summary>
        [XmlArray("LearnedParameterKeys")]
        [XmlArrayItem("Key")]
        public List<string> LearnedParameterKeys { get; set; } = new List<string>();
    }
}

