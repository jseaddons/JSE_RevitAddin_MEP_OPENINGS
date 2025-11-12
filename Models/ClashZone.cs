using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

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
        
        // ✅ FIX 3: Backing fields to store X/Y/Z independently of XYZ object (24 bytes total, not 48)
        // This prevents XML serialization from reading 0 when IntersectionPoint is (0,0,0) or null
        private double _intersectionPointX = 0.0;
        private double _intersectionPointY = 0.0;
        private double _intersectionPointZ = 0.0;
        
        /// <summary>
        /// The intersection point where the clash occurs
        /// ✅ FIX 3: Computed from backing fields to avoid duplicate storage (reduces memory by 24 bytes)
        /// </summary>
        [XmlIgnore]
        public XYZ IntersectionPoint
        {
            get => new XYZ(_intersectionPointX, _intersectionPointY, _intersectionPointZ);
            set
            {
                // ✅ FIX 3: Only update backing fields, don't store XYZ object
                _intersectionPointX = value?.X ?? 0.0;
                _intersectionPointY = value?.Y ?? 0.0;
                _intersectionPointZ = value?.Z ?? 0.0;
            }
        }
        
        /// <summary>
        /// XML serializable intersection point X coordinate
        /// </summary>
        public double IntersectionPointX
        {
            get => _intersectionPointX;
            set => _intersectionPointX = value;
        }
        
        /// <summary>
        /// XML serializable intersection point Y coordinate
        /// </summary>
        public double IntersectionPointY
        {
            get => _intersectionPointY;
            set => _intersectionPointY = value;
        }
        
        /// <summary>
        /// XML serializable intersection point Z coordinate
        /// </summary>
        public double IntersectionPointZ
        {
            get => _intersectionPointZ;
            set => _intersectionPointZ = value;
        }
        
        /// <summary>
        /// The bounding box of the clash zone
        /// </summary>
        [XmlIgnore]
        public BoundingBoxXYZ ClashBoundingBox { get; set; }
        
        /// <summary>
        /// ✅ NEW: XML-serializable bounding box coordinates for clustering
        /// These are calculated during sleeve placement and used for cheap proximity detection
        /// </summary>
        public double SleeveBoundingBoxMinX { get; set; } = 0.0;
        public double SleeveBoundingBoxMinY { get; set; } = 0.0;
        public double SleeveBoundingBoxMinZ { get; set; } = 0.0;
        public double SleeveBoundingBoxMaxX { get; set; } = 0.0;
        public double SleeveBoundingBoxMaxY { get; set; } = 0.0;
        public double SleeveBoundingBoxMaxZ { get; set; } = 0.0;
        
        /// <summary>
        /// The diameter/size of the MEP element at this clash point
        /// </summary>
        public double MepElementSize { get; set; }
        
        // Suppress legacy field in XML output
        public bool ShouldSerializeMepElementSize()
        {
            return false;
        }
        
        /// <summary>
        /// ✅ NEW: Detailed MEP element size information with insulation data
        /// </summary>
        public MepElementSize MepElementSizeData { get; set; } = new MepElementSize();
        
        /// <summary>
        /// The clearance required for this clash zone
        /// </summary>
        public double RequiredClearance { get; set; }
        
        /// <summary>
        /// Whether this clash zone has been resolved (individual sleeve placed)
        /// ✅ FLAG PERSISTENCE: Stored in both Global XML and Filter XML so refresh + placement share a single view of flag state
        /// (Global XML remains the source of truth; Filter XML copy assists diagnostics and legacy tools.)
        /// </summary>
        public bool IsResolved { get; set; } = false;
        
        /// <summary>
        /// Whether this clash zone has been resolved by cluster sleeve
        /// ✅ FLAG PERSISTENCE: Stored in both Global XML and Filter XML for transparency; Global XML is still authoritative.
        /// </summary>
        public bool IsClusterResolved { get; set; } = false;
        
        /// <summary>
        /// ✅ DUCT-DAMPER COMBO FLAG: Indicates this duct is near a damper and should be skipped
        /// Set to true when duct-damper combo is detected, prevents re-checking on subsequent refreshes
        /// </summary>
        public bool HasDamperNearby { get; set; } = false;
        
        /// <summary>
        /// REMOVED: IsClustered flag - replaced by MarkedForClusteringSleeveProcess
        /// </summary>

        /// <summary>
        /// Flag indicating if this clash zone was detected in the current refresh
        /// true = detected in current refresh (new clash)
        /// false = loaded from previous XML (old clash)
        /// Used for debugging filtering effectiveness
        /// </summary>
        public bool IsCurrentClash { get; set; } = false; // ✅ CRITICAL FIX: Default to false for XML-loaded clashes
        
        /// <summary>
        /// CLEAR FLAG: Indicates this sleeve should be processed for cluster placement
        /// true = sleeve is proximate to other sleeves and should be clustered
        /// false = sleeve should remain individual (not proximate)
        /// null = not yet processed for clustering
        /// This replaces the confusing IsClustered flag logic
        /// </summary>
        public bool? MarkedForClusteringSleeveProcess { get; set; } = null;
        
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
        /// ✅ STORAGE: Original SleeveInstanceId stored BEFORE cluster placement sets it to -1
        /// This allows cleanup to find individual sleeves even after they're marked as cluster-resolved
        /// </summary>
        public int AfterClusterSleevePlacedSleeveInstanceId { get; set; } = -1;

        /// <summary>
        /// ✅ NEW: Cluster sleeve bounding box coordinates (for XML-serializable cleanup detection)
        /// These are saved after placing cluster sleeves to enable cheap cleanup of individual sleeves within clusters
        /// </summary>
        public double ClusterSleeveBoundingBoxMinX { get; set; } = 0.0;
        public double ClusterSleeveBoundingBoxMinY { get; set; } = 0.0;
        public double ClusterSleeveBoundingBoxMinZ { get; set; } = 0.0;
        public double ClusterSleeveBoundingBoxMaxX { get; set; } = 0.0;
        public double ClusterSleeveBoundingBoxMaxY { get; set; } = 0.0;
        public double ClusterSleeveBoundingBoxMaxZ { get; set; } = 0.0;

        /// <summary>
        /// The placed sleeve instance ID (integer value for serialization and tracking)
        /// </summary>
        public int SleeveInstanceId { get; set; } = -1;
        
        /// <summary>
        /// The family name of the placed sleeve (e.g., "RectangularOpeningOnWall")
        /// </summary>
        public string SleeveFamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Width of the placed sleeve (in Revit internal units)
        /// Used for proximity calculation and clustering
        /// </summary>
        public double SleeveWidth { get; set; } = 0.0;

        /// <summary>
        /// Height of the placed sleeve (in Revit internal units)
        /// Used for proximity calculation and clustering
        /// </summary>
        public double SleeveHeight { get; set; } = 0.0;

        /// <summary>
        /// Diameter of the placed sleeve (in Revit internal units)
        /// Used for circular sleeves proximity calculation
        /// </summary>
        public double SleeveDiameter { get; set; } = 0.0;
        
        // ✅ FIX 3: Backing fields for sleeve placement point (24 bytes total, not 48)
        private double _sleevePlacementPointX = 0.0;
        private double _sleevePlacementPointY = 0.0;
        private double _sleevePlacementPointZ = 0.0;
        
        /// <summary>
        /// The actual placement point of the sleeve (in Revit internal units)
        /// Used for simple proximity calculation: check X diff and Y diff
        /// ✅ FIX 3: Computed from backing fields to avoid duplicate storage (reduces memory by 24 bytes)
        /// </summary>
        [XmlIgnore]
        public XYZ SleevePlacementPoint
        {
            get => new XYZ(_sleevePlacementPointX, _sleevePlacementPointY, _sleevePlacementPointZ);
            set
            {
                // ✅ FIX 3: Only update backing fields, don't store XYZ object
                _sleevePlacementPointX = value?.X ?? 0.0;
                _sleevePlacementPointY = value?.Y ?? 0.0;
                _sleevePlacementPointZ = value?.Z ?? 0.0;
            }
        }
        
        /// <summary>
        /// XML serializable sleeve placement point X coordinate
        /// </summary>
        public double SleevePlacementPointX
        {
            get => _sleevePlacementPointX;
            set => _sleevePlacementPointX = value;
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Y coordinate
        /// </summary>
        public double SleevePlacementPointY
        {
            get => _sleevePlacementPointY;
            set => _sleevePlacementPointY = value;
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Z coordinate
        /// </summary>
        public double SleevePlacementPointZ
        {
            get => _sleevePlacementPointZ;
            set => _sleevePlacementPointZ = value;
        }
        
        /// <summary>
        /// Sleeve placement point in ACTIVE document coordinates (for proximity calculation)
        /// </summary>
        [XmlIgnore]
        public XYZ SleevePlacementPointActiveDocument { get; set; }
        
        /// <summary>
        /// XML serializable sleeve placement point X coordinate in active document
        /// </summary>
        public double SleevePlacementPointActiveDocumentX
        {
            get => SleevePlacementPointActiveDocument?.X ?? 0.0;
            set { 
                if (SleevePlacementPointActiveDocument == null) 
                    SleevePlacementPointActiveDocument = new XYZ(value, 0, 0); 
                else 
                    SleevePlacementPointActiveDocument = new XYZ(value, SleevePlacementPointActiveDocument.Y, SleevePlacementPointActiveDocument.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Y coordinate in active document
        /// </summary>
        public double SleevePlacementPointActiveDocumentY
        {
            get => SleevePlacementPointActiveDocument?.Y ?? 0.0;
            set { 
                if (SleevePlacementPointActiveDocument == null) 
                    SleevePlacementPointActiveDocument = new XYZ(0, value, 0); 
                else 
                    SleevePlacementPointActiveDocument = new XYZ(SleevePlacementPointActiveDocument.X, value, SleevePlacementPointActiveDocument.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable sleeve placement point Z coordinate in active document
        /// </summary>
        public double SleevePlacementPointActiveDocumentZ
        {
            get => SleevePlacementPointActiveDocument?.Z ?? 0.0;
            set { 
                if (SleevePlacementPointActiveDocument == null) 
                    SleevePlacementPointActiveDocument = new XYZ(0, 0, value); 
                else 
                    SleevePlacementPointActiveDocument = new XYZ(SleevePlacementPointActiveDocument.X, SleevePlacementPointActiveDocument.Y, value);
            }
        }
        
        /// <summary>
        /// Ensures SleevePlacementPointActiveDocument is properly reconstructed from XML-serializable properties
        /// Call this after XML deserialization to ensure SleevePlacementPointActiveDocument is not null
        /// </summary>
        public void EnsureSleevePlacementPointActiveDocumentReconstructed()
        {
            if (SleevePlacementPointActiveDocument == null && (SleevePlacementPointActiveDocumentX != 0 || SleevePlacementPointActiveDocumentY != 0 || SleevePlacementPointActiveDocumentZ != 0))
            {
                SleevePlacementPointActiveDocument = new XYZ(SleevePlacementPointActiveDocumentX, SleevePlacementPointActiveDocumentY, SleevePlacementPointActiveDocumentZ);
            }
        }
        
        /// <summary>
        /// Ensures SleevePlacementPoint is properly reconstructed from XML-serializable properties
        /// Call this after XML deserialization to ensure SleevePlacementPoint is not null
        /// </summary>
        public void EnsureSleevePlacementPointReconstructed()
        {
            if (SleevePlacementPoint == null && (SleevePlacementPointX != 0 || SleevePlacementPointY != 0 || SleevePlacementPointZ != 0))
            {
                SleevePlacementPoint = new XYZ(SleevePlacementPointX, SleevePlacementPointY, SleevePlacementPointZ);
            }
        }
        
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
        /// ⚠️ CRITICAL PROPERTY - DO NOT REMOVE ⚠️
        /// Pre-calculated wall direction vector for robust X-wall/Y-wall detection
        /// For walls: actual wall direction (not normal) - calculated during refresh
        /// For framing: framing direction vector
        /// For floors: not applicable (use MEP orientation)
        /// This enables robust sleeve rotation without expensive Revit API calls during placement
        /// </summary>
        [XmlIgnore]
        public XYZ WallDirection { get; set; }
        
        /// <summary>
        /// XML serializable wall direction X coordinate
        /// </summary>
        public double WallDirectionX
        {
            get => WallDirection?.X ?? 0.0;
            set { 
                if (WallDirection == null) 
                    WallDirection = new XYZ(value, 0, 0); 
                else 
                    WallDirection = new XYZ(value, WallDirection.Y, WallDirection.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable wall direction Y coordinate
        /// </summary>
        public double WallDirectionY
        {
            get => WallDirection?.Y ?? 0.0;
            set { 
                if (WallDirection == null) 
                    WallDirection = new XYZ(0, value, 0); 
                else 
                    WallDirection = new XYZ(WallDirection.X, value, WallDirection.Z); 
            }
        }
        
        /// <summary>
        /// XML serializable wall direction Z coordinate
        /// </summary>
        public double WallDirectionZ
        {
            get => WallDirection?.Z ?? 0.0;
            set { 
                if (WallDirection == null) 
                    WallDirection = new XYZ(0, 0, value); 
                else 
                    WallDirection = new XYZ(WallDirection.X, WallDirection.Y, value); 
            }
        }
        
        /// <summary>
        /// Pre-calculated wall direction type ("X-WALL", "Y-WALL", "FRAMING", "FLOOR")
        /// Determined during refresh for efficient sleeve rotation logic
        /// </summary>
        public string WallDirectionType { get; set; } = string.Empty;
        
        /// <summary>
        /// The thickness of the structural element (for depth calculation)
        /// </summary>
        public double StructuralElementThickness { get; set; } = 0.0;
        
        /// <summary>
        /// Wall thickness (for walls only)
        /// </summary>
        public double WallThickness { get; set; } = 0.0;
        
        /// <summary>
        /// Structural framing parameter 'b' thickness (for structural framing only)
        /// </summary>
        public double FramingThickness { get; set; } = 0.0;
        
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
        
        /// <summary>
        /// ✅ NEW: Set bounding box coordinates from a Revit BoundingBoxXYZ
        /// This is called during sleeve placement to save bounding box for cheap clustering
        /// Coordinates are saved in world coordinates - clustering will use appropriate 2D projection
        /// </summary>
        public void SetSleeveBoundingBox(BoundingBoxXYZ boundingBox)
        {
            if (boundingBox != null)
            {
                // ✅ CORRECT: Save world coordinates as-is - clustering will use appropriate 2D projection
                SleeveBoundingBoxMinX = boundingBox.Min.X;
                SleeveBoundingBoxMinY = boundingBox.Min.Y;
                SleeveBoundingBoxMinZ = boundingBox.Min.Z;
                SleeveBoundingBoxMaxX = boundingBox.Max.X;
                SleeveBoundingBoxMaxY = boundingBox.Max.Y;
                SleeveBoundingBoxMaxZ = boundingBox.Max.Z;
            }
        }
        
        /// <summary>
        /// ✅ MEMORY OPTIMIZATION: Clear all Revit API objects (XYZ, BoundingBoxXYZ) after coordinates are extracted
        /// Call this after clash zone is created and all coordinates are saved to XML-serializable properties
        /// </summary>
        public void ClearRevitApiObjects()
        {
            // ✅ CRITICAL FIX: Do NOT clear IntersectionPoint - it's a computed property that reads from backing fields
            // Clearing it would zero out _intersectionPointX/Y/Z, losing the intersection point data!
            // IntersectionPoint getter creates a NEW XYZ from backing fields, so there's no heavy object to clear
            // ClashBoundingBox = null;  // ✅ PRESERVED: Keep bounding box for normalization fallback
            SleevePlacementPoint = null;
            SleevePlacementPointActiveDocument = null;
            WallDirection = null;
            StructuralElementNormal = null;
            MepElementOrientation = null;
            // Note: ElementId objects are value types (structs), so no need to clear them
        }
        
        /// <summary>
        /// ✅ CORRECT: Calculate minimum distance between two rectangles and check if within tolerance
        /// Uses 2D coordinates based on host type and orientation
        /// </summary>
        public bool IsBoundingBoxOverlapping(ClashZone other, double toleranceDistance)
        {
            if (other == null) return false;
            
            double minDistance;
            
            // ✅ DEBUG: Log orientation values
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect1: Min=({SleeveBoundingBoxMinX:F3}, {SleeveBoundingBoxMinY:F3}), Max=({SleeveBoundingBoxMaxX:F3}, {SleeveBoundingBoxMaxY:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect2: Min=({other.SleeveBoundingBoxMinX:F3}, {other.SleeveBoundingBoxMinY:F3}), Max=({other.SleeveBoundingBoxMaxX:F3}, {other.SleeveBoundingBoxMaxY:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] HostType={StructuralElementType}, Orientation={MepElementOrientationDirection}\n");
            
            // ✅ CORRECT ALGORITHM: Calculate actual minimum distance between rectangles
            if (StructuralElementType == "Floor")
            {
                // Floor sleeves: Use X,Y distance only (ignore Z coordinate)
                minDistance = CalculateMinimumDistance2D(
                    SleeveBoundingBoxMinX, SleeveBoundingBoxMinY, SleeveBoundingBoxMaxX, SleeveBoundingBoxMaxY,
                    other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY);
            }
            else if (StructuralElementType == "Wall" || StructuralElementType == "Structural Framing")
            {
                // ✅ FIX: Use orientation from grouping logic instead of MepElementOrientationDirection
                // For walls, we know the orientation from the sleeve placement logic
                if (MepElementOrientationDirection == "X" || MepElementOrientationDirection == null)
                {
                    // Wall/Framing sleeves (X orientation): Use X,Z distance only (ignore Y coordinate)
                    minDistance = CalculateMinimumDistance2D(
                        SleeveBoundingBoxMinX, SleeveBoundingBoxMinZ, SleeveBoundingBoxMaxX, SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxZ);
                }
                else if (MepElementOrientationDirection == "Y")
                {
                    // Wall/Framing sleeves (Y orientation): Use Y,Z distance only (ignore X coordinate)
                    minDistance = CalculateMinimumDistance2D(
                        SleeveBoundingBoxMinY, SleeveBoundingBoxMinZ, SleeveBoundingBoxMaxY, SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                }
                else
                {
                    // Default for walls: Use Y,Z distance (most walls are Y-oriented)
                    minDistance = CalculateMinimumDistance2D(
                        SleeveBoundingBoxMinY, SleeveBoundingBoxMinZ, SleeveBoundingBoxMaxY, SleeveBoundingBoxMaxZ,
                        other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
                }
            }
            else
            {
                // Fallback: Use 3D distance for unknown host types
                minDistance = CalculateMinimumDistance3D(
                    SleeveBoundingBoxMinX, SleeveBoundingBoxMinY, SleeveBoundingBoxMinZ, 
                    SleeveBoundingBoxMaxX, SleeveBoundingBoxMaxY, SleeveBoundingBoxMaxZ,
                    other.SleeveBoundingBoxMinX, other.SleeveBoundingBoxMinY, other.SleeveBoundingBoxMinZ,
                    other.SleeveBoundingBoxMaxX, other.SleeveBoundingBoxMaxY, other.SleeveBoundingBoxMaxZ);
            }
            
            return minDistance <= toleranceDistance;
        }
        
        /// <summary>
        /// Calculate minimum distance between two 2D rectangles
        /// Returns 0 if they overlap, otherwise the shortest distance between any two points
        /// </summary>
        private double CalculateMinimumDistance2D(double minX1, double minY1, double maxX1, double maxY1,
                                                double minX2, double minY2, double maxX2, double maxY2)
        {
            // Check if rectangles overlap
            bool xOverlap = maxX1 >= minX2 && minX1 <= maxX2;
            bool yOverlap = maxY1 >= minY2 && minY1 <= maxY2;
            
            // Debug logging
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect1: Min=({minX1:F3}, {minY1:F3}), Max=({maxX1:F3}, {maxY1:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] Rect2: Min=({minX2:F3}, {minY2:F3}), Max=({maxX2:F3}, {maxY2:F3})\n");
            DebugLogger.Info($"[DISTANCE-DEBUG] X-overlap: {xOverlap}, Y-overlap: {yOverlap}\n");
            
            if (xOverlap && yOverlap)
            {
                DebugLogger.Info($"[DISTANCE-DEBUG] Result: OVERLAP (distance = 0)\n");
                return 0.0; // Rectangles overlap
            }
            
            // Calculate minimum distance between non-overlapping rectangles
            double dx = Math.Max(0, Math.Max(minX1 - maxX2, minX2 - maxX1));
            double dy = Math.Max(0, Math.Max(minY1 - maxY2, minY2 - maxY1));
            double distance = Math.Sqrt(dx * dx + dy * dy);
            
            DebugLogger.Info($"[DISTANCE-DEBUG] dx={dx:F3}, dy={dy:F3}, distance={distance:F3} feet ({distance * 304.8:F1}mm)\n");
            
            return distance;
        }
        
        /// <summary>
        /// Calculate minimum distance between two 3D bounding boxes
        /// Returns 0 if they overlap, otherwise the shortest distance between any two points
        /// </summary>
        private double CalculateMinimumDistance3D(double minX1, double minY1, double minZ1, double maxX1, double maxY1, double maxZ1,
                                                  double minX2, double minY2, double minZ2, double maxX2, double maxY2, double maxZ2)
        {
            // Check if bounding boxes overlap
            bool xOverlap = maxX1 >= minX2 && minX1 <= maxX2;
            bool yOverlap = maxY1 >= minY2 && minY1 <= maxY2;
            bool zOverlap = maxZ1 >= minZ2 && minZ1 <= maxZ2;
            
            if (xOverlap && yOverlap && zOverlap)
            {
                return 0.0; // Bounding boxes overlap
            }
            
            // Calculate minimum distance between non-overlapping bounding boxes
            double dx = Math.Max(0, Math.Max(minX1 - maxX2, minX2 - maxX1));
            double dy = Math.Max(0, Math.Max(minY1 - maxY2, minY2 - maxY1));
            double dz = Math.Max(0, Math.Max(minZ1 - maxZ2, minZ2 - maxZ1));
            
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
    }
    
    /// <summary>
    /// ✅ TREE STRUCTURE: FileComboGroup for Filter XML (groups clash zones by file combo)
    /// Similar to Global XML's FileComboGroup but contains full ClashZone objects instead of just entries
    /// </summary>
    public class FilterFileComboGroup
    {
        [XmlAttribute("LinkedFile")] public string LinkedFile { get; set; } = string.Empty;
        [XmlAttribute("HostFile")] public string HostFile { get; set; } = string.Empty;
        [XmlAttribute("ProcessedAt")] public DateTime ProcessedAt { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Clash zones for this file combo
        /// </summary>
        [XmlArray("ClashZones")]
        [XmlArrayItem("ClashZone")]
        public List<ClashZone> ClashZones { get; set; } = new List<ClashZone>();
        
        /// <summary>
        /// Creates a normalized key for comparison (case-insensitive, removes path info)
        /// ✅ FIX: Handles old format ": X : location Shared" pattern
        /// </summary>
        public string GetNormalizedKey()
        {
            Func<string, string> norm = s =>
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                var trimmed = s;
                
                // ✅ FIX: Remove old format ": X : location Shared" pattern (e.g., ": 12 : location Shared")
                // This handles legacy file combo names from Global XML
                var locationMatch = System.Text.RegularExpressions.Regex.Match(trimmed, @":\s*\d+\s*:\s*location\s+Shared", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (locationMatch.Success)
                {
                    trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
                }
                
                var idxParen = trimmed.IndexOf('(');
                if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
                trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                return trimmed.Trim();
            };
            
            return $"{norm(LinkedFile)}|{norm(HostFile)}";
        }
    }
    
    /// <summary>
    /// ✅ TREE STRUCTURE: FilterGroup for Filter XML (groups file combos by filter)
    /// Similar to Global XML's FilterGroup but contains full ClashZone objects instead of just entries
    /// </summary>
    public class FilterGroupForStorage
    {
        [XmlAttribute("Name")] public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// File combo groups for this filter
        /// </summary>
        [XmlElement("FileCombo")]
        public List<FilterFileComboGroup> FileCombos { get; set; } = new List<FilterFileComboGroup>();
    }

    /// <summary>
    /// Container for storing clash zones in a profile
    /// ✅ TREE STRUCTURE: Same hierarchical structure as Global XML - Filter → FileCombo → ClashZones
    /// </summary>
    [XmlRoot("ClashZoneStorage")]
    public class ClashZoneStorage
    {
        /// <summary>
        /// ✅ TREE STRUCTURE: Hierarchical organization - Filter → FileCombo → ClashZones
        /// This is the PRIMARY structure for Filter XML (same as Global XML)
        /// </summary>
        [XmlArray("Filters")]
        [XmlArrayItem("Filter")]
        public List<FilterGroupForStorage> Filters { get; set; } = new List<FilterGroupForStorage>();
        
        /// <summary>
        /// ⚠️ DEPRECATED: Flat structure - kept for backward compatibility during migration
        /// Will be migrated to hierarchical structure (Filters) on first load
        /// </summary>
        [XmlArray("ClashZones")]
        [XmlArrayItem("ClashZone")]
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

        /// <summary>
        /// Returns all clash zones in this storage using the primary tree structure,
        /// falling back to the deprecated flat list for backward compatibility.
        /// </summary>
        [XmlIgnore]
        public List<ClashZone> AllZones
        {
            get
            {
                var result = new List<ClashZone>();
                var seen = new HashSet<Guid>();

                if (Filters != null)
                {
                    foreach (var filterGroup in Filters)
                    {
                        if (filterGroup?.FileCombos == null) continue;

                        foreach (var fileCombo in filterGroup.FileCombos)
                        {
                            if (fileCombo?.ClashZones == null) continue;

                            foreach (var cz in fileCombo.ClashZones)
                            {
                                if (cz == null) continue;
                                if (seen.Add(cz.Id))
                                {
                                    result.Add(cz);
                                }
                            }
                        }
                    }
                }

                if (ClashZones != null && ClashZones.Count > 0)
                {
                    foreach (var cz in ClashZones)
                    {
                        if (cz == null) continue;
                        if (seen.Add(cz.Id))
                        {
                            result.Add(cz);
                        }
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// Helper method to enumerate all clash zones.
        /// </summary>
        public IEnumerable<ClashZone> EnumerateAllZones()
        {
            return AllZones;
        }
    }
}