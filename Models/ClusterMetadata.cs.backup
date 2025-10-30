using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Metadata for a cluster opening - tracks all MEP elements contained within
    /// Used for generating opening schedules and coordination
    /// </summary>
    [XmlRoot("ClusterMetadata")]
    public class ClusterMetadata
    {
        /// <summary>
        /// Unique cluster identifier (Revit Element ID)
        /// </summary>
        public int ClusterId { get; set; }
        
        /// <summary>
        /// Cluster mark parameter (user-facing identifier like "DC-001")
        /// </summary>
        public string ClusterMark { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster family name (ClusterOpeningOnWallX, ClusterOpeningOnWallY, ClusterOpeningOnSlab)
        /// </summary>
        public string ClusterFamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// MEP Category for this cluster (ALL elements must be same category)
        /// Values: "Ducts", "Pipes", "Cable Trays", "Duct Accessories"
        /// </summary>
        public string MepCategory { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element type (Wall, Floor, Framing)
        /// </summary>
        public string HostType { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element ID
        /// </summary>
        public int HostElementId { get; set; }
        
        /// <summary>
        /// Host element name/type
        /// </summary>
        public string HostName { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster center point (X, Y, Z in feet)
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
        
        /// <summary>
        /// Cluster bounding box width (in mm)
        /// </summary>
        public double ClusterWidthMm { get; set; }
        
        /// <summary>
        /// Cluster bounding box height (in mm)
        /// </summary>
        public double ClusterHeightMm { get; set; }
        
        /// <summary>
        /// Level name where cluster is located
        /// </summary>
        public string LevelName { get; set; } = string.Empty;
        
        /// <summary>
        /// When this cluster was created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// List of MEP elements contained in this cluster (ALL same category)
        /// </summary>
        [XmlArray("MepElements")]
        [XmlArrayItem("MepElement")]
        public List<MepElementInfo> MepElements { get; set; } = new List<MepElementInfo>();
        
        /// <summary>
        /// List of individual sleeve IDs that were replaced by this cluster
        /// </summary>
        [XmlArray("ReplacedSleeves")]
        [XmlArrayItem("SleeveId")]
        public List<int> ReplacedSleeveIds { get; set; } = new List<int>();
        
        /// <summary>
        /// Cluster tolerance used (JoinOpeningsDistance in mm)
        /// </summary>
        public double ClusterToleranceMm { get; set; }
        
        /// <summary>
        /// Number of MEP elements in this cluster
        /// </summary>
        [XmlIgnore]
        public int MepElementCount => MepElements?.Count ?? 0;
        
        /// <summary>
        /// Total cross-sectional area of all MEP elements (in mm²)
        /// </summary>
        [XmlIgnore]
        public double TotalMepAreaMm2
        {
            get
            {
                double total = 0;
                if (MepElements != null)
                {
                    foreach (var mep in MepElements)
                    {
                        total += mep.CrossSectionalAreaMm2;
                    }
                }
                return total;
            }
        }
        
        /// <summary>
        /// Cluster opening area (in mm²)
        /// </summary>
        [XmlIgnore]
        public double ClusterAreaMm2 => ClusterWidthMm * ClusterHeightMm;
        
        /// <summary>
        /// Percentage of cluster area occupied by MEP elements
        /// </summary>
        [XmlIgnore]
        public double OccupancyPercentage => ClusterAreaMm2 > 0 ? (TotalMepAreaMm2 / ClusterAreaMm2) * 100.0 : 0.0;
    }
    
    /// <summary>
    /// Information about a single MEP element within a cluster
    /// </summary>
    public class MepElementInfo
    {
        /// <summary>
        /// MEP element ID
        /// </summary>
        public int ElementId { get; set; }
        
        /// <summary>
        /// MEP element unique ID (persistent across sessions)
        /// </summary>
        public string UniqueId { get; set; } = string.Empty;
        
        /// <summary>
        /// MEP category (Ducts, Pipes, Cable Trays, Duct Accessories)
        /// </summary>
        public string Category { get; set; } = string.Empty;
        
        /// <summary>
        /// System abbreviation (HVAC, PLB, ELEC, etc.)
        /// </summary>
        public string SystemAbbreviation { get; set; } = string.Empty;
        
        /// <summary>
        /// System name (Supply Air, Domestic Hot Water, etc.)
        /// </summary>
        public string SystemName { get; set; } = string.Empty;
        
        /// <summary>
        /// System type (Supply Air, Return Air, Exhaust Air, etc.)
        /// </summary>
        public string SystemType { get; set; } = string.Empty;
        
        /// <summary>
        /// Element shape (Round, Rectangular, Oval)
        /// </summary>
        public string Shape { get; set; } = string.Empty;
        
        /// <summary>
        /// Width or Diameter in mm
        /// </summary>
        public double WidthMm { get; set; }
        
        /// <summary>
        /// Height in mm (for rectangular elements)
        /// </summary>
        public double HeightMm { get; set; }
        
        /// <summary>
        /// Diameter in mm (for round elements)
        /// </summary>
        public double DiameterMm { get; set; }
        
        /// <summary>
        /// Size formatted as string (e.g., "Ø300", "400×200")
        /// </summary>
        [XmlIgnore]
        public string SizeFormatted
        {
            get
            {
                if (Shape == "Round" || Shape == "Circular")
                    return $"Ø{DiameterMm:F0}";
                else if (Shape == "Rectangular")
                    return $"{WidthMm:F0}×{HeightMm:F0}";
                else
                    return $"{WidthMm:F0}";
            }
        }
        
        /// <summary>
        /// Cross-sectional area in mm²
        /// </summary>
        public double CrossSectionalAreaMm2 { get; set; }
        
        /// <summary>
        /// Insulation type (None, Normal, Enhanced)
        /// </summary>
        public string InsulationType { get; set; } = "None";
        
        /// <summary>
        /// Insulation thickness in mm
        /// </summary>
        public double InsulationThicknessMm { get; set; }
        
        /// <summary>
        /// Family name
        /// </summary>
        public string FamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Family type name
        /// </summary>
        public string FamilyTypeName { get; set; } = string.Empty;
        
        /// <summary>
        /// Comments from MEP element
        /// </summary>
        public string Comments { get; set; } = string.Empty;
        
        /// <summary>
        /// Location of this MEP element
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
    }
    
    /// <summary>
    /// Simple XYZ point for serialization
    /// </summary>
    public class XyzPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        
        public override string ToString()
        {
            return $"({X:F2}, {Y:F2}, {Z:F2})";
        }
    }
    
    /// <summary>
    /// Container for all cluster metadata
    /// Used for generating opening schedules
    /// </summary>
    [XmlRoot("ClusterMetadataCollection")]
    public class ClusterMetadataCollection
    {
        [XmlArray("Clusters")]
        [XmlArrayItem("Cluster")]
        public List<ClusterMetadata> Clusters { get; set; } = new List<ClusterMetadata>();
        
        /// <summary>
        /// When this collection was created/updated
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Project name
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;
        
        /// <summary>
        /// Filter name that created these clusters
        /// </summary>
        public string FilterName { get; set; } = string.Empty;
        
        /// <summary>
        /// MEP category for these clusters (Ducts, Pipes, Cable Trays, Duct Accessories)
        /// </summary>
        public string Category { get; set; } = string.Empty;
        
        /// <summary>
        /// Total number of clusters
        /// </summary>
        [XmlIgnore]
        public int TotalClusters => Clusters?.Count ?? 0;
        
        /// <summary>
        /// Total number of MEP elements across all clusters
        /// </summary>
        [XmlIgnore]
        public int TotalMepElements
        {
            get
            {
                int total = 0;
                if (Clusters != null)
                {
                    foreach (var cluster in Clusters)
                    {
                        total += cluster.MepElementCount;
                    }
                }
                return total;
            }
        }
    }
}
