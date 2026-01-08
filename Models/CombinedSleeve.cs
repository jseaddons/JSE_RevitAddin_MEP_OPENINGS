using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a cross-category combined sleeve that encompasses multiple
    /// individual and/or cluster sleeves from different MEP categories.
    /// 
    /// Example: A combined sleeve containing 2 pipe sleeves and 1 duct sleeve
    /// </summary>
    public class CombinedSleeve
    {
        // ============================================================================
        // PRIMARY IDENTIFIERS
        // ============================================================================
        
        /// <summary>
        /// Database primary key (auto-increment)
        /// </summary>
        public int CombinedSleeveId { get; set; }
        
        /// <summary>
        /// Revit element ID (ElementId.IntegerValue)
        /// </summary>
        public int CombinedInstanceId { get; set; }
        
        /// <summary>
        /// Deterministic GUID generated from constituent sleeve IDs
        /// Used to prevent duplicate combined sleeve rows in database
        /// </summary>
        public string DeterministicGuid { get; set; }
        
        // ============================================================================
        // FOREIGN KEYS
        // ============================================================================
        
        /// <summary>
        /// File combination ID (links to FileCombos table)
        /// </summary>
        public int ComboId { get; set; }
        
        /// <summary>
        /// Filter ID (links to Filters table)
        /// </summary>
        public int FilterId { get; set; }
        
        // ============================================================================
        // CATEGORIES
        // ============================================================================
        
        /// <summary>
        /// List of MEP categories included in this combined sleeve
        /// Example: ["Pipes", "Ducts"] or ["Pipes", "Ducts", "Cable Trays"]
        /// </summary>
        public List<string> Categories { get; set; } = new List<string>();
        
        // ============================================================================
        // GEOMETRY - BOUNDING BOX
        // ============================================================================
        
        /// <summary>
        /// Bounding box minimum X coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMinX { get; set; }
        
        /// <summary>
        /// Bounding box minimum Y coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMinY { get; set; }
        
        /// <summary>
        /// Bounding box minimum Z coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMinZ { get; set; }
        
        /// <summary>
        /// Bounding box maximum X coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMaxX { get; set; }
        
        /// <summary>
        /// Bounding box maximum Y coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMaxY { get; set; }
        
        /// <summary>
        /// Bounding box maximum Z coordinate (world space, feet)
        /// </summary>
        public double BoundingBoxMaxZ { get; set; }
        
        // ============================================================================
        // GEOMETRY - DIMENSIONS
        // ============================================================================
        
        /// <summary>
        /// Combined width (X-axis dimension, feet)
        /// </summary>
        public double CombinedWidth { get; set; }
        
        /// <summary>
        /// Combined height (Y-axis dimension, feet)
        /// </summary>
        public double CombinedHeight { get; set; }
        
        /// <summary>
        /// Combined depth (Z-axis dimension, feet)
        /// </summary>
        public double CombinedDepth { get; set; }
        
        // ============================================================================
        // PLACEMENT
        // ============================================================================
        
        /// <summary>
        /// Placement point X coordinate (world space, feet)
        /// </summary>
        public double PlacementX { get; set; }
        
        /// <summary>
        /// Placement point Y coordinate (world space, feet)
        /// </summary>
        public double PlacementY { get; set; }
        
        /// <summary>
        /// Placement point Z coordinate (world space, feet)
        /// </summary>
        public double PlacementZ { get; set; }
        
        /// <summary>
        /// Rotation angle in degrees (0-360)
        /// </summary>
        public double RotationAngleDeg { get; set; }
        
        // ============================================================================
        // HOST INFORMATION
        // ============================================================================
        
        /// <summary>
        /// Host element type: 'Wall', 'Floor', 'Roof', 'Framing'
        /// </summary>
        public string HostType { get; set; }
        
        /// <summary>
        /// Host orientation: 'X', 'Y', 'Z'
        /// </summary>
        public string HostOrientation { get; set; }

        /// <summary>
        /// Elevation from Level (internally calculated or from Revit param)
        /// Used for reliable Bottom of Opening calculation.
        /// </summary>
        public double? ElevationFromLevel { get; set; }
        
        // ============================================================================
        // CORNER COORDINATES (World Space)
        // ============================================================================
        // Calculated after placement, stored for downstream processes
        // (parameter transfer, visualization, etc.)
        // ============================================================================
        
        /// <summary>Corner 1 X coordinate (world space, feet)</summary>
        public double Corner1X { get; set; }
        
        /// <summary>Corner 1 Y coordinate (world space, feet)</summary>
        public double Corner1Y { get; set; }
        
        /// <summary>Corner 1 Z coordinate (world space, feet)</summary>
        public double Corner1Z { get; set; }
        
        /// <summary>Corner 2 X coordinate (world space, feet)</summary>
        public double Corner2X { get; set; }
        
        /// <summary>Corner 2 Y coordinate (world space, feet)</summary>
        public double Corner2Y { get; set; }
        
        /// <summary>Corner 2 Z coordinate (world space, feet)</summary>
        public double Corner2Z { get; set; }
        
        /// <summary>Corner 3 X coordinate (world space, feet)</summary>
        public double Corner3X { get; set; }
        
        /// <summary>Corner 3 Y coordinate (world space, feet)</summary>
        public double Corner3Y { get; set; }
        
        /// <summary>Corner 3 Z coordinate (world space, feet)</summary>
        public double Corner3Z { get; set; }
        
        /// <summary>Corner 4 X coordinate (world space, feet)</summary>
        public double Corner4X { get; set; }
        
        /// <summary>Corner 4 Y coordinate (world space, feet)</summary>
        public double Corner4Y { get; set; }
        
        /// <summary>Corner 4 Z coordinate (world space, feet)</summary>
        public double Corner4Z { get; set; }
        
        // ============================================================================
        // TIMESTAMPS
        // ============================================================================
        
        /// <summary>
        /// Timestamp when this combined sleeve was created
        /// </summary>
        public DateTime CreatedAt { get; set; }
        
        /// <summary>
        /// Timestamp when this combined sleeve was last updated
        /// </summary>
        public DateTime UpdatedAt { get; set; }

        /// <summary>
        /// The Family Name of the placed sleeve (e.g. "Rectangular Sleeve", "Round Sleeve")
        /// </summary>
        public string SleeveFamilyName { get; set; }
        
        // ============================================================================
        // NAVIGATION PROPERTIES
        // ============================================================================
        
        /// <summary>
        /// List of constituent sleeves (individual or cluster) that make up this combined sleeve
        /// </summary>
        public List<SleeveConstituent> Constituents { get; set; } = new List<SleeveConstituent>();
    }
}
