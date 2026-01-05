using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data
{
    /// <summary>
    /// Data Transfer Object for clustering operations.
    /// Replaces anonymous/dynamic types for better performance and type safety.
    /// </summary>
    public class ClusteringSleeveDto
    {
        public int SleeveInstanceId { get; set; }
        public ClashZone ClashZone { get; set; }
        public FamilyInstance FamilyInstance { get; set; }
        public string HostType { get; set; }
        public string Orientation { get; set; }
        public string Category { get; set; }
        public string SystemType { get; set; }
        public bool IsCircular { get; set; }
        
        // Caching potentially expensive property lookups
        public double Width { get; set; }
        public double Height { get; set; }
        public double Diameter { get; set; }
        
        // For spatial grid optimization
        public XYZ LocationPoint { get; set; }
        public BoundingBoxXYZ BoundingBox { get; set; }
        
        // For rotated element support
        public double RotationAngle { get; set; }
        public bool IsRotated { get; set; }
        
        public ClusteringSleeveDto()
        {
        }
    }
}
