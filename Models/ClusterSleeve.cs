using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a cluster sleeve record from the database.
    /// Used for retrieving persisted corner data for manual join operations.
    /// </summary>
    public class ClusterSleeve
    {
        public int ClusterSleeveId { get; set; }
        public long ClusterInstanceId { get; set; }
        public string Category { get; set; } = string.Empty;
        
        // Corners 1-4 X/Y/Z
        public double? Corner1X { get; set; }
        public double? Corner1Y { get; set; }
        public double? Corner1Z { get; set; }
        
        public double? Corner2X { get; set; }
        public double? Corner2Y { get; set; }
        public double? Corner2Z { get; set; }
        
        public double? Corner3X { get; set; }
        public double? Corner3Y { get; set; }
        public double? Corner3Z { get; set; }
        
        public double? Corner4X { get; set; }
        public double? Corner4Y { get; set; }
        public double? Corner4Z { get; set; }
        
        public double? RotationAngleDeg { get; set; }
        
        // Host Info
        public string HostType { get; set; }
        public string HostOrientation { get; set; }
    }
}
