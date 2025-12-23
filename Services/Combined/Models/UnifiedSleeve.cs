using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models
{
    /// <summary>
    /// Unified abstraction for both individual and cluster sleeves.
    /// Enables uniform proximity detection logic regardless of sleeve type.
    /// 
    /// This abstraction allows Agent B's proximity detection algorithm to treat
    /// individual sleeves (ClashZones) and cluster sleeves uniformly without
    /// needing separate logic for each type.
    /// </summary>
    public class UnifiedSleeve
    {
        // ============================================================================
        // IDENTIFIERS
        // ============================================================================
        
        /// <summary>
        /// Unique identifier: ClashZoneGuid (for individual) or ClusterInstanceId (for cluster)
        /// </summary>
        public string Id { get; set; }
        
        /// <summary>
        /// Type of sleeve: Individual or Cluster
        /// </summary>
        public SleeveType Type { get; set; }
        
        /// <summary>
        /// MEP category: 'Pipes', 'Ducts', 'Cable Trays', 'Conduits'
        /// </summary>
        public string Category { get; set; }
        
        // ============================================================================
        // GEOMETRY
        // ============================================================================
        
        /// <summary>
        /// Bounding box in world coordinates
        /// </summary>
        public BoundingBoxXYZ BoundingBox { get; set; }
        
        /// <summary>
        /// Placement point (center of sleeve)
        /// </summary>
        public XYZ PlacementPoint { get; set; }
        
        /// <summary>
        /// Four corner coordinates in world space
        /// Used for accurate proximity calculations and combined bounding box
        /// </summary>
        public List<XYZ> Corners { get; set; }
        
        // ============================================================================
        // METADATA
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
        /// Rotation angle in degrees
        /// </summary>
        public double RotationAngleDeg { get; set; }
        
        // ============================================================================
        // SOURCE DATA
        // ============================================================================
        
        /// <summary>
        /// Reference to original source data (ClashZone or ClusterSleeve object)
        /// Allows retrieval of full data after proximity detection
        /// </summary>
        public object SourceData { get; set; }
        
        // ============================================================================
        // HELPER METHODS
        // ============================================================================
        
        /// <summary>
        /// Gets the center point of the sleeve
        /// </summary>
        public XYZ GetCenter()
        {
            if (PlacementPoint != null)
                return PlacementPoint;
            
            if (BoundingBox != null)
                return (BoundingBox.Min + BoundingBox.Max) / 2.0;
            
            throw new InvalidOperationException("UnifiedSleeve has no placement point or bounding box");
        }
        
        /// <summary>
        /// Calculates distance to another sleeve (center-to-center)
        /// </summary>
        public double GetDistanceTo(UnifiedSleeve other)
        {
            if (other == null)
                throw new ArgumentNullException(nameof(other));
            
            return this.GetCenter().DistanceTo(other.GetCenter());
        }
        
        /// <summary>
        /// Checks if this sleeve is within proximity threshold of another sleeve
        /// </summary>
        public bool IsWithinProximity(UnifiedSleeve other, double threshold)
        {
            if (other == null)
                throw new ArgumentNullException(nameof(other));
            
            return GetDistanceTo(other) <= threshold;
        }
        
        /// <summary>
        /// Checks if this sleeve's bounding box intersects with another sleeve's bounding box
        /// (with optional expansion by threshold)
        /// </summary>
        public bool BoundingBoxIntersects(UnifiedSleeve other, double expansionThreshold = 0.0)
        {
            if (other == null || this.BoundingBox == null || other.BoundingBox == null)
                return false;
            
            var thisMin = this.BoundingBox.Min - new XYZ(expansionThreshold, expansionThreshold, expansionThreshold);
            var thisMax = this.BoundingBox.Max + new XYZ(expansionThreshold, expansionThreshold, expansionThreshold);
            var otherMin = other.BoundingBox.Min;
            var otherMax = other.BoundingBox.Max;
            
            return thisMin.X <= otherMax.X && thisMax.X >= otherMin.X &&
                   thisMin.Y <= otherMax.Y && thisMax.Y >= otherMin.Y &&
                   thisMin.Z <= otherMax.Z && thisMax.Z >= otherMin.Z;
        }
        
        /// <summary>
        /// Creates a UnifiedSleeve from a ClashZone (individual sleeve)
        /// </summary>
        public static UnifiedSleeve FromClashZone(JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone clashZone)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
            
            var corners = new List<XYZ>
            {
                new XYZ(clashZone.SleeveCorner1X ?? 0.0, clashZone.SleeveCorner1Y ?? 0.0, clashZone.SleeveCorner1Z ?? 0.0),
                new XYZ(clashZone.SleeveCorner2X ?? 0.0, clashZone.SleeveCorner2Y ?? 0.0, clashZone.SleeveCorner2Z ?? 0.0),
                new XYZ(clashZone.SleeveCorner3X ?? 0.0, clashZone.SleeveCorner3Y ?? 0.0, clashZone.SleeveCorner3Z ?? 0.0),
                new XYZ(clashZone.SleeveCorner4X ?? 0.0, clashZone.SleeveCorner4Y ?? 0.0, clashZone.SleeveCorner4Z ?? 0.0)
            };
            
            var bbox = new BoundingBoxXYZ();
            
            // Check if stored BBox is valid (volume > 0 or at least dimensions > 0)
            bool hasValidStoredBBox = 
                clashZone.SleeveBoundingBoxMaxX > clashZone.SleeveBoundingBoxMinX ||
                clashZone.SleeveBoundingBoxMaxY > clashZone.SleeveBoundingBoxMinY;

            if (hasValidStoredBBox)
            {
                bbox.Min = new XYZ(
                    clashZone.SleeveBoundingBoxMinX, 
                    clashZone.SleeveBoundingBoxMinY, 
                    clashZone.SleeveBoundingBoxMinZ);
                bbox.Max = new XYZ(
                    clashZone.SleeveBoundingBoxMaxX, 
                    clashZone.SleeveBoundingBoxMaxY, 
                    clashZone.SleeveBoundingBoxMaxZ);
            }
            else
            {
                // ✅ CRITICAL FIX: Derive BBox from Corners if stored BBox is missing (e.g. Synthetic Zones)
                double minX = corners.Min(c => c.X);
                double minY = corners.Min(c => c.Y);
                double minZ = corners.Min(c => c.Z);
                double maxX = corners.Max(c => c.X);
                double maxY = corners.Max(c => c.Y);
                double maxZ = corners.Max(c => c.Z);

                // Handle 2D case (flat Z) by adding height if available
                if (Math.Abs(maxZ - minZ) < 0.001)
                {
                     double height = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : clashZone.SleeveDiameter;
                     if (height > 0) maxZ += height;
                }

                bbox.Min = new XYZ(minX, minY, minZ);
                bbox.Max = new XYZ(maxX, maxY, maxZ);
            }
            
            return new UnifiedSleeve
            {
                Id = clashZone.Id.ToString(),
                Type = SleeveType.Individual,
                Category = clashZone.MepElementCategory,
                BoundingBox = bbox,
                PlacementPoint = clashZone.IntersectionPoint, // Use IntersectionPoint as placement point
                Corners = corners,
                HostType = clashZone.StructuralElementType,
                HostOrientation = clashZone.HostOrientation,
                RotationAngleDeg = clashZone.MepElementRotationAngle * (180.0 / Math.PI), // Convert Rad to Deg
                SourceData = clashZone
            };
        }
        
        /// <summary>
        /// Creates a UnifiedSleeve from a ClusterSleeveData (cluster sleeve)
        /// </summary>
        public static UnifiedSleeve FromClusterSleeve(JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData clusterSleeve)
        {
            if (clusterSleeve == null)
                throw new ArgumentNullException(nameof(clusterSleeve));
            
            // Populate corners from ClusterSleeveData (corner columns now exist in DB)
            var corners = new List<XYZ>
            {
                new XYZ(clusterSleeve.Corner1X, clusterSleeve.Corner1Y, clusterSleeve.Corner1Z),
                new XYZ(clusterSleeve.Corner2X, clusterSleeve.Corner2Y, clusterSleeve.Corner2Z),
                new XYZ(clusterSleeve.Corner3X, clusterSleeve.Corner3Y, clusterSleeve.Corner3Z),
                new XYZ(clusterSleeve.Corner4X, clusterSleeve.Corner4Y, clusterSleeve.Corner4Z)
            };
            
            var bbox = new BoundingBoxXYZ
            {
                Min = new XYZ(clusterSleeve.BoundingBoxMinX, clusterSleeve.BoundingBoxMinY, clusterSleeve.BoundingBoxMinZ),
                Max = new XYZ(clusterSleeve.BoundingBoxMaxX, clusterSleeve.BoundingBoxMaxY, clusterSleeve.BoundingBoxMaxZ)
            };
            
            var placementPoint = new XYZ(clusterSleeve.PlacementX, clusterSleeve.PlacementY, clusterSleeve.PlacementZ);
            
            return new UnifiedSleeve
            {
                Id = clusterSleeve.ClusterInstanceId.ToString(),
                Type = SleeveType.Cluster,
                Category = clusterSleeve.Category,
                BoundingBox = bbox,
                PlacementPoint = placementPoint,
                Corners = corners,
                HostType = clusterSleeve.HostType,
                HostOrientation = clusterSleeve.HostOrientation,
                RotationAngleDeg = clusterSleeve.RotationAngleDeg,
                SourceData = clusterSleeve
            };
        }
    }
    
    /// <summary>
    /// Enum representing the type of sleeve
    /// </summary>
    public enum SleeveType
    {
        /// <summary>
        /// Individual sleeve (single clash zone)
        /// </summary>
        Individual,
        
        /// <summary>
        /// Cluster sleeve (multiple clash zones grouped together)
        /// </summary>
        Cluster
    }
}
