using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services;

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
        public static UnifiedSleeve FromClashZone(ClashZone clashZone)
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

            bool isClusterResolvedZone = clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0;
            if (isClusterResolvedZone)
            {
                double c1x = clashZone.SleeveCorner1X ?? 0.0;
                bool cornersSet = Math.Abs(c1x) > 1e-6 || Math.Abs(clashZone.SleeveCorner2X ?? 0.0) > 1e-6;
                DebugLogger.Info($"[FromClashZone] cluster-resolved ZoneId={clashZone.Id} ClusterInstanceId={clashZone.ClusterSleeveInstanceId} CornersSet={cornersSet} (C1X={c1x:F3})");
            }

            // ✅ FIX: Extrude Corners for Floor Sleeves (Flat Z) to ensure ProximityGroup has 3D volume
            // Without this, rangeZ is 0, leading to Depth=0 for combined floor sleeves.
            if (string.Equals(clashZone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                double minZ = corners.Min(c => c.Z);
                double maxZ = corners.Max(c => c.Z);
                
                if (Math.Abs(maxZ - minZ) < 0.001)
                {
                     double height = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : 
                                    (clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : 1.0);
                     
                     // Add 4 extruded corners (Top Face)
                     var topCorners = corners.Select(c => new XYZ(c.X, c.Y, c.Z + height)).ToList();
                     corners.AddRange(topCorners);
                }
            }
            
            var bbox = new BoundingBoxXYZ();
            string sleeveId = clashZone.Id.ToString();
            var sleeveType = SleeveType.Individual;
            XYZ placementPoint = clashZone.IntersectionPoint;

            // ✅ CLUSTER-RESOLVED: Use cluster bounding box and cluster id for proximity (one sleeve per cluster)
            bool isClusterResolved = clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0;
            bool hasValidClusterBBox = clashZone.ClusterSleeveBoundingBoxMaxX > clashZone.ClusterSleeveBoundingBoxMinX ||
                                       clashZone.ClusterSleeveBoundingBoxMaxY > clashZone.ClusterSleeveBoundingBoxMinY;

            if (isClusterResolved && hasValidClusterBBox)
            {
                bbox.Min = new XYZ(
                    clashZone.ClusterSleeveBoundingBoxMinX,
                    clashZone.ClusterSleeveBoundingBoxMinY,
                    clashZone.ClusterSleeveBoundingBoxMinZ);
                bbox.Max = new XYZ(
                    clashZone.ClusterSleeveBoundingBoxMaxX,
                    clashZone.ClusterSleeveBoundingBoxMaxY,
                    clashZone.ClusterSleeveBoundingBoxMaxZ);
                sleeveId = clashZone.ClusterSleeveInstanceId.ToString();
                sleeveType = SleeveType.Cluster;
                placementPoint = new XYZ(
                    (bbox.Min.X + bbox.Max.X) * 0.5,
                    (bbox.Min.Y + bbox.Max.Y) * 0.5,
                    (bbox.Min.Z + bbox.Max.Z) * 0.5);

                if (Math.Abs(bbox.Max.Z - bbox.Min.Z) < 0.001)
                {
                    double height = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight :
                                   (clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : 1.0);
                    bbox.Max = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + height);
                    placementPoint = new XYZ(placementPoint.X, placementPoint.Y, placementPoint.Z + height * 0.5);
                }
            }
            else
            {
                // Individual sleeve: use stored BBox or derive from corners
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

                    if (Math.Abs(bbox.Max.Z - bbox.Min.Z) < 0.001)
                    {
                        double height = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight :
                                       (clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : 1.0);
                        bbox.Max = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + height);
                    }
                }
                else
                {
                    double minX = corners.Min(c => c.X);
                    double minY = corners.Min(c => c.Y);
                    double minZ = corners.Min(c => c.Z);
                    double maxX = corners.Max(c => c.X);
                    double maxY = corners.Max(c => c.Y);
                    double maxZ = corners.Max(c => c.Z);

                    if (Math.Abs(maxZ - minZ) < 0.001)
                    {
                        double height = clashZone.SleeveHeight > 0 ? clashZone.SleeveHeight : clashZone.SleeveDiameter;
                        if (height > 0) maxZ += height;
                    }

                    bbox.Min = new XYZ(minX, minY, minZ);
                    bbox.Max = new XYZ(maxX, maxY, maxZ);
                }
            }

            return new UnifiedSleeve
            {
                Id = sleeveId,
                Type = sleeveType,
                Category = clashZone.MepElementCategory,
                BoundingBox = bbox,
                PlacementPoint = placementPoint,
                Corners = corners,
                HostType = clashZone.StructuralElementType,
                HostOrientation = clashZone.HostOrientation,
                RotationAngleDeg = clashZone.MepElementRotationAngle * (180.0 / Math.PI),
                SourceData = clashZone
            };
        }
        
        /// <summary>
        /// Creates a UnifiedSleeve from a ClusterSleeveData (cluster sleeve)
        /// </summary>
        public static UnifiedSleeve FromClusterSleeve(ClusterSleeveData clusterSleeve)
        {
            if (clusterSleeve == null)
                throw new ArgumentNullException(nameof(clusterSleeve));
            
            // Populate corners from ClusterSleeveData (corner columns now exist in DB)
            var corners = new List<XYZ>
            {
                new XYZ(clusterSleeve.Corner1X, clusterSleeve.Corner1Y, clusterSleeve.Corner1Z),
                new XYZ(clusterSleeve.Corner2X, clusterSleeve.Corner2Y, clusterSleeve.Corner2Z),
                new XYZ(clusterSleeve.Corner3X, clusterSleeve.Corner3Y, clusterSleeve.Corner3Z),
                new XYZ(clusterSleeve.Corner3X, clusterSleeve.Corner3Y, clusterSleeve.Corner3Z),
                new XYZ(clusterSleeve.Corner4X, clusterSleeve.Corner4Y, clusterSleeve.Corner4Z)
            };

            // ✅ FIX: Extrude Corners for Floor Clusters (Flat Z)
            if (string.Equals(clusterSleeve.HostType, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                double minZ = corners.Min(c => c.Z);
                double maxZ = corners.Max(c => c.Z);

                if (Math.Abs(maxZ - minZ) < 0.001)
                {
                     double height = clusterSleeve.ClusterDepth > 0 ? clusterSleeve.ClusterDepth : 
                                    (clusterSleeve.ClusterHeight > 0 ? clusterSleeve.ClusterHeight : 1.0);
                     
                     var topCorners = corners.Select(c => new XYZ(c.X, c.Y, c.Z + height)).ToList();
                     corners.AddRange(topCorners);
                }
            }
            
            // ✅ IMPROVED LOGIC: Respect "High Quality" stored BBox if it has volume
            double dbMinZ = clusterSleeve.BoundingBoxMinZ;
            double dbMaxZ = clusterSleeve.BoundingBoxMaxZ;
            bool dbHasVolume = Math.Abs(dbMaxZ - dbMinZ) > 0.001;

            BoundingBoxXYZ bbox;

            if (dbHasVolume)
            {
                // Trusted Source: Database has valid 3D volume (e.g. User's Duct scenario)
                bbox = new BoundingBoxXYZ
                {
                    Min = new XYZ(clusterSleeve.BoundingBoxMinX, clusterSleeve.BoundingBoxMinY, clusterSleeve.BoundingBoxMinZ),
                    Max = new XYZ(clusterSleeve.BoundingBoxMaxX, clusterSleeve.BoundingBoxMaxY, clusterSleeve.BoundingBoxMaxZ)
                };
            }
            else
            {
                // Fallback: DB BBox is flat (Symbolic lines). Derive from Corners + Inflate.
                double minX = corners.Min(c => c.X);
                double minY = corners.Min(c => c.Y);
                double minZ = corners.Min(c => c.Z);
                double maxX = corners.Max(c => c.X);
                double maxY = corners.Max(c => c.Y);
                double maxZ = corners.Max(c => c.Z);

                // Derived is likely flat (since corners are planar), so we inflate
                double height = clusterSleeve.ClusterDepth > 0 ? clusterSleeve.ClusterDepth : 
                               (clusterSleeve.ClusterHeight > 0 ? clusterSleeve.ClusterHeight : 1.0); // 1.0 ft default
                
                maxZ += height; // Inflate upwards

                bbox = new BoundingBoxXYZ
                {
                    Min = new XYZ(minX, minY, minZ),
                    Max = new XYZ(maxX, maxY, maxZ)
                };
            }

            var placementPoint = new XYZ(clusterSleeve.PlacementX, clusterSleeve.PlacementY, clusterSleeve.PlacementZ);
            
            var result = new UnifiedSleeve
            {
                Id = clusterSleeve.ClusterInstanceId.ToString(),
                Type = SleeveType.Cluster,
                Category = clusterSleeve.Category,
                BoundingBox = bbox,
                PlacementPoint = placementPoint,
                Corners = corners,
                HostOrientation = clusterSleeve.HostOrientation,
                HostType = clusterSleeve.HostType, // ✅ FIX: Map HostType
                RotationAngleDeg = clusterSleeve.RotationAngleDeg,
                SourceData = clusterSleeve
            };

            return result;
        }

        /// <summary>
        /// Creates a UnifiedSleeve from a ClusterSleeve Model (used by Manual Join / ClashZoneRepository)
        /// </summary>
        public static UnifiedSleeve FromClusterSleeve(JSE_RevitAddin_MEP_OPENINGS.Models.ClusterSleeve clusterSleeve)
        {
            if (clusterSleeve == null)
                throw new ArgumentNullException(nameof(clusterSleeve));
            
            // Populate corners from ClusterSleeve Model
            // Handle nullables by coalescing to 0.0
            var corners = new List<XYZ>
            {
                new XYZ(clusterSleeve.Corner1X ?? 0.0, clusterSleeve.Corner1Y ?? 0.0, clusterSleeve.Corner1Z ?? 0.0),
                new XYZ(clusterSleeve.Corner2X ?? 0.0, clusterSleeve.Corner2Y ?? 0.0, clusterSleeve.Corner2Z ?? 0.0),
                new XYZ(clusterSleeve.Corner3X ?? 0.0, clusterSleeve.Corner3Y ?? 0.0, clusterSleeve.Corner3Z ?? 0.0),
                new XYZ(clusterSleeve.Corner3X ?? 0.0, clusterSleeve.Corner3Y ?? 0.0, clusterSleeve.Corner3Z ?? 0.0),
                new XYZ(clusterSleeve.Corner4X ?? 0.0, clusterSleeve.Corner4Y ?? 0.0, clusterSleeve.Corner4Z ?? 0.0)
            };

            // ✅ FIX: Extrude Corners for Floor Clusters (Flat Z)
            if (string.Equals(clusterSleeve.HostType, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                double localMinZ = corners.Min(c => c.Z);
                double localMaxZ = corners.Max(c => c.Z);

                if (Math.Abs(localMaxZ - localMinZ) < 0.001)
                {
                     double height = 1.0; // Default if not available
                     var topCorners = corners.Select(c => new XYZ(c.X, c.Y, c.Z + height)).ToList();
                     corners.AddRange(topCorners);
                }
            }
            
            // Model doesn't store BBox explicitly.
            // So we must derive BBox from Corners.
            
            double minX = corners.Min(c => c.X);
            double minY = corners.Min(c => c.Y);
            double minZ = corners.Min(c => c.Z);
            double maxX = corners.Max(c => c.X);
            double maxY = corners.Max(c => c.Y);
            double maxZ = corners.Max(c => c.Z);

            var bbox = new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
            
            // ✅ CRITICAL FIX: Inflate flat bounding boxes (e.g. Floors) for Cluster Sleeves (Model overload)
            double zDiff = bbox.Max.Z - bbox.Min.Z;
            if (Math.Abs(zDiff) < 0.001)
            {
                // Model doesn't have dimensions easily accessible here, use default 1.0 ft
                double height = 1.0; 
                var newMax = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z + height);
                bbox.Max = newMax;
            }

            // Global Center
            var placementPoint = (bbox.Min + bbox.Max) / 2.0;
            
            var result = new UnifiedSleeve
            {
                Id = clusterSleeve.ClusterInstanceId.ToString(),
                Type = SleeveType.Cluster,
                Category = clusterSleeve.Category,
                BoundingBox = bbox,
                PlacementPoint = placementPoint,
                Corners = corners,
                HostType = clusterSleeve.HostType,
                HostOrientation = clusterSleeve.HostOrientation,
                RotationAngleDeg = clusterSleeve.RotationAngleDeg ?? 0.0,
                SourceData = clusterSleeve
            };

            return result;
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
