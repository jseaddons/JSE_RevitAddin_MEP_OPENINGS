using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models
{
    /// <summary>
    /// DTO for cluster sleeves feeding combined clustering (Phase 1-2, DB/CPU only).
    /// </summary>
    public class ClusterSleeveInfo
    {
        public ClusterSleeveInfo() 
        {
            ClashZoneIds = new List<Guid>();
        }

        public Guid ClashZoneId { get; set; }
        public List<Guid> ClashZoneIds { get; set; }
        
        public int ClusterSleeveInstanceId { get; set; }
        public int CombinedClusterInstanceId { get; set; }
        public string Category { get; set; }
        public int SourceSleeveInstanceId { get; set; }
        
        // Additional properties for ViewModel compatibility
        public int SleeveInstanceId { get; set; }
        public string CategoryName 
        { 
            get => Category; 
            set => Category = value; 
        }
        public XYZ SleeveCenter { get; set; }

        public double ClusterSleeveBoundingBoxMinX { get; set; }
        public double ClusterSleeveBoundingBoxMinY { get; set; }
        public double ClusterSleeveBoundingBoxMinZ { get; set; }
        public double ClusterSleeveBoundingBoxMaxX { get; set; }
        public double ClusterSleeveBoundingBoxMaxY { get; set; }
        public double ClusterSleeveBoundingBoxMaxZ { get; set; }

        public double SleeveWidth { get; set; }
        public double SleeveHeight { get; set; }
        public double SleeveDiameter { get; set; }
        
        // Level information for host plane
        public double Level { get; set; } = 0.0;
        
        // Host type (Wall, Floor, etc.)
        // Host type (Wall, Floor, etc.)
        public string HostType { get; set; } = "Wall";
        public string HostOrientation { get; set; }

        // Compatibility aliases for Phase3And4
        public double Width => SleeveWidth;
        public double Height => SleeveHeight;

        public bool CombinedClusterIncorporated { get; set; }
        public bool IsCombinedResolved { get; set; }
        public string CombinedClusterCategories { get; set; } = string.Empty;

        public Dictionary<string, string> ParameterSnapshot { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public bool HasValidClusterBounds =>
            ClusterSleeveBoundingBoxMaxX > ClusterSleeveBoundingBoxMinX &&
            ClusterSleeveBoundingBoxMaxY > ClusterSleeveBoundingBoxMinY &&
            ClusterSleeveBoundingBoxMaxZ > ClusterSleeveBoundingBoxMinZ;

        // Corner coordinates for precise geometric proximity (from ClashZone or ClusterSleeves DB)
        public double Corner1X { get; set; }
        public double Corner1Y { get; set; }
        public double Corner1Z { get; set; }
        public double Corner2X { get; set; }
        public double Corner2Y { get; set; }
        public double Corner2Z { get; set; }
        public double Corner3X { get; set; }
        public double Corner3Y { get; set; }
        public double Corner3Z { get; set; }
        public double Corner4X { get; set; }
        public double Corner4Y { get; set; }
        public double Corner4Z { get; set; }

        public bool HasValidCorners =>
            (Corner1X != 0 || Corner1Y != 0 || Corner1Z != 0) &&
            (Corner2X != 0 || Corner2Y != 0 || Corner2Z != 0) &&
            (Corner3X != 0 || Corner3Y != 0 || Corner3Z != 0) &&
            (Corner4X != 0 || Corner4Y != 0 || Corner4Z != 0);

        private ClusterSleeveInfo(ClashZone zone) : this()
        {
            ClashZoneId = zone.Id;
            ClashZoneIds.Add(zone.Id);
            
            ClusterSleeveInstanceId = zone.ClusterSleeveInstanceId;
            CombinedClusterInstanceId = zone.CombinedClusterInstanceId;
            Category = zone.MepElementCategory;
            SourceSleeveInstanceId = zone.AfterClusterSleevePlacedSleeveInstanceId > 0
                ? zone.AfterClusterSleevePlacedSleeveInstanceId
                : zone.SleeveInstanceId;
            
            SleeveInstanceId = zone.SleeveInstanceId; 

            // Populate corner data FIRST so we can use it for BBox calculation if needed
            Corner1X = zone.SleeveCorner1X ?? 0;
            Corner1Y = zone.SleeveCorner1Y ?? 0;
            Corner1Z = zone.SleeveCorner1Z ?? 0;
            Corner2X = zone.SleeveCorner2X ?? 0;
            Corner2Y = zone.SleeveCorner2Y ?? 0;
            Corner2Z = zone.SleeveCorner2Z ?? 0;
            Corner3X = zone.SleeveCorner3X ?? 0;
            Corner3Y = zone.SleeveCorner3Y ?? 0;
            Corner3Z = zone.SleeveCorner3Z ?? 0;
            Corner4X = zone.SleeveCorner4X ?? 0;
            Corner4Y = zone.SleeveCorner4Y ?? 0;
            Corner4Z = zone.SleeveCorner4Z ?? 0;

            // Bounding Box Logic: Use explicit fields if available, otherwise derive from Corners
            if (zone.ClusterSleeveBoundingBoxMaxX > zone.ClusterSleeveBoundingBoxMinX)
            {
                ClusterSleeveBoundingBoxMinX = zone.ClusterSleeveBoundingBoxMinX;
                ClusterSleeveBoundingBoxMinY = zone.ClusterSleeveBoundingBoxMinY;
                ClusterSleeveBoundingBoxMinZ = zone.ClusterSleeveBoundingBoxMinZ;
                ClusterSleeveBoundingBoxMaxX = zone.ClusterSleeveBoundingBoxMaxX;
                ClusterSleeveBoundingBoxMaxY = zone.ClusterSleeveBoundingBoxMaxY;
                ClusterSleeveBoundingBoxMaxZ = zone.ClusterSleeveBoundingBoxMaxZ;
            }
            else if (HasValidCorners)
            {
                // Derive AABB from Corners
                var xs = new[] { Corner1X, Corner2X, Corner3X, Corner4X };
                var ys = new[] { Corner1Y, Corner2Y, Corner3Y, Corner4Y };
                var zs = new[] { Corner1Z, Corner2Z, Corner3Z, Corner4Z }; // Simplification (assuming Z is constant-ish or taking min/max)
                
                // Note: Corner3 and 4 might imply thickness/depth, or 1/2/3/4 is just the 2D footprint.
                // Assuming typical 4-corner footprint. For 3D BBox we need min/max of all.
                
                ClusterSleeveBoundingBoxMinX = Math.Min(Math.Min(Corner1X, Corner2X), Math.Min(Corner3X, Corner4X));
                ClusterSleeveBoundingBoxMaxX = Math.Max(Math.Max(Corner1X, Corner2X), Math.Max(Corner3X, Corner4X));
                
                ClusterSleeveBoundingBoxMinY = Math.Min(Math.Min(Corner1Y, Corner2Y), Math.Min(Corner3Y, Corner4Y));
                ClusterSleeveBoundingBoxMaxY = Math.Max(Math.Max(Corner1Y, Corner2Y), Math.Max(Corner3Y, Corner4Y));
                
                // For Z, if corners are just a 2D plane (common in some representations), we might need to rely on stored height or Zone Z.
                // But usually Corners are 3D.
                ClusterSleeveBoundingBoxMinZ = Math.Min(Math.Min(Corner1Z, Corner2Z), Math.Min(Corner3Z, Corner4Z));
                ClusterSleeveBoundingBoxMaxZ = Math.Max(Math.Max(Corner1Z, Corner2Z), Math.Max(Corner3Z, Corner4Z));
                
                // IMPORTANT: If this is a 2D plane (MaxZ == MinZ), apply height? 
                // Repository usually stores full 3D corners or 2D footprint. 
                // If 2D (diff < tolerance), we might need to query SleeveHeight.
                if (Math.Abs(ClusterSleeveBoundingBoxMaxZ - ClusterSleeveBoundingBoxMinZ) < 0.001)
                {
                     double height = zone.SleeveHeight > 0 ? zone.SleeveHeight : zone.SleeveDiameter;
                     if (height > 0) ClusterSleeveBoundingBoxMaxZ += height;
                }
            }
            else
            {
                // Fallback to zone.SleeveBoundingBox... logic or all zeros
                ClusterSleeveBoundingBoxMinX = zone.ClusterSleeveBoundingBoxMinX; 
                // ... leave as 0 if source is 0
            }

            SleeveWidth = zone.SleeveWidth;
            SleeveHeight = zone.SleeveHeight;
            SleeveDiameter = zone.SleeveDiameter;

            CombinedClusterIncorporated = zone.CombinedClusterIncorporated;
            CombinedClusterCategories = zone.CombinedClusterCategories;

            // Host Info
            HostType = zone.StructuralElementType;
            HostOrientation = zone.HostOrientation;

            // Store Original ClashZone for downstream consumers (Proximity Depth, Cleanup)
            OriginalZone = zone;
        }

        public ClashZone OriginalZone { get; set; }

        public static ClusterSleeveInfo? FromClashZone(ClashZone zone)
        {
            if (zone == null)
            {
                return null;
            }

            // ✅ FIX: Accept both individual sleeves AND cluster sleeves for Combined Sleeve discovery
            // Also accept unplaced zones (SleeveInstanceId = -1) for candidate discovery
            bool isIndividualSleeve = zone.SleeveInstanceId > 0;
            bool isClusterSleeve = zone.IsClusterResolved && zone.ClusterSleeveInstanceId > 0;
            bool isUnplacedCandidate = zone.SleeveInstanceId == -1 && !string.IsNullOrEmpty(zone.ClashZoneGuid);
            
            if (!isIndividualSleeve && !isClusterSleeve && !isUnplacedCandidate)
            {
                return null;
            }

            return new ClusterSleeveInfo(zone);
        }

        public void ReplaceParameterSnapshot(IDictionary<string, string> snapshot)
        {
            ParameterSnapshot.Clear();
            if (snapshot == null)
            {
                return;
            }

            foreach (var kvp in snapshot)
            {
                ParameterSnapshot[kvp.Key] = kvp.Value;
            }
        }
    }
}
