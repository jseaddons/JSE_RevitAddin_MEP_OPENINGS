using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models
{
    /// <summary>
    /// DTO for cluster sleeves feeding combined clustering (Phase 1-2, DB/CPU only).
    /// </summary>
    public class ClusterSleeveInfo
    {
        public Guid ClashZoneId { get; }
        public int ClusterSleeveInstanceId { get; }
        public int CombinedClusterInstanceId { get; set; }
        public string Category { get; }
        public int SourceSleeveInstanceId { get; }

        public double ClusterSleeveBoundingBoxMinX { get; }
        public double ClusterSleeveBoundingBoxMinY { get; }
        public double ClusterSleeveBoundingBoxMinZ { get; }
        public double ClusterSleeveBoundingBoxMaxX { get; }
        public double ClusterSleeveBoundingBoxMaxY { get; }
        public double ClusterSleeveBoundingBoxMaxZ { get; }

        public double SleeveWidth { get; }
        public double SleeveHeight { get; }
        public double SleeveDiameter { get; }
        
        // Level information for host plane
        public double Level { get; set; } = 0.0;
        
        // Host type (Wall, Floor, etc.)
        public string HostType { get; set; } = "Wall";

        // Compatibility aliases for Phase3And4
        public double Width => SleeveWidth;
        public double Height => SleeveHeight;

        public bool CombinedClusterIncorporated { get; set; }
        public string CombinedClusterCategories { get; set; } = string.Empty;

        public Dictionary<string, string> ParameterSnapshot { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public bool HasValidClusterBounds =>
            ClusterSleeveBoundingBoxMaxX > ClusterSleeveBoundingBoxMinX &&
            ClusterSleeveBoundingBoxMaxY > ClusterSleeveBoundingBoxMinY &&
            ClusterSleeveBoundingBoxMaxZ > ClusterSleeveBoundingBoxMinZ;

        private ClusterSleeveInfo(ClashZone zone)
        {
            ClashZoneId = zone.Id;
            ClusterSleeveInstanceId = zone.ClusterSleeveInstanceId;
            CombinedClusterInstanceId = zone.CombinedClusterInstanceId;
            Category = zone.MepElementCategory;
            SourceSleeveInstanceId = zone.AfterClusterSleevePlacedSleeveInstanceId > 0
                ? zone.AfterClusterSleevePlacedSleeveInstanceId
                : zone.SleeveInstanceId;

            ClusterSleeveBoundingBoxMinX = zone.ClusterSleeveBoundingBoxMinX;
            ClusterSleeveBoundingBoxMinY = zone.ClusterSleeveBoundingBoxMinY;
            ClusterSleeveBoundingBoxMinZ = zone.ClusterSleeveBoundingBoxMinZ;
            ClusterSleeveBoundingBoxMaxX = zone.ClusterSleeveBoundingBoxMaxX;
            ClusterSleeveBoundingBoxMaxY = zone.ClusterSleeveBoundingBoxMaxY;
            ClusterSleeveBoundingBoxMaxZ = zone.ClusterSleeveBoundingBoxMaxZ;

            SleeveWidth = zone.SleeveWidth;
            SleeveHeight = zone.SleeveHeight;
            SleeveDiameter = zone.SleeveDiameter;

            CombinedClusterIncorporated = zone.CombinedClusterIncorporated;
            CombinedClusterCategories = zone.CombinedClusterCategories;
        }

        public static ClusterSleeveInfo? FromClashZone(ClashZone zone)
        {
            if (zone == null)
            {
                return null;
            }

            if (!zone.IsClusterResolved || zone.ClusterSleeveInstanceId <= 0)
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
