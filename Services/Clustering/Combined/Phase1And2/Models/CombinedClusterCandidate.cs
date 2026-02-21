using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models
{
    /// <summary>
    /// Represents a combined cluster candidate built from multiple category-specific cluster sleeves.
    /// </summary>
    public class CombinedClusterCandidate
    {
        public CombinedClusterCandidate(long combinedClusterInstanceId, IEnumerable<ClusterSleeveInfo> members)
        {
            CombinedClusterInstanceId = combinedClusterInstanceId;
            MemberClusters = new List<ClusterSleeveInfo>();
            CategoriesInvolved = new List<string>();
            IncorporatedIndividualSleeves = new List<ClashZone>();
            
            foreach (var member in members)
            {
                AddMember(member);
            }
        }

        public long CombinedClusterInstanceId { get; }
        public List<ClusterSleeveInfo> MemberClusters { get; }
        public List<string> CategoriesInvolved { get; }
        public List<ClashZone> IncorporatedIndividualSleeves { get; set; }

        // Bounding box components (for compatibility)
        public double CombinedBoundingBoxMinX { get; internal set; } = double.MaxValue;
        public double CombinedBoundingBoxMinY { get; internal set; } = double.MaxValue;
        public double CombinedBoundingBoxMinZ { get; internal set; } = double.MaxValue;
        public double CombinedBoundingBoxMaxX { get; internal set; } = double.MinValue;
        public double CombinedBoundingBoxMaxY { get; internal set; } = double.MinValue;
        public double CombinedBoundingBoxMaxZ { get; internal set; } = double.MinValue;

        // Derived bounding box as BoundingBoxXYZ for Phase3And4 compatibility
        public BoundingBoxXYZ CombinedBoundingBox
        {
            get
            {
                var bbox = new BoundingBoxXYZ();
                bbox.Min = new XYZ(CombinedBoundingBoxMinX, CombinedBoundingBoxMinY, CombinedBoundingBoxMinZ);
                bbox.Max = new XYZ(CombinedBoundingBoxMaxX, CombinedBoundingBoxMaxY, CombinedBoundingBoxMaxZ);
                return bbox;
            }
        }

        // Derived dimensions
        public double CombinedWidth => Math.Max(0, CombinedBoundingBoxMaxX - CombinedBoundingBoxMinX);
        public double CombinedHeight => Math.Max(0, CombinedBoundingBoxMaxZ - CombinedBoundingBoxMinZ);

        // Host/Orientation metadata (from first member for now)
        public string HostType { get; set; } = "Wall";
        public string Orientation { get; set; } = "X-Wall";
        public double Level { get; set; } = 0.0;

        public AggregatedParameterSnapshot AggregatedSnapshot { get; set; }

        private void AddMember(ClusterSleeveInfo member)
        {
            if (member == null)
                return;

            MemberClusters.Add(member);
            if (!CategoriesInvolved.Contains(member.Category))
            {
                CategoriesInvolved.Add(member.Category);
            }

            CombinedBoundingBoxMinX = Math.Min(CombinedBoundingBoxMinX, member.ClusterSleeveBoundingBoxMinX);
            CombinedBoundingBoxMinY = Math.Min(CombinedBoundingBoxMinY, member.ClusterSleeveBoundingBoxMinY);
            CombinedBoundingBoxMinZ = Math.Min(CombinedBoundingBoxMinZ, member.ClusterSleeveBoundingBoxMinZ);
            CombinedBoundingBoxMaxX = Math.Max(CombinedBoundingBoxMaxX, member.ClusterSleeveBoundingBoxMaxX);
            CombinedBoundingBoxMaxY = Math.Max(CombinedBoundingBoxMaxY, member.ClusterSleeveBoundingBoxMaxY);
            CombinedBoundingBoxMaxZ = Math.Max(CombinedBoundingBoxMaxZ, member.ClusterSleeveBoundingBoxMaxZ);

            member.CombinedClusterInstanceId = CombinedClusterInstanceId;
            member.CombinedClusterCategories = string.Join(",", CategoriesInvolved.OrderBy(c => c));
            member.CombinedClusterIncorporated = true;
        }
    }
}
