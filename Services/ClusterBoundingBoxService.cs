using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ClusterBoundingBoxResult
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public XYZ Midpoint { get; set; }
        public double ZSum { get; set; }
        public bool IsXAxisMEP { get; set; }
        public bool IsDuctFloorCluster { get; set; }
    }

    public static class ClusterBoundingBoxService
    {
        // Returns bounding box, midpoint, and swap info for a cluster
        public static ClusterBoundingBoxResult GetClusterBoundingBox(
            IList<FamilyInstance> cluster,
            IDictionary<FamilyInstance, XYZ> groupLocations)
        {
            var result = new ClusterBoundingBoxResult();
            if (cluster.Count == 0)
                return result;

            // Gather all 8 corners of each element's bounding box, transformed to model coordinates
            List<XYZ> allCorners = new List<XYZ>();
            double zSum = 0;
            foreach (var s in cluster)
            {
                BoundingBoxXYZ bbox = s.get_BoundingBox(null);
                if (bbox == null) continue;
                Transform t = s.GetTransform();
                foreach (XYZ corner in GetBoundingBoxCorners(bbox))
                {
                    allCorners.Add(t.OfPoint(corner));
                }
                zSum += groupLocations[s].Z;
            }
            if (allCorners.Count == 0)
                return result;

            double minX = allCorners.Min(pt => pt.X);
            double maxX = allCorners.Max(pt => pt.X);
            double minY = allCorners.Min(pt => pt.Y);
            double maxY = allCorners.Max(pt => pt.Y);
            double minZ = allCorners.Min(pt => pt.Z);
            double maxZ = allCorners.Max(pt => pt.Z);

            double width = maxX - minX;
            double height = maxY - minY;
            result.Width = width;
            result.Height = height;
            result.Midpoint = new XYZ((minX + maxX) / 2, (minY + maxY) / 2, (minZ + maxZ) / 2);
            result.ZSum = zSum;
            result.IsDuctFloorCluster = false;
            result.IsXAxisMEP = false;

            // Cluster summary logging for cable tray wall clusters
            if (cluster.Count > 0) {
                string famName = cluster[0].Symbol?.Family?.Name ?? "";
                if (famName.Equals("CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase)) {
                    string clusterLogPath = @"C:\\JSE_CSharp_Projects\\JSE_RevitAddin_MEP_OPENINGS\\JSE_RevitAddin_MEP_OPENINGS\\Logs\\CableTrayWallBBoxStepLog.txt";
                    System.IO.File.AppendAllText(clusterLogPath, $"\n=== CLUSTER SUMMARY ===\n");
                    System.IO.File.AppendAllText(clusterLogPath, $"Cluster Width (Y): {result.Width}\n");
                    System.IO.File.AppendAllText(clusterLogPath, $"Cluster Height (X): {result.Height}\n");
                    System.IO.File.AppendAllText(clusterLogPath, $"Cluster Midpoint: ({result.Midpoint.X}, {result.Midpoint.Y}, {result.Midpoint.Z})\n");
                }
            }
            return result;
        }

        // Helper: Get all 8 corners of a bounding box
        private static IEnumerable<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ bbox)
        {
            yield return new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z);
        }
    }
}
