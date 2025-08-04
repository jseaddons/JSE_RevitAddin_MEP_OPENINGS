
#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class ClusterBoundingBoxServices
    {
        // Returns bounding box dimensions and midpoint for a cluster of sleeves
        public static (double width, double height, double depth, XYZ mid) GetClusterBoundingBox(List<FamilyInstance>? cluster)
        {
            if (cluster == null || cluster.Count == 0)
                return (0, 0, 0, XYZ.Zero);

            BoundingBoxXYZ? combinedBbox = null;

            foreach (var s in cluster)
            {
                // Use get_BoundingBox(null) which returns an axis-aligned bounding box in the model's coordinate system.
                // This correctly accounts for the element's geometry, size, and orientation.
                BoundingBoxXYZ? bbox = s.get_BoundingBox(null);
                if (bbox == null || !bbox.Enabled) continue;

                if (combinedBbox == null)
                {
                    // Defensive copy to avoid mutating original bbox
                    combinedBbox = new BoundingBoxXYZ
                    {
                        Min = bbox.Min,
                        Max = bbox.Max,
                        Enabled = bbox.Enabled
                    };
                }
                else
                {
                    // Union of the existing combined box and the new sleeve's box
                    combinedBbox.Min = new XYZ(Math.Min(combinedBbox.Min.X, bbox.Min.X),
                                               Math.Min(combinedBbox.Min.Y, bbox.Min.Y),
                                               Math.Min(combinedBbox.Min.Z, bbox.Min.Z));
                    combinedBbox.Max = new XYZ(Math.Max(combinedBbox.Max.X, bbox.Max.X),
                                               Math.Max(combinedBbox.Max.Y, bbox.Max.Y),
                                               Math.Max(combinedBbox.Max.Z, bbox.Max.Z));
                }
            }

            if (combinedBbox == null)
                return (0, 0, 0, XYZ.Zero);

            double widthVal = combinedBbox.Max.X - combinedBbox.Min.X;
            double heightVal = combinedBbox.Max.Y - combinedBbox.Min.Y;
            double depthVal = combinedBbox.Max.Z - combinedBbox.Min.Z;
            XYZ mid = (combinedBbox.Min + combinedBbox.Max) / 2.0;

            return (widthVal, heightVal, depthVal, mid);
        }
    }
}