using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ClusterSleeveDuplicationService
    {
        public static void JoinNearbyOpenings(Document doc, Models.SettingsModel settings)
        {
            // Get all sleeves in the document
            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab"))
                .ToList();

            // Create a list of sleeves to be deleted
            var sleevesToDelete = new List<ElementId>();
            var newClusters = new List<FamilyInstance>();

            // Iterate through all pairs of sleeves
            for (int i = 0; i < sleeves.Count; i++)
            {
                for (int j = i + 1; j < sleeves.Count; j++)
                {
                    var sleeve1 = sleeves[i];
                    var sleeve2 = sleeves[j];

                    // Skip if either sleeve is already marked for deletion
                    if (sleevesToDelete.Contains(sleeve1.Id) || sleevesToDelete.Contains(sleeve2.Id))
                        continue;

                    // Get the location of the sleeves
                    var location1 = (sleeve1.Location as LocationPoint)?.Point;
                    var location2 = (sleeve2.Location as LocationPoint)?.Point;

                    if (location1 != null && location2 != null)
                    {
                        // Check if the distance between the sleeves is less than the threshold
                        if (location1.DistanceTo(location2) < settings.JoinOpeningsDistance)
                        {
                            // Calculate the bounding box that encloses both sleeves
                            var bbox1 = sleeve1.get_BoundingBox(null);
                            var bbox2 = sleeve2.get_BoundingBox(null);

                            if (bbox1 != null && bbox2 != null)
                            {
                                var min = new XYZ(Math.Min(bbox1.Min.X, bbox2.Min.X), Math.Min(bbox1.Min.Y, bbox2.Min.Y), Math.Min(bbox1.Min.Z, bbox2.Min.Z));
                                var max = new XYZ(Math.Max(bbox1.Max.X, bbox2.Max.X), Math.Max(bbox1.Max.Y, bbox2.Max.Y), Math.Max(bbox1.Max.Z, bbox2.Max.Z));

                                var clusterBbox = new BoundingBoxXYZ { Min = min, Max = max };

                                // Find a suitable family symbol for the cluster sleeve
                                // TODO: Make this configurable
                                var clusterSymbol = new FilteredElementCollector(doc)
                                    .OfClass(typeof(FamilySymbol))
                                    .OfCategory(BuiltInCategory.OST_GenericModel)
                                    .Cast<FamilySymbol>()
                                    .FirstOrDefault(s => s.Family.Name.Contains("ClusterOpening"));

                                if (clusterSymbol != null)
                                {
                                    // Create a new cluster sleeve
                                    var newCluster = doc.Create.NewFamilyInstance(clusterBbox.Min, clusterSymbol, doc.ActiveView);
                                    newClusters.Add(newCluster);

                                    // Set the parameters of the new cluster sleeve
                                    var width = clusterBbox.Max.X - clusterBbox.Min.X;
                                    var height = clusterBbox.Max.Y - clusterBbox.Min.Y;
                                    var depth = clusterBbox.Max.Z - clusterBbox.Min.Z;

                                    newCluster.LookupParameter("Width").Set(width);
                                    newCluster.LookupParameter("Height").Set(height);
                                    newCluster.LookupParameter("Depth").Set(depth);
                                }

                                // Mark the original sleeves for deletion
                                sleevesToDelete.Add(sleeve1.Id);
                                sleevesToDelete.Add(sleeve2.Id);
                            }
                        }
                    }
                }
            }

            // Delete the original sleeves
            if (sleevesToDelete.Any())
            {
                using (var transaction = new Transaction(doc, "Delete original sleeves"))
                {
                    transaction.Start();

                    doc.Delete(sleevesToDelete);
                    transaction.Commit();
                }
            }
        }
    }
}
