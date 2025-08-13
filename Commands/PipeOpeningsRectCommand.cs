using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class PipeOpeningsRectCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            int placedCount = 0;
            int deletedCount = 0;

            // Collect all placed pipe sleeves (assuming circular PS# family instances)
            double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
            double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
            double zTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("PS#"))
                .ToList();
            
            if (sleeves.Count == 0)
            {
                return Result.Succeeded;
            }

            // Find rectangular opening symbol
            var allFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            // Use ClusterOpeningOnWallX family for clusters instead of PipeOpeningOnWallRect
            var rectSymbol = allFamilySymbols
                .FirstOrDefault(sym => sym.Family.Name.Equals("ClusterOpeningOnWallX", System.StringComparison.OrdinalIgnoreCase));
            if (rectSymbol == null)
            {
                TaskDialog.Show("Error", "Please load the ClusterOpeningOnWallX family.");
                return Result.Failed;
            }

            using (var tx = new Transaction(doc, "Place Rectangular Pipe Openings"))
            {
                DebugLogger.Log("Starting transaction: Place Rectangular Pipe Openings");
                tx.Start();
                if (!rectSymbol.IsActive)
                    rectSymbol.Activate();

                // Build map of sleeve to insertion point (use LocationPoint if available)
                var sleeveLocations = sleeves.ToDictionary(
                    s => s,
                    s =>
                    {
                        if (s.Location is LocationPoint lp)
                            return lp.Point;
                        return s.GetTransform().Origin;
                    });

                // Collect all placed pipe sleeves (rectangular and circular) for suppression
                double suppressionTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm
                var allPipeSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("Pipe"))
                    .ToList();
                var allPipeSleeveLocations = allPipeSleeves.ToDictionary(
                    s => s,
                    s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin);

                // Ensure all sleeves have a valid Level and Schedule Level parameter before clustering
                foreach (var sleeve in sleeves)
                {
                    DebugLogger.Log($"Processing sleeve {sleeve.Id.Value} for level assignment");
                    // Try to get reference level from parameter or helper
                    Level? refLevelNullable = HostLevelHelper.GetHostReferenceLevel(doc, sleeve);
                    Level refLevel;
                    if (refLevelNullable == null)
                    {
                        var pt = (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin;
                        refLevel = GetNearestPositiveZLevel(doc, pt);
                        DebugLogger.Log($"Using nearest positive Z level for sleeve {sleeve.Id.Value}: {refLevel?.Name ?? "null"}");
                    }
                    else
                    {
                        refLevel = refLevelNullable;
                        DebugLogger.Log($"Got reference level from helper for sleeve {sleeve.Id.Value}: {refLevel.Name}");
                    }
                    if (refLevel != null)
                    {
                        Parameter levelParam = sleeve.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                            ?? sleeve.LookupParameter("Level")
                            ?? sleeve.LookupParameter("Reference Level");
                        if (levelParam != null && !levelParam.IsReadOnly && levelParam.StorageType == StorageType.ElementId)
                        {
                            levelParam.Set(refLevel.Id);
                            DebugLogger.Log($"Set Level parameter for sleeve {sleeve.Id.Value} to {refLevel.Name}");
                        }
                        // Set Schedule Level if available
                        var schedLevelParam = sleeve.LookupParameter("Schedule Level");
                        if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                        {
                            DebugLogger.Log($"Setting Schedule Level for sleeve {sleeve.Id.Value}, StorageType: {schedLevelParam.StorageType}");
                            if (schedLevelParam.StorageType == StorageType.ElementId)
                            {
                                schedLevelParam.Set(refLevel.Id);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.Value} Schedule Level to ElementId: {refLevel.Id.Value} ({refLevel.Name})");
                            }
                            else if (schedLevelParam.StorageType == StorageType.String)
                            {
                                schedLevelParam.Set(refLevel.Name);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.Value} Schedule Level to String: '{refLevel.Name}'");
                            }
                            else if (schedLevelParam.StorageType == StorageType.Integer)
                            {
                                schedLevelParam.Set(refLevel.Id.Value);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.Value} Schedule Level to Integer: {refLevel.Id.Value} ({refLevel.Name})");
                            }
                        }
                        else
                        {
                            DebugLogger.Log($"Sleeve {sleeve.Id.Value} has no Schedule Level parameter or it's read-only");
                        }
                    }
                    else
                    {
                        DebugLogger.Log($"No reference level found for sleeve {sleeve.Id.Value}");
                    }
                }

                // Group sleeves into clusters where any two within toleranceDistance are connected
                var clusters = new List<List<FamilyInstance>>();
                var unprocessed = new List<FamilyInstance>(sleeves);
                while (unprocessed.Count > 0)
                {
                    var queue = new Queue<FamilyInstance>();
                    var cluster = new List<FamilyInstance>();
                    queue.Enqueue(unprocessed[0]);
                    unprocessed.RemoveAt(0);
                    while (queue.Count > 0)
                    {
                        var inst = queue.Dequeue();
                        cluster.Add(inst);
                        // find neighbors within tolerance (planar XY only)
                        var o1 = sleeveLocations[inst];
                        var neighbors = unprocessed.Where(s =>
                        {
                            XYZ o2 = sleeveLocations[s];
                            // skip if vertical offset exceeds threshold
                            if (Math.Abs(o1.Z - o2.Z) > zTolerance)
                                return false;
                            // compute planar XY gap
                            double dx = o1.X - o2.X;
                            double dy = o1.Y - o2.Y;
                            double planar = Math.Sqrt(dx * dx + dy * dy);
                            // account for sleeve diameters (edge-to-edge gap)
                            double dia1 = inst.LookupParameter("Diameter")?.AsDouble() ?? 0;
                            double dia2 = s.LookupParameter("Diameter")?.AsDouble() ?? 0;
                            double gap = planar - (dia1 / 2.0 + dia2 / 2.0);
                            double gapMm = UnitUtils.ConvertFromInternalUnits(gap, UnitTypeId.Millimeters);
                            DebugLogger.Log($"Edge gap between {inst.Id} and {s.Id}: {gapMm:F1} mm");
                            return gap <= toleranceDist;
                        }).ToList();
                        foreach (var n in neighbors)
                        {
                            queue.Enqueue(n);
                            unprocessed.Remove(n);
                        }
                    }
                    clusters.Add(cluster);
                }
                // Log cluster summary
                DebugLogger.Log($"Cluster formation complete. Total clusters: {clusters.Count}");
                for (int ci = 0; ci < clusters.Count; ci++)
                {
                    DebugLogger.Log($"Cluster {ci}: {clusters[ci].Count} sleeves");
                }
                // Process each cluster: one rectangle per cluster of size >=2
                foreach (var cluster in clusters)
                {
                    if (cluster.Count < 2)
                        continue;
                    DebugLogger.Log($"Cluster of {cluster.Count} sleeves detected for replacement.");
                    // Use ClusterBoundingBoxServices to get bounding box and midpoint (like RectangularSleeveClusterCommandV2)
                    var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(cluster);
                    // Use reference level from first sleeve in cluster
                    Level? refLevelNullable = HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]);
                    Level refLevel;
                    if (refLevelNullable == null)
                    {
                        var pt = (cluster[0].Location as LocationPoint)?.Point ?? cluster[0].GetTransform().Origin;
                        refLevel = GetNearestPositiveZLevel(doc, pt);
                    }
                    else
                    {
                        refLevel = refLevelNullable;
                    }
                    // Duplicate suppression
                    double clusterSuppressionTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    var existingRects = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol.Family.Name.Equals("ClusterOpeningOnWallX", System.StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    bool duplicateFound = existingRects.Any(rect => {
                        var loc = (rect.Location as LocationPoint)?.Point ?? rect.GetTransform().Origin;
                        double dist = mid.DistanceTo(loc);
                        return dist <= clusterSuppressionTol;
                    });
                    if (duplicateFound)
                    {
                        DebugLogger.Log($"Suppressed duplicate rectangular opening at {mid} (existing rectangular opening within 100mm)");
                        continue;
                    }
                    // Place the cluster sleeve family instance at the cluster midpoint
                    FamilyInstance inst = doc.Create.NewFamilyInstance(mid, rectSymbol, refLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    // Set Width, Height, Depth parameters mapping model bbox dimensions correctly
                    var widthParam = inst.LookupParameter("Width");
                    var heightParam = inst.LookupParameter("Height");
                    var depthParam = inst.LookupParameter("Depth");
                    // For Depth, use host thickness if available (like RectangularSleeveClusterCommandV2)
                    double hostThickness = width;
                    var wall = cluster[0].Host as Wall;
                    if (wall != null)
                    {
                        hostThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                        DebugLogger.Log($"[CLUSTER] Wall thickness used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                    }
                    if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(height);   // Y
                    if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(depth); // Z
                    if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(hostThickness); // X
                    placedCount++;
                    DebugLogger.Log($"Rectangular opening created with id {inst.Id.Value} (total placed: {placedCount})");
                    // Delete originals
                    foreach (var s in cluster)
                    {
                        doc.Delete(s.Id);
                        deletedCount++;
                    }
                    DebugLogger.Log($"Deleted {cluster.Count} circular sleeves (total deleted: {deletedCount})");
                }
                tx.Commit();
            }

            // Simple summary dialog - no intrusive details
            // Info dialog removed as requested
            return Result.Succeeded;
        }

        // Helper to get the nearest positive Z level (at or above the given Z)
        public static Level GetNearestPositiveZLevel(Document doc, XYZ point)
        {
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            // Only consider levels at or above the point's Z
            var aboveOrAt = levels.Where(l => l.Elevation >= point.Z).OrderBy(l => l.Elevation - point.Z).ToList();
            if (aboveOrAt.Any())
                return aboveOrAt.First();
            // Fallback: nearest level (if all are below)
            return levels.OrderBy(l => Math.Abs(l.Elevation - point.Z)).FirstOrDefault();
        }

        // Helper to get the nearest level strictly below the given Z
        public static Level GetNearestLevelBelowZ(Document doc, XYZ point)
        {
            var levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            // Only consider levels strictly below the point's Z
            var below = levels.Where(l => l.Elevation < point.Z).OrderByDescending(l => l.Elevation).ToList();
            if (below.Any())
                return below.First();
            // Fallback: lowest level in the project
            return levels.OrderBy(l => l.Elevation).FirstOrDefault();
        }
    }
    }
