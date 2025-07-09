using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class PipeOpeningsRectCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.InitLogFile("PipeOpeningsRectCommand");
            DebugLogger.Log("PipeOpeningsRectCommand: Execute started.");
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            int placedCount = 0;
            int deletedCount = 0;

            double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
            double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
            double zTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);

            // Define all types to process (wall and slab)
            var types = new[] {
                new { Name = "Pipe", Family = "PipeOpeningOnWall", RectFamily = "PipeOpeningOnWallRect", SymbolPrefix = "PS#", HostType = "Wall" },
                new { Name = "Duct", Family = "DuctOpeningOnWall", RectFamily = "DuctOpeningOnWallRect", SymbolPrefix = "DS#", HostType = "Wall" },
                new { Name = "CableTray", Family = "CableTrayOpeningOnWall", RectFamily = "CableTrayOpeningOnWallRect", SymbolPrefix = "CT#", HostType = "Wall" },
                new { Name = "Pipe", Family = "PipeOpeningOnSlab", RectFamily = "PipeOpeningOnSlabRect", SymbolPrefix = "PS#", HostType = "Floor" },
                new { Name = "Duct", Family = "DuctOpeningOnSlab", RectFamily = "DuctOpeningOnSlabRect", SymbolPrefix = "DS#", HostType = "Floor" },
                new { Name = "CableTray", Family = "CableTrayOpeningOnSlab", RectFamily = "CableTrayOpeningOnSlabRect", SymbolPrefix = "CT#", HostType = "Floor" }
            };

            using (var tx = new Transaction(doc, "Place Rectangular MEP Cluster Openings"))
            {
                DebugLogger.Log("Starting transaction: Place Rectangular MEP Cluster Openings");
                tx.Start();

                foreach (var t in types)
                {
                    // Collect sleeves for this type (wall or slab)
                    var sleeves = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol.Family.Name == t.Family)
                        .Where(fi => fi.Symbol.Name.StartsWith(t.SymbolPrefix))
                        .ToList();
                    DebugLogger.Log($"Collected {sleeves.Count} {t.Name.ToLower()} sleeves for {t.HostType}. Tolerance: {toleranceMm:F1} mm");

                    // Find rectangular opening symbol for this type
                    var rectSymbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(sym => sym.Family.Name == t.RectFamily);
                    if (rectSymbol == null)
                    {
                        DebugLogger.Log($"{t.RectFamily} family not found. Skipping {t.Name.ToLower()} {t.HostType.ToLower()} cluster replacement.");
                        continue;
                    }
                    if (!rectSymbol.IsActive)
                        rectSymbol.Activate();

                    // Build map of sleeve to insertion point
                    var sleeveLocations = sleeves.ToDictionary(
                        s => s,
                        s => (s.Location is LocationPoint lp) ? lp.Point : s.GetTransform().Origin);

                    // For slab, log extra info
                    if (t.HostType == "Floor")
                        DebugLogger.Log($"[SLAB] Processing {t.Name} sleeves for slab clustering.");

                    // Group sleeves into clusters (same logic for wall/slab)
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
                            var o1 = sleeveLocations[inst];
                            var neighbors = unprocessed.Where(s =>
                            {
                                XYZ o2 = sleeveLocations[s];
                                if (Math.Abs(o1.Z - o2.Z) > zTolerance)
                                    return false;
                                double dx = o1.X - o2.X;
                                double dy = o1.Y - o2.Y;
                                double planar = Math.Sqrt(dx * dx + dy * dy);
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
                    DebugLogger.Log($"Cluster formation complete for {t.Name.ToLower()} ({t.HostType}). Total clusters: {clusters.Count}");
                    for (int ci = 0; ci < clusters.Count; ci++)
                        DebugLogger.Log($"Cluster {ci}: {clusters[ci].Count} sleeves");

                    foreach (var cluster in clusters)
                    {
                        if (cluster.Count < 2)
                        {
                            if (t.HostType == "Floor")
                                DebugLogger.Log($"[SLAB] Skipping cluster (only {cluster.Count} sleeve(s))");
                            continue;
                        }
                        DebugLogger.Log($"Cluster of {cluster.Count} {t.Name.ToLower()} sleeves detected for replacement. HostType: {t.HostType}");
                        var clusterCenter = new XYZ(
                            cluster.Average(s => sleeveLocations[s].X),
                            cluster.Average(s => sleeveLocations[s].Y),
                            cluster.Average(s => sleeveLocations[s].Z)
                        );
                        // Focused suppression logging: log all nearby openings, their type, and suppression reason
                        var allNearbyOpenings = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi =>
                            {
                                var fiLocation = (fi.Location as LocationPoint)?.Point;
                                if (fiLocation == null) return false;
                                var distance = clusterCenter.DistanceTo(fiLocation);
                                return distance <= toleranceDist;
                            })
                            .ToList();
                        DebugLogger.Log($"Suppression check: Found {allNearbyOpenings.Count} FamilyInstances within {toleranceMm:F1}mm of cluster center {clusterCenter}.");
                        foreach (var fi in allNearbyOpenings)
                        {
                            var fiLocation = (fi.Location as LocationPoint)?.Point;
                            var distance = fiLocation != null ? clusterCenter.DistanceTo(fiLocation) : -1;
                            string fam = fi.Symbol?.Family?.Name ?? "<null>";
                            string sym = fi.Symbol?.Name ?? "<null>";
                            bool isRect = fam == t.RectFamily;
                            bool isIndiv = fam == t.Family;
                            DebugLogger.Log($"  - Candidate ID {fi.Id.IntegerValue}, Family: {fam}, Symbol: {sym}, Location: {fiLocation}, Distance: {UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm, Type: {(isRect ? "RECT_CLUSTER" : (isIndiv ? "INDIVIDUAL" : "OTHER"))}");
                        }
                        var existingRectOpenings = allNearbyOpenings
                            .Where(fi => fi.Symbol?.Family?.Name != null && fi.Symbol.Family.Name == t.RectFamily)
                            .ToList();
                        var existingIndivOpenings = allNearbyOpenings
                            .Where(fi => fi.Symbol?.Family?.Name != null && fi.Symbol.Family.Name == t.Family)
                            .ToList();
                        if (existingRectOpenings.Any())
                        {
                            DebugLogger.Log($"DUPLICATION SUPPRESSION: Skipping cluster at {clusterCenter} - found {existingRectOpenings.Count} existing rectangular opening(s) within {toleranceMm:F1}mm tolerance");
                            foreach (var existing in existingRectOpenings)
                            {
                                var existingLocation = (existing.Location as LocationPoint)?.Point;
                                var distance = clusterCenter.DistanceTo(existingLocation);
                                DebugLogger.Log($"  - SUPPRESSING: Rectangular opening ID {existing.Id.IntegerValue} at distance {UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm, Family: {existing.Symbol.Family.Name}, Symbol: {existing.Symbol.Name}, Location: {existingLocation}");
                            }
                            continue;
                        }
                        if (existingIndivOpenings.Any())
                        {
                            DebugLogger.Log($"Suppression: Found {existingIndivOpenings.Count} individual sleeve(s) (Family: {t.Family}) within {toleranceMm:F1}mm, but NOT suppressing cluster (cluster will replace individuals).");
                        }
                        // Compute extents
                        var xMinEdge = double.MaxValue; var xMaxEdge = double.MinValue;
                        var yMinEdge = double.MaxValue; var yMaxEdge = double.MinValue;
                        foreach (var s in cluster)
                        {
                            var o = sleeveLocations[s];
                            double dia = s.LookupParameter("Diameter")?.AsDouble() ?? 0.0;
                            double r = dia / 2.0;
                            xMinEdge = Math.Min(xMinEdge, o.X - r);
                            xMaxEdge = Math.Max(xMaxEdge, o.X + r);
                            yMinEdge = Math.Min(yMinEdge, o.Y - r);
                            yMaxEdge = Math.Max(yMaxEdge, o.Y + r);
                        }
                        double midX = (xMinEdge + xMaxEdge) / 2.0;
                        double midY = (yMinEdge + yMaxEdge) / 2.0;
                        double midZ = cluster.Average(s => sleeveLocations[s].Z);
                        var midPoint = new XYZ(midX, midY, midZ);
                        double widthInternal = xMaxEdge - xMinEdge;
                        double heightInternal = yMaxEdge - yMinEdge;
                        double widthMm = UnitUtils.ConvertFromInternalUnits(widthInternal, UnitTypeId.Millimeters);
                        double heightMm = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters);
                        double minDimMm = 20.0;
                        if (widthMm < minDimMm || heightMm < minDimMm)
                        {
                            DebugLogger.Log($"[VALIDATION] Skipping placement: width={widthMm:F1}mm, height={heightMm:F1}mm below minimum {minDimMm}mm. Cluster center: {midPoint}");
                            continue;
                        }
                        DebugLogger.Log($"Placing rectangular opening at {midPoint} with width {widthMm:F1} mm, height {heightMm:F1} mm");
                        double fallbackDepth = 0.0;
                        if (t.HostType == "Wall")
                        {
                            var wallHost = cluster[0].Host as Wall;
                            fallbackDepth = wallHost != null ? wallHost.Width : 0.0;
                            if (wallHost != null)
                                DebugLogger.Log($"Wall thickness for depth fallback: {UnitUtils.ConvertFromInternalUnits(fallbackDepth, UnitTypeId.Millimeters):F1} mm");
                            else
                                DebugLogger.Log("No wall host found - families are unhosted Generic Models in linked walls");
                        }
                        else if (t.HostType == "Floor")
                        {
                            // For slab/floor clusters: always unhosted, no fallback, no structural check, no host required.
                            // Use only the max Depth from the individual sleeves in the cluster.
                            fallbackDepth = 0.0;
                        }
                        var sleeveDepths = cluster.Select(s => s.LookupParameter("Depth")?.AsDouble() ?? 0.0);
                        double depthInternal = sleeveDepths.DefaultIfEmpty(fallbackDepth).Max();
                        DebugLogger.Log($"Calculated opening depth: {UnitUtils.ConvertFromInternalUnits(depthInternal, UnitTypeId.Millimeters):F1} mm");
                        // For slab/floor clusters: do not require or use Level for placement.
                        // Aspect ratio logic for width/height swap (same as RectangularSleeveClusterCommand)
                        double finalWidth = widthInternal;
                        double finalHeight = heightInternal;
                        if (t.HostType == "Wall")
                        {
                            double aspectRatio = widthInternal / heightInternal;
                            if (aspectRatio > 1.5)
                            {
                                // Horizontal arrangement - keep original
                            }
                            else if (aspectRatio < 0.67)
                            {
                                // Vertical arrangement - swap
                                finalWidth = heightInternal;
                                finalHeight = widthInternal;
                            }
                            else
                            {
                                // Nearly square - default to vertical wall behavior (swap)
                                finalWidth = heightInternal;
                                finalHeight = widthInternal;
                            }
                        }
                        // For Floor (slab), do not swap width/height
                        FamilyInstance rectInst = null;
                        try
                        {
                            // IMPORTANT: For Generic Model, work plane-based families (as used for sleeves), the correct overload is NewFamilyInstance(XYZ, FamilySymbol, StructuralType.NonStructural)
                            // Do NOT use the Level argument for these families, especially for slab/floor placements, as it can cause failures or incorrect placement.
                            rectInst = doc.Create.NewFamilyInstance(midPoint, rectSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            rectInst.LookupParameter("Depth")?.Set(depthInternal);
                            rectInst.LookupParameter("Width")?.Set(finalWidth);
                            rectInst.LookupParameter("Height")?.Set(finalHeight);
                            var origRot = (cluster[0].Location as LocationPoint)?.Rotation ?? 0;
                            if (origRot != 0)
                            {
                                var axis = Line.CreateBound(midPoint, midPoint + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(doc, rectInst.Id, axis, origRot);
                            }
                            DebugLogger.Log($"Rectangular {t.Name.ToLower()} opening placed with computed Width/Height/Depth and original rotation.");
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[ERROR] Exception during rectangular opening placement: {ex.Message}");
                        }
                        bool instanceExists = rectInst != null && doc.GetElement(rectInst.Id) != null;
                        if (rectInst != null && instanceExists)
                        {
                            placedCount++;
                            DebugLogger.Log($"Rectangular opening created with id {rectInst.Id.IntegerValue} (total placed: {placedCount})");
                            foreach (var s in cluster)
                            {
                                doc.Delete(s.Id);
                                deletedCount++;
                            }
                            DebugLogger.Log($"Deleted {cluster.Count} original sleeves (total deleted: {deletedCount})");
                        }
                        else if (rectInst != null && !instanceExists)
                        {
                            DebugLogger.Log($"[ERROR] Rectangular opening instance {rectInst.Id.IntegerValue} was deleted by Revit after creation (likely family error).");
                        }
                    }
                }
                DebugLogger.Log($"Transaction committed for rectangular openings.");
                DebugLogger.Log($"Summary: {placedCount} rectangular openings placed, {deletedCount} original sleeves deleted.");
                tx.Commit();
            }
            DebugLogger.Log("PipeOpeningsRectCommand: Execute completed.");
            TaskDialog.Show("Done", $"Rectangular openings placed: {placedCount}. Original sleeves deleted: {deletedCount}.");
            return Result.Succeeded;
        }
    }
}
