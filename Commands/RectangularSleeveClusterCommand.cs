using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class RectangularSleeveClusterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.InitLogFile("RectangularSleeveClusterCommand");
            DebugLogger.Log("RectangularSleeveClusterCommand: Execute started.");
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            int placedCount = 0;
            int deletedCount = 0;

            // DEBUG: Log all available wall sleeve families and symbols in the model
            try
            {
                var allWallSleeveInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol != null && fi.Symbol.Family != null && fi.Symbol.Family.Name.ToLower().Contains("onwall"))
                    .ToList();
                var familySymbolSet = new HashSet<string>();
                foreach (var fi in allWallSleeveInstances)
                {
                    string fam = fi.Symbol.Family.Name;
                    string sym = fi.Symbol.Name;
                    familySymbolSet.Add($"Family: {fam}, Symbol: {sym}");
                }
                var allWallFamiliesDoc = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AllWallSleeveFamilies.txt");
                File.WriteAllLines(allWallFamiliesDoc, familySymbolSet.OrderBy(x => x));
                DebugLogger.Log($"[DEBUG] Wrote all wall sleeve families and symbols to {allWallFamiliesDoc}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DEBUG] Failed to write wall sleeve families: {ex.Message}");
            }

            double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
            double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
            double zTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);

            // Separate wall and slab sleeves, and group wall sleeves by type
            var wallSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol != null && fi.Symbol.Family != null &&
                    fi.Symbol.Family.Name.EndsWith("OnWall", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var slabSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol != null && fi.Symbol.Family != null &&
                    fi.Symbol.Family.Name.EndsWith("OnSlab", StringComparison.OrdinalIgnoreCase))
                .ToList();
            DebugLogger.Log($"Collected {wallSleeves.Count} wall sleeves, {slabSleeves.Count} slab sleeves. Tolerance: {toleranceMm:F1} mm");

            // Find rectangular opening symbols for each type
            var rectSymbols = new Dictionary<string, FamilySymbol>();
            var allRectSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();
            rectSymbols["Pipe"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("PipeOpeningOnWallRect", StringComparison.OrdinalIgnoreCase));
            rectSymbols["Duct"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("DuctOpeningOnWallRect", StringComparison.OrdinalIgnoreCase));
            rectSymbols["CableTray"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("CableTrayOpeningOnWallRect", StringComparison.OrdinalIgnoreCase));
            var slabRectSymbol = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.EndsWith("OnSlabRect", StringComparison.OrdinalIgnoreCase));
            if (rectSymbols["Pipe"] == null || rectSymbols["Duct"] == null || rectSymbols["CableTray"] == null)
            {
                TaskDialog.Show("Error", $"One or more rectangular cluster sleeve families (wall) not found. Pipe: {(rectSymbols["Pipe"] != null ? "OK" : "MISSING")}, Duct: {(rectSymbols["Duct"] != null ? "OK" : "MISSING")}, CableTray: {(rectSymbols["CableTray"] != null ? "OK" : "MISSING")}");
                return Result.Failed;
            }
            if (slabRectSymbol == null)
            {
                DebugLogger.Log("Rectangular cluster sleeve family (slab) not found. Slab clusters will be skipped.");
            }
            DebugLogger.Log($"Rectangular wall symbols: Pipe={rectSymbols["Pipe"]?.Name}, Duct={rectSymbols["Duct"]?.Name}, CableTray={rectSymbols["CableTray"]?.Name}, slab symbol: {(slabRectSymbol != null ? slabRectSymbol.Name : "NOT FOUND")}");

            using (var tx = new Transaction(doc, "Place Rectangular Cluster Sleeves"))
            {
                DebugLogger.Log("Starting transaction: Place Rectangular Cluster Sleeves");
                tx.Start();
                // Activate all wall rect symbols
                foreach (var sym in rectSymbols.Values)
                {
                    if (sym != null && !sym.IsActive) sym.Activate();
                }
                if (slabRectSymbol != null && !slabRectSymbol.IsActive)
                    slabRectSymbol.Activate();

                // --- WALL CLUSTER LOGIC ---
                var sleeveLocations = wallSleeves.ToDictionary(
                    s => s,
                    s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin);
                // Group wall sleeves by type (Pipe, Duct, CableTray)
                var wallTypeGroups = wallSleeves.GroupBy(s =>
                {
                    var fam = s.Symbol.Family.Name;
                    if (fam.Equals("PipeOpeningOnWall", StringComparison.OrdinalIgnoreCase)) return "Pipe";
                    if (fam.Equals("DuctOpeningOnWall", StringComparison.OrdinalIgnoreCase)) return "Duct";
                    if (fam.Equals("CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase)) return "CableTray";
                    return "Other";
                });
                int wallClusterCount = 0;
                foreach (var group in wallTypeGroups)
                {
                    if (group.Key == "Other") continue;
                    var groupSleeves = group.ToList();
                    var groupUnprocessed = new List<FamilyInstance>(groupSleeves);
                    var groupClusters = new List<List<FamilyInstance>>();
                    while (groupUnprocessed.Count > 0)
                    {
                        var queue = new Queue<FamilyInstance>();
                        var cluster = new List<FamilyInstance>();
                        queue.Enqueue(groupUnprocessed[0]);
                        groupUnprocessed.RemoveAt(0);
                        while (queue.Count > 0)
                        {
                            var inst = queue.Dequeue();
                            cluster.Add(inst);
                            var o1 = sleeveLocations[inst];
                            var neighbors = groupUnprocessed.Where(s =>
                            {
                                XYZ o2 = sleeveLocations[s];
                                if (Math.Abs(o1.Z - o2.Z) > zTolerance)
                                    return false;
                                double w1 = inst.LookupParameter("Width")?.AsDouble() ?? 0;
                                double w2 = s.LookupParameter("Width")?.AsDouble() ?? 0;
                                double h1 = inst.LookupParameter("Height")?.AsDouble() ?? 0;
                                double h2 = s.LookupParameter("Height")?.AsDouble() ?? 0;
                                double dx = o1.X - o2.X;
                                double dy = o1.Y - o2.Y;
                                double planar = Math.Sqrt(dx * dx + dy * dy);
                                double gap = planar - (w1 / 2.0 + w2 / 2.0 + h1 / 2.0 + h2 / 2.0) / 2.0;
                                return gap <= toleranceDist;
                            }).ToList();
                            foreach (var n in neighbors)
                            {
                                queue.Enqueue(n);
                                groupUnprocessed.Remove(n);
                            }
                        }
                        groupClusters.Add(cluster);
                    }
                    DebugLogger.Log($"Wall cluster formation complete for {group.Key}. Total clusters: {groupClusters.Count}");
                    foreach (var cluster in groupClusters)
                    {
                        if (cluster.Count < 2)
                            continue;
                        DebugLogger.Log($"Wall cluster of {cluster.Count} sleeves detected for replacement. Type: {group.Key}");
                        var clusterCenter = new XYZ(
                            cluster.Average(s => sleeveLocations[s].X),
                            cluster.Average(s => sleeveLocations[s].Y),
                            cluster.Average(s => sleeveLocations[s].Z)
                        );
                        // Suppression: Only check for this type's rectangular family
                        var existingRectOpenings = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name != null &&
                                   fi.Symbol.Family.Name.Equals(rectSymbols[group.Key].Family.Name, StringComparison.OrdinalIgnoreCase))
                            .Where(fi =>
                            {
                                var fiLocation = (fi.Location as LocationPoint)?.Point;
                                if (fiLocation == null) return false;
                                var distance = clusterCenter.DistanceTo(fiLocation);
                                return distance <= toleranceDist;
                            })
                            .ToList();
                        if (existingRectOpenings.Any())
                        {
                            DebugLogger.Log($"DUPLICATION SUPPRESSION: Skipping wall cluster at {clusterCenter} - found {existingRectOpenings.Count} existing rectangular opening(s) of type {group.Key} within {toleranceMm:F1}mm tolerance");
                            foreach (var existing in existingRectOpenings)
                            {
                                var existingLoc = (existing.Location as LocationPoint)?.Point;
                                string fam = existing.Symbol?.Family?.Name ?? "<null>";
                                string sym = existing.Symbol?.Name ?? "<null>";
                                int id = existing.Id.IntegerValue;
                                DebugLogger.Log($"  - Existing opening: Family={fam}, Symbol={sym}, ID={id}, Location={existingLoc}");
                            }
                            continue;
                        }
                        // Log all sleeve parameters in the cluster for debugging
                        DebugLogger.Log($"--- Cluster debug info ---");
                        int idx = 0;
                        foreach (var s in cluster)
                        {
                            var o = sleeveLocations[s];
                            string fam = s.Symbol?.Family?.Name ?? "<null>";
                            string type = s.Symbol?.Name ?? "<null>";
                            double w = s.LookupParameter("Width")?.AsDouble() ?? 0.0;
                            double h = s.LookupParameter("Height")?.AsDouble() ?? 0.0;
                            double d = s.LookupParameter("Depth")?.AsDouble() ?? 0.0;
                            DebugLogger.Log($"Sleeve {idx}: Id={s.Id.IntegerValue}, Family={fam}, Type={type}, Loc=({o.X:F2},{o.Y:F2},{o.Z:F2}), Width={UnitUtils.ConvertFromInternalUnits(w, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(h, UnitTypeId.Millimeters):F1}mm, Depth={UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters):F1}mm");
                            idx++;
                        }
                        // Compute extents
                        double xMinEdge = double.MaxValue, xMaxEdge = double.MinValue;
                        double yMinEdge = double.MaxValue, yMaxEdge = double.MinValue;
                        foreach (var s in cluster)
                        {
                            var o = sleeveLocations[s];
                            double w = s.LookupParameter("Width")?.AsDouble() ?? 0.0;
                            double h = s.LookupParameter("Height")?.AsDouble() ?? 0.0;
                            xMinEdge = Math.Min(xMinEdge, o.X - w / 2.0);
                            xMaxEdge = Math.Max(xMaxEdge, o.X + w / 2.0);
                            yMinEdge = Math.Min(yMinEdge, o.Y - h / 2.0);
                            yMaxEdge = Math.Max(yMaxEdge, o.Y + h / 2.0);
                        }
                        double midX = (xMinEdge + xMaxEdge) / 2.0;
                        double midY = (yMinEdge + yMaxEdge) / 2.0;
                        double midZ = cluster.Average(s => sleeveLocations[s].Z);
                        var midPoint = new XYZ(midX, midY, midZ);
                        double widthInternal = xMaxEdge - xMinEdge;
                        double heightInternal = yMaxEdge - yMinEdge;
                        var sleeveDepths = cluster.Select(s => s.LookupParameter("Depth")?.AsDouble() ?? 0.0);
                        double depthInternal = sleeveDepths.DefaultIfEmpty(0.0).Max();
                        var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
                        if (level == null)
                        {
                            DebugLogger.Log("Level not found, skipping wall cluster.");
                            continue;
                        }
                        // Aspect ratio logic for width/height swap
                        double finalWidth = widthInternal;
                        double finalHeight = heightInternal;
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
                        var rectSymbol = rectSymbols[group.Key];
                        if (!rectSymbol.IsActive) rectSymbol.Activate();
                        var rectInst = doc.Create.NewFamilyInstance(midPoint, rectSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        // Find intersecting structural framing element at the cluster location
                        double bTypeDepth = depthInternal;
                        try
                        {
                            var framingCollector = new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilyInstance))
                                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                .Cast<FamilyInstance>()
                                .ToList();
                            FamilyInstance intersectingFraming = null;
                            foreach (var framing in framingCollector)
                            {
                                var framingBox = framing.get_BoundingBox(null);
                                if (framingBox != null)
                                {
                                    if (framingBox.Min.X <= midPoint.X && midPoint.X <= framingBox.Max.X &&
                                        framingBox.Min.Y <= midPoint.Y && midPoint.Y <= framingBox.Max.Y &&
                                        framingBox.Min.Z <= midPoint.Z && midPoint.Z <= framingBox.Max.Z)
                                    {
                                        intersectingFraming = framing;
                                        break;
                                    }
                                }
                            }
                            if (intersectingFraming != null)
                            {
                                var symbol = doc.GetElement(intersectingFraming.GetTypeId()) as FamilySymbol;
                                var bParam = symbol?.LookupParameter("b");
                                if (bParam != null && bParam.StorageType == StorageType.Double)
                                {
                                    bTypeDepth = bParam.AsDouble();
                                    DebugLogger.Log($"[RectangularSleeveClusterCommand] Using 'b' type parameter from intersecting structural framing: {UnitUtils.ConvertFromInternalUnits(bTypeDepth, UnitTypeId.Millimeters):F1} mm");
                                }
                                else
                                {
                                    DebugLogger.Log("[RectangularSleeveClusterCommand] 'b' type parameter not found on intersecting structural framing, using fallback.");
                                }
                            }
                            else
                            {
                                DebugLogger.Log("[RectangularSleeveClusterCommand] No intersecting structural framing found, using fallback for wall thickness.");
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[RectangularSleeveClusterCommand] Exception while finding 'b' type parameter: {ex.Message}");
                        }
                        rectInst.LookupParameter("Depth")?.Set(bTypeDepth);
                        rectInst.LookupParameter("Width")?.Set(finalWidth);
                        rectInst.LookupParameter("Height")?.Set(finalHeight);
                        DebugLogger.Log($"Rectangular wall cluster sleeve placed at {midPoint} with width {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, height {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm, type: {group.Key}");
                        if (rectInst != null)
                        {
                            placedCount++;
                            wallClusterCount++;
                            DebugLogger.Log($"Rectangular wall cluster sleeve created with id {rectInst.Id.IntegerValue} (total placed: {placedCount})");
                            foreach (var s in cluster)
                            {
                                doc.Delete(s.Id);
                                deletedCount++;
                            }
                            DebugLogger.Log($"Deleted {cluster.Count} original wall sleeves (total deleted: {deletedCount})");
                        }
                    }
                }

                // --- SLAB CLUSTER LOGIC ---
                // --- SLAB CLUSTER LOGIC (by type) ---
                var slabRectSymbols = new Dictionary<string, FamilySymbol>();
                slabRectSymbols["Pipe"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("PipeOpeningOnSlabRect", StringComparison.OrdinalIgnoreCase));
                slabRectSymbols["Duct"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("DuctOpeningOnSlabRect", StringComparison.OrdinalIgnoreCase));
                slabRectSymbols["CableTray"] = allRectSymbols.FirstOrDefault(sym => sym.Family.Name.Equals("CableTrayOpeningOnSlabRect", StringComparison.OrdinalIgnoreCase));
                // Group slab sleeves by type
                var slabTypeGroups = slabSleeves.GroupBy(s =>
                {
                    var fam = s.Symbol.Family.Name;
                    if (fam.Equals("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase)) return "Pipe";
                    if (fam.Equals("DuctOpeningOnSlab", StringComparison.OrdinalIgnoreCase)) return "Duct";
                    if (fam.Equals("CableTrayOpeningOnSlab", StringComparison.OrdinalIgnoreCase)) return "CableTray";
                    return "Other";
                });
                foreach (var group in slabTypeGroups)
                {
                    if (group.Key == "Other") continue;
                    var groupSleeves = group.ToList();
                    DebugLogger.Log($"[SLAB] Processing group '{group.Key}' with {groupSleeves.Count} sleeves.");
                    var slabSleeveLocations = groupSleeves.ToDictionary(
                        s => s,
                        s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin);
                    var groupUnprocessed = new List<FamilyInstance>(groupSleeves);
                    var structuralClusters = new List<List<FamilyInstance>>();
                    while (groupUnprocessed.Count > 0)
                    {
                        var queue = new Queue<FamilyInstance>();
                        var cluster = new List<FamilyInstance>();
                        queue.Enqueue(groupUnprocessed[0]);
                        groupUnprocessed.RemoveAt(0);
                        while (queue.Count > 0)
                        {
                            var inst = queue.Dequeue();
                            cluster.Add(inst);
                            var o1 = slabSleeveLocations[inst];
                            var neighbors = groupUnprocessed.Where(s =>
                            {
                                XYZ o2 = slabSleeveLocations[s];
                                if (Math.Abs(o1.Z - o2.Z) > zTolerance)
                                    return false;
                                double dx = o1.X - o2.X;
                                double dy = o1.Y - o2.Y;
                                double planar = Math.Sqrt(dx * dx + dy * dy);
                                double w1 = inst.LookupParameter("Width")?.AsDouble() ?? 0;
                                double w2 = s.LookupParameter("Width")?.AsDouble() ?? 0;
                                double h1 = inst.LookupParameter("Height")?.AsDouble() ?? 0;
                                double h2 = s.LookupParameter("Height")?.AsDouble() ?? 0;
                                double gap = planar - (w1 / 2.0 + w2 / 2.0 + h1 / 2.0 + h2 / 2.0) / 2.0;
                                return gap <= toleranceDist;
                            }).ToList();
                            foreach (var n in neighbors)
                            {
                                queue.Enqueue(n);
                                groupUnprocessed.Remove(n);
                            }
                        }
                        // Always add cluster for replacement, regardless of floor's structural status
                        structuralClusters.Add(cluster);
                    }
                    DebugLogger.Log($"[SLAB] Cluster formation complete for {group.Key}. Total clusters: {structuralClusters.Count}");
                    int slabClusterIdx = 0;
                    foreach (var cluster in structuralClusters)
                    {
                        DebugLogger.Log($"[SLAB] Cluster {slabClusterIdx}: {cluster.Count} sleeves");
                        slabClusterIdx++;
                        if (cluster.Count < 2)
                        {
                            DebugLogger.Log($"[SLAB] Skipping cluster {slabClusterIdx - 1} (only {cluster.Count} sleeve(s))");
                            continue;
                        }
                        var clusterCenter = new XYZ(
                            cluster.Average(s => slabSleeveLocations[s].X),
                            cluster.Average(s => slabSleeveLocations[s].Y),
                            cluster.Average(s => slabSleeveLocations[s].Z)
                        );
                        DebugLogger.Log($"[SLAB] Cluster {slabClusterIdx - 1} center: {clusterCenter}");
                        var rectSymbol = slabRectSymbols[group.Key];
                        if (rectSymbol == null)
                        {
                            DebugLogger.Log($"[ERROR] No rectangular slab family found for {group.Key}, skipping cluster at {clusterCenter}");
                            continue;
                        }
                        if (!rectSymbol.IsActive) rectSymbol.Activate();
                        var existingRectOpenings = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(fi => fi.Symbol?.Family?.Name != null &&
                                   fi.Symbol.Family.Name.Equals(rectSymbol.Family.Name, StringComparison.OrdinalIgnoreCase))
                            .Where(fi =>
                            {
                                var fiLocation = (fi.Location as LocationPoint)?.Point;
                                if (fiLocation == null) return false;
                                var distance = clusterCenter.DistanceTo(fiLocation);
                                DebugLogger.Log($"[SLAB] Checking for duplicate: Existing {fi.Id} at {fiLocation}, distance={UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters):F1}mm");
                                return distance <= toleranceDist;
                            })
                            .ToList();
                        DebugLogger.Log($"[SLAB] Found {existingRectOpenings.Count} existing rectangular slab openings within tolerance at cluster center");
                        // Store the reference level of each sleeve in the cluster
                        var levelCounts = new Dictionary<ElementId, int>();
                        var levelRefs = new Dictionary<ElementId, Level>();
                        foreach (var s in cluster)
                        {
                            var refLevelParam = s.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                            if (refLevelParam != null && refLevelParam.StorageType == StorageType.ElementId)
                            {
                                var lvlId = refLevelParam.AsElementId();
                                if (lvlId != ElementId.InvalidElementId)
                                {
                                    if (!levelCounts.ContainsKey(lvlId))
                                        levelCounts[lvlId] = 0;
                                    levelCounts[lvlId]++;
                                    if (!levelRefs.ContainsKey(lvlId))
                                    {
                                        var lvl = doc.GetElement(lvlId) as Level;
                                        if (lvl != null)
                                            levelRefs[lvlId] = lvl;
                                    }
                                }
                            }
                        }
                        // Pick the most common level among the cluster sleeves
                        Level clusterLevel = null;
                        if (levelCounts.Count > 0)
                        {
                            var mostCommon = levelCounts.OrderByDescending(kvp => kvp.Value).First();
                            clusterLevel = levelRefs[mostCommon.Key];
                        }
                        else
                        {
                            // fallback: nearest below
                            var allLevels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                            clusterLevel = allLevels
                                .Where(lvl => lvl.Elevation <= clusterCenter.Z)
                                .OrderByDescending(lvl => lvl.Elevation)
                                .FirstOrDefault();
                        }
                        if (existingRectOpenings.Any())
                        {
                            DebugLogger.Log($"[SLAB] DUPLICATION SUPPRESSION: Skipping slab cluster at {clusterCenter} - found {existingRectOpenings.Count} existing rectangular slab opening(s) within {toleranceMm:F1}mm tolerance");
                            foreach (var existing in existingRectOpenings)
                            {
                                var existingLoc = (existing.Location as LocationPoint)?.Point;
                                DebugLogger.Log($"[SLAB]   - Existing opening: Family={existing.Symbol?.Family?.Name ?? "<null>"}, Symbol={existing.Symbol?.Name ?? "<null>"}, ID={existing.Id.IntegerValue}, Location={existingLoc}");
                            }
                            continue;
                        }
                        // Compute extents
                        double xMinEdge = double.MaxValue, xMaxEdge = double.MinValue;
                        double yMinEdge = double.MaxValue, yMaxEdge = double.MinValue;
                        foreach (var s in cluster)
                        {
                            var o = slabSleeveLocations[s];
                            double w = s.LookupParameter("Width")?.AsDouble() ?? 0.0;
                            double h = s.LookupParameter("Height")?.AsDouble() ?? 0.0;
                            xMinEdge = Math.Min(xMinEdge, o.X - w / 2.0);
                            xMaxEdge = Math.Max(xMaxEdge, o.X + w / 2.0);
                            yMinEdge = Math.Min(yMinEdge, o.Y - h / 2.0);
                            yMaxEdge = Math.Max(yMaxEdge, o.Y + h / 2.0);
                        }
                        double midX = (xMinEdge + xMaxEdge) / 2.0;
                        double midY = (yMinEdge + yMaxEdge) / 2.0;
                        double midZ = cluster.Average(s => slabSleeveLocations[s].Z);
                        var midPoint = new XYZ(midX, midY, midZ);
                        double widthInternal = xMaxEdge - xMinEdge;
                        double heightInternal = yMaxEdge - yMinEdge;
                        double widthMm = UnitUtils.ConvertFromInternalUnits(widthInternal, UnitTypeId.Millimeters);
                        double heightMm = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters);
                        double minDimMm = 20.0;
                        if (widthMm < minDimMm || heightMm < minDimMm)
                        {
                            DebugLogger.Log($"[SLAB][VALIDATION] Skipping placement: width={widthMm:F1}mm, height={heightMm:F1}mm below minimum {minDimMm}mm. Cluster center: {midPoint}");
                            continue;
                        }
                        DebugLogger.Log($"[SLAB] Placing rectangular opening at {midPoint} with width {widthMm:F1} mm, height {heightMm:F1} mm");
                        // The cluster is a set of individual sleeves to be replaced by a single cluster (rectangular) sleeve.
                        // Use only the max Depth from the individual sleeves in the cluster. No fallback to host floor.
                        var individualSleeveDepths = cluster.Select(s => s.LookupParameter("Depth")?.AsDouble() ?? 0.0);
                        double depthInternal = individualSleeveDepths.DefaultIfEmpty(0.0).Max();
                        DebugLogger.Log($"[SLAB] Calculated opening depth (from individual sleeves in cluster): {UnitUtils.ConvertFromInternalUnits(depthInternal, UnitTypeId.Millimeters):F1} mm");
                        // For Floor (slab), do not swap width/height
                        // No Level required for unhosted, work plane-based families. Always attempt placement if cluster is valid and family symbol exists.
                        FamilyInstance rectInst = null;
                        try
                        {
                            // IMPORTANT: For Generic Model, work plane-based families (slab sleeves), DO NOT pass a Level argument!
                            // Use the overload: NewFamilyInstance(XYZ, FamilySymbol, StructuralType.NonStructural)
                            rectInst = doc.Create.NewFamilyInstance(midPoint, rectSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            rectInst.LookupParameter("Depth")?.Set(depthInternal);
                            rectInst.LookupParameter("Width")?.Set(widthInternal);
                            rectInst.LookupParameter("Height")?.Set(heightInternal);
                            DebugLogger.Log($"[SLAB] Rectangular slab cluster sleeve placed at {midPoint} with width {UnitUtils.ConvertFromInternalUnits(widthInternal, UnitTypeId.Millimeters):F1} mm, height {UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters):F1} mm, type: {group.Key}");
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[SLAB][ERROR] Exception during rectangular opening placement: {ex.Message}");
                        }
                        bool instanceExists = rectInst != null && doc.GetElement(rectInst.Id) != null;
                        if (rectInst != null && instanceExists)
                        {
                            placedCount++;
                            DebugLogger.Log($"[SLAB] Rectangular slab cluster sleeve created with id {rectInst.Id.IntegerValue} (total placed: {placedCount})");
                            foreach (var s in cluster)
                            {
                                doc.Delete(s.Id);
                                deletedCount++;
                            }
                            DebugLogger.Log($"[SLAB] Deleted {cluster.Count} original slab sleeves (total deleted: {deletedCount})");
                        }
                        else if (rectInst != null && !instanceExists)
                        {
                            DebugLogger.Log($"[SLAB][ERROR] Rectangular opening instance {rectInst.Id.IntegerValue} was deleted by Revit after creation (likely family error).");
                        }
                    }
                }
                tx.Commit();
                DebugLogger.Log("Transaction committed for rectangular cluster sleeves.");
                DebugLogger.Log($"Summary: {placedCount} rectangular cluster sleeves placed, {deletedCount} original sleeves deleted.");
            }
            DebugLogger.Log("RectangularSleeveClusterCommand: Execute completed.");
            TaskDialog.Show("Done", $"Rectangular cluster sleeves placed: {placedCount}. Original sleeves deleted: {deletedCount}.");
            return Result.Succeeded;
        }
    }
}
