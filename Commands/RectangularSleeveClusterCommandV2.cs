using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class RectangularSleeveClusterCommandV2 : IExternalCommand
    {
        // Helper struct for grouping key
        struct SleeveGroupKey
        {
            public string hostType;
            public string systemType;
            public string orientation;
            public SleeveGroupKey(string hostType, string systemType, string orientation)
            {
                this.hostType = hostType;
                this.systemType = systemType;
                this.orientation = orientation;
            }
            // Override Equals and GetHashCode for dictionary/grouping
            public override bool Equals(object obj)
            {
                if (!(obj is SleeveGroupKey)) return false;
                var other = (SleeveGroupKey)obj;
                return hostType == other.hostType && systemType == other.systemType && orientation == other.orientation;
            }
            public override int GetHashCode()
            {
                return (hostType, systemType, orientation).GetHashCode();
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string logFileName = $"RectangularSleeveClusterCommandV2_{timestamp}.log";
            try { DebugLogger.InitLogFile(logFileName); }
            catch (Exception ex) { TaskDialog.Show("Logging Error", $"Failed to initialize log file: {logFileName}\n{ex.Message}"); }
            DebugLogger.Log($"Assembly version: {Assembly.GetExecutingAssembly().GetName().Version}");
            DebugLogger.Log($"Assembly build timestamp: {System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location):o}");
            DebugLogger.Log("*** RectangularSleeveClusterCommandV2 START ***");
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            int placedCount = 0;
            int deletedCount = 0;

            double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
            double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);

            // Collect all placed rectangular sleeves (PS Rectangular family instances)
            var rawSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.EndsWith("OpeningOnWall", StringComparison.OrdinalIgnoreCase)
                          || fi.Symbol.Family.Name.EndsWith("OpeningOnSlab", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Use SectionBoxHelper to reduce to only elements visible in the active 3D section box
            List<FamilyInstance> sleeves;
            try
            {
                var rawElements = rawSleeves.Cast<Element>().Select(e => (element: (Element)e, transform: (Transform?)null)).ToList();
                var filtered = SectionBoxHelper.FilterElementsBySectionBox(uiDoc, rawElements);
                sleeves = filtered.Select(t => t.element).OfType<FamilyInstance>().ToList();
                DebugLogger.Log($"[RectangularCluster] Raw sleeves={rawSleeves.Count}, Filtered by section box={sleeves.Count}");
                // If section-box filtering returns zero but we had raw sleeves, assume the section
                // box is positioned such that nothing passed the filter. Fall back to raw collection
                // to avoid silently doing nothing when the user expects clustering.
                if (sleeves.Count == 0 && rawSleeves.Count > 0)
                {
                    DebugLogger.Log($"[RectangularCluster] Section-box filtering yielded 0 results; falling back to raw collection of {rawSleeves.Count} sleeves.");
                    sleeves = rawSleeves;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RectangularCluster] SectionBox filtering failed: {ex.Message}; falling back to raw collection");
                sleeves = rawSleeves;
            }

            using (var tx = new Transaction(doc, "Place Clustered Rectangular Openings V2"))
            {
                tx.Start();

                // Debug targets: user-provided sleeve ids that must cluster
                var debugTargetIds = new HashSet<int> { 775507, 775509, 775510 };

                var sleeveGroups = sleeves.GroupBy(sleeve => {
                    // Use HostOrientation parameter only
                    var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
                    string effectiveOrientation = hostOrientationParam != null ? hostOrientationParam.AsString() : "";
                    // HostType from family name
                    var famName = sleeve.Symbol.Family.Name.ToLower();
                    string hostType = famName.Contains("wall") ? "Wall" : (famName.Contains("slab") || famName.Contains("floor") ? "Floor" : "Unknown");
                    // SystemType from family name (now includes pipe)
                    string systemType = famName.Contains("duct") ? "Duct" 
                                        : famName.Contains("cabletray") ? "CableTray" 
                                        : famName.Contains("pipe") ? "Pipe"
                                        : "Unknown";
                    return new SleeveGroupKey(hostType, systemType, effectiveOrientation);
                });

                // Diagnostic: log group counts and sample ids to understand why no clusters form
                try
                {
                    foreach (var g in sleeveGroups)
                    {
                        var list = g.ToList();
                        var ids = list.Select(fi => fi.Id.IntegerValue.ToString()).Take(10).ToList();
                        DebugLogger.Log($"[ClusterDiag] Group hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[ClusterDiag] Failed to log sleeveGroups: {ex.Message}");
                }

                // For each group, form clusters based on edge-to-edge distance (≤ 100mm)
                Dictionary<SleeveGroupKey, List<List<FamilyInstance>>> clustersByGroup = new Dictionary<SleeveGroupKey, List<List<FamilyInstance>>>();
                foreach (var group in sleeveGroups)
                {
                    // Build a list of family sleeves and precompute centers and bboxes once
                    var familySleeves = group.ToList();
                    var centers = new Dictionary<FamilyInstance, XYZ>(familySleeves.Count);
                    var bboxes = new Dictionary<FamilyInstance, BoundingBoxXYZ>(familySleeves.Count);
                    foreach (var s in familySleeves)
                    {
                        var center = (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin;
                        centers[s] = center;
                        try { var bb = s.get_BoundingBox(null); if (bb != null) bboxes[s] = bb; } catch { }
                    }

                    // Targeted diagnostic: if this group contains any of the target IDs, dump center and bbox info
                    try
                    {
                        if (familySleeves.Any(fi => debugTargetIds.Contains(fi.Id.IntegerValue)))
                        {
                            DebugLogger.Log($"[TARGET-DIAG] Group contains target IDs ({string.Join(",", familySleeves.Where(fi => debugTargetIds.Contains(fi.Id.IntegerValue)).Select(fi => fi.Id.IntegerValue))}). Dumping centers and bboxes:");
                            foreach (var fi in familySleeves)
                            {
                                int id = fi.Id.IntegerValue;
                                var c = centers.ContainsKey(fi) ? centers[fi] : (fi.Location as LocationPoint)?.Point ?? fi.GetTransform().Origin;
                                BoundingBoxXYZ? bb = null;
                                try { bb = bboxes.ContainsKey(fi) ? bboxes[fi] : fi.get_BoundingBox(null); } catch { }
                                if (bb != null)
                                {
                                    DebugLogger.Log($"[TARGET-DIAG] ID={id}, Center=({c.X:F4},{c.Y:F4},{c.Z:F4}), BBoxMin=({bb.Min.X:F4},{bb.Min.Y:F4},{bb.Min.Z:F4}), BBoxMax=({bb.Max.X:F4},{bb.Max.Y:F4},{bb.Max.Z:F4})");
                                }
                                else
                                {
                                    DebugLogger.Log($"[TARGET-DIAG] ID={id}, Center=({c.X:F4},{c.Y:F4},{c.Z:F4}), BBox=NULL");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[TARGET-DIAG] Failed to write target diagnostics: {ex.Message}");
                    }

                    // Spatial hash grid using toleranceDist as cell size to limit neighbor searches
                    var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();
                    double cellSize = toleranceDist > 0.0 ? toleranceDist : UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    // Build grid by bounding-box extents (so border proximity is respected)
                    foreach (var s in familySleeves)
                    {
                        BoundingBoxXYZ? bb = null;
                        try { bb = bboxes.ContainsKey(s) ? bboxes[s] : s.get_BoundingBox(null); } catch { }
                        if (bb != null)
                        {
                            int min_ix = (int)Math.Floor(bb.Min.X / cellSize);
                            int max_ix = (int)Math.Floor(bb.Max.X / cellSize);
                            int min_iy = (int)Math.Floor(bb.Min.Y / cellSize);
                            int max_iy = (int)Math.Floor(bb.Max.Y / cellSize);
                            int min_iz = (int)Math.Floor(bb.Min.Z / cellSize);
                            int max_iz = (int)Math.Floor(bb.Max.Z / cellSize);
                            for (int gx = min_ix; gx <= max_ix; gx++)
                                for (int gy = min_iy; gy <= max_iy; gy++)
                                    for (int gz = min_iz; gz <= max_iz; gz++)
                                    {
                                        var key = (gx, gy, gz);
                                        if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                                        list.Add(s);
                                    }
                        }
                        else
                        {
                            // Fallback to center-based bucketing when bbox unavailable
                            var c = centers[s];
                            int ix = (int)Math.Floor(c.X / cellSize);
                            int iy = (int)Math.Floor(c.Y / cellSize);
                            int iz = (int)Math.Floor(c.Z / cellSize);
                            var key = (ix, iy, iz);
                            if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                            list.Add(s);
                        }
                    }

                    var groupClusters = new List<List<FamilyInstance>>();
                    var unprocessedSet = new HashSet<FamilyInstance>(familySleeves);

                    while (unprocessedSet.Count > 0)
                    {
                        var start = unprocessedSet.First();
                        var queue = new Queue<FamilyInstance>();
                        var cluster = new List<FamilyInstance>();
                        queue.Enqueue(start);
                        unprocessedSet.Remove(start);

                        while (queue.Count > 0)
                        {
                            var inst = queue.Dequeue();
                            cluster.Add(inst);

                            // Find candidate neighbors using buckets overlapped by the instance bbox expanded by tolerance
                            BoundingBoxXYZ o1_bbox = bboxes.ContainsKey(inst) ? bboxes[inst] : inst.get_BoundingBox(null);
                            var candidates = new List<FamilyInstance>();
                            if (o1_bbox != null)
                            {
                                double exMinX = o1_bbox.Min.X - toleranceDist;
                                double exMaxX = o1_bbox.Max.X + toleranceDist;
                                double exMinY = o1_bbox.Min.Y - toleranceDist;
                                double exMaxY = o1_bbox.Max.Y + toleranceDist;
                                double exMinZ = o1_bbox.Min.Z - toleranceDist;
                                double exMaxZ = o1_bbox.Max.Z + toleranceDist;

                                int min_ix = (int)Math.Floor(exMinX / cellSize);
                                int max_ix = (int)Math.Floor(exMaxX / cellSize);
                                int min_iy = (int)Math.Floor(exMinY / cellSize);
                                int max_iy = (int)Math.Floor(exMaxY / cellSize);
                                int min_iz = (int)Math.Floor(exMinZ / cellSize);
                                int max_iz = (int)Math.Floor(exMaxZ / cellSize);

                                for (int gx = min_ix; gx <= max_ix; gx++)
                                    for (int gy = min_iy; gy <= max_iy; gy++)
                                        for (int gz = min_iz; gz <= max_iz; gz++)
                                        {
                                            var key = (gx, gy, gz);
                                            if (grid.TryGetValue(key, out var bucket)) candidates.AddRange(bucket);
                                        }
                            }
                            else
                            {
                                // Fallback to center-neighbor search (previous 3x3x3)
                                var c = centers[inst];
                                int ix = (int)Math.Floor(c.X / cellSize);
                                int iy = (int)Math.Floor(c.Y / cellSize);
                                int iz = (int)Math.Floor(c.Z / cellSize);
                                for (int dx = -1; dx <= 1; dx++)
                                    for (int dy = -1; dy <= 1; dy++)
                                        for (int dz = -1; dz <= 1; dz++)
                                        {
                                            var key = (ix + dx, iy + dy, iz + dz);
                                            if (grid.TryGetValue(key, out var bucket)) candidates.AddRange(bucket);
                                        }
                            }

                            // Check actual bbox overlap with tolerance only for candidates
                            // o1_bbox already computed above for candidate search; reuse it
                            var neighbors = new List<FamilyInstance>();
                            foreach (var s in candidates)
                            {
                                if (!unprocessedSet.Contains(s)) continue;
                                if (s == inst) continue;
                                var o2_bbox = bboxes.ContainsKey(s) ? bboxes[s] : s.get_BoundingBox(null);
                                if (o1_bbox == null || o2_bbox == null) continue;

                                bool xOverlap = o1_bbox.Max.X >= o2_bbox.Min.X - toleranceDist && o1_bbox.Min.X <= o2_bbox.Max.X + toleranceDist;
                                bool yOverlap = o1_bbox.Max.Y >= o2_bbox.Min.Y - toleranceDist && o1_bbox.Min.Y <= o2_bbox.Max.Y + toleranceDist;
                                bool zOverlap = o1_bbox.Max.Z >= o2_bbox.Min.Z - toleranceDist && o1_bbox.Min.Z <= o2_bbox.Max.Z + toleranceDist;
                                if (xOverlap && yOverlap && zOverlap) neighbors.Add(s);
                            }

                            foreach (var n in neighbors)
                            {
                                if (unprocessedSet.Remove(n)) queue.Enqueue(n);
                            }
                        }

                        groupClusters.Add(cluster);
                    }

                    clustersByGroup[group.Key] = groupClusters;
                }

                // Process each cluster within each group
                foreach (var groupEntry in clustersByGroup)
                {
                    var groupKey = groupEntry.Key;
                    var clusters = groupEntry.Value;

                    foreach (var cluster in clusters)
                    {
                            try
                            {
                                var sampleIds = cluster.Select(fi => fi.Id.IntegerValue.ToString()).Take(10).ToList();
                                DebugLogger.Log($"[ClusterDiag] Processing cluster size={cluster.Count}, sampleIds={string.Join(",", sampleIds)}");
                            }
                            catch { }

                        // Skip pipe clusters on Wall or Structural Framing only (let PipeOpeningsRectCommand handle them)
                        if (groupKey.systemType == "Pipe" && (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing"))
                            continue;
                    
                    
                        if (cluster.Count <= 1) continue; // Skip individual sleeves and empty clusters
                        
                        string familyName = "";
                        // Always use ClusterOpeningOnWallX family for Wall and Structural Framing clusters
                        if ((groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing"))
                        {
                            familyName = "ClusterOpeningOnWallX";
                        }
                        else if (groupKey.hostType == "Floor")
                        {
                            familyName = "ClusterOpeningOnSlab";
                        }
                        else
                        {
                            DebugLogger.Log($"Unknown host/orientation for cluster group, skipping. HostType={groupKey.hostType}, Orientation={groupKey.orientation}");
                            continue;
                        }

                        // Detect if this is a pipe cluster (all sleeves are pipe system)
                        bool isPipeCluster = groupKey.systemType == "Pipe";

                        // Find a FamilySymbol for this system type and host type
                        var allClusterSymbols = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                            .ToList();
                        if (allClusterSymbols.Count == 0)
                        {
                            DebugLogger.Log($"No suitable cluster family found for name '{familyName}'. Please load the required family and try again.");
                            TaskDialog.Show("Missing Family", $"No suitable cluster family found for name '{familyName}'. Please load the required family and try again.");
                            continue;
                        }
                        var clusterSymbol = allClusterSymbols.First();
                        if (!clusterSymbol.IsActive) clusterSymbol.Activate();

                        // Use ClusterBoundingBoxServices to get bounding box and midpoint
                        var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(cluster);

                        // Use reference level from first sleeve in cluster
                        Level? refLevel = HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]);
                        if (refLevel == null)
                        {
                            DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                            continue;
                        }

                        


                        // --- Cluster sleeve duplicate suppression ---
                        double clusterSuppressionTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                        // Use section-box and hostType aware cluster bounds check to avoid scanning all clusters
                        BoundingBoxXYZ? sectionBoxForDoc = null;
                        try { if (uiDoc.ActiveView is View3D vb2) sectionBoxForDoc = SectionBoxHelper.GetSectionBoxBounds(vb2); } catch { }
                        string clusterHostType = groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing" ? "ClusterOpeningOnWallX" : "ClusterOpeningOnSlab";
                        bool clusterExists = OpeningDuplicationChecker.IsLocationWithinClusterBounds(doc, mid, clusterSuppressionTol, hostType: clusterHostType, sectionBox: sectionBoxForDoc);
                        if (clusterExists)
                        {
                            DebugLogger.Log($"Suppression: Existing cluster sleeve found within {UnitUtils.ConvertFromInternalUnits(clusterSuppressionTol, UnitTypeId.Millimeters):F0}mm at {mid}, skipping placement. (optimized)");
                            continue;
                        }

                        // Always use bounding box center (mid) for placement, just like X-axis
                        // Place the cluster sleeve family instance at the cluster midpoint
                        FamilyInstance inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, refLevel!, StructuralType.NonStructural);

                        // Determine if rotation is needed based on host and orientation
                        double rotationAngle = 0.0;
                        if (groupKey.hostType == "Wall")
                        {
                            if (groupKey.orientation == "Y")
                            {
                                // Rotate 90 degrees around vertical axis for Y-oriented walls
                                rotationAngle = Math.PI / 2;
                            }
                            // For X-oriented walls, no rotation needed (0 radians)
                        }

                        // Apply rotation if needed
                        if (rotationAngle != 0.0)
                        {
                            XYZ axisOrigin = mid;
                            XYZ axisDirection = XYZ.BasisZ; // vertical axis

                            Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                            ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
                        }

                        // Now set Width, Height, Depth parameters mapping model bbox dimensions correctly
                        var widthParam = inst.LookupParameter("Width");
                        var heightParam = inst.LookupParameter("Height");
                        var depthParam = inst.LookupParameter("Depth");

                        if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                        {
                            // Map to match family created in Left view:
                            // Width parameter = height (Y), Height parameter = depth (Z), Depth parameter = width (X)
                            if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(height);   // Y
                            if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(depth); // Z

                            // For Depth, use host thickness if available (mimic CableTraySleevePlacer), but mapped to X
                            double hostThickness = width;
                            if (groupKey.hostType == "Wall")
                            {
                                var wall = cluster[0].Host as Wall;
                                if (wall != null)
                                {
                                    hostThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                                    DebugLogger.Log($"[CLUSTER] Wall thickness used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                                }
                            }
                            else if (groupKey.hostType == "Structural Framing")
                            {
                                var framing = cluster[0].Host as FamilyInstance;
                                if (framing != null)
                                {
                                    var framingType = framing.Symbol;
                                    var bParam = framingType.LookupParameter("b");
                                    if (bParam != null && bParam.StorageType == StorageType.Double)
                                    {
                                        hostThickness = bParam.AsDouble();
                                        DebugLogger.Log($"[CLUSTER] Framing 'b' parameter used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                                    }
                                }
                            }
                            if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(hostThickness); // X
                        }
                        else
                        {
                            // For other hosts (e.g., Floor), use bounding box values as before
                            if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(width);
                            if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(height);
                            if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(depth);
                        }
                        placedCount++;
                        // Debug logging for cluster placement
                        var placementPoints = cluster.Select(s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin).ToList();
                        DebugLogger.Log($"[CLUSTER-DEBUG] Cluster placement points:");
                        foreach (var pt in placementPoints)
                            DebugLogger.Log($"[CLUSTER-DEBUG]   - ({pt.X:F3}, {pt.Y:F3}, {pt.Z:F3})");
                        DebugLogger.Log($"[CLUSTER-DEBUG] BoundingBox Min=({cluster.Min(s => s.get_BoundingBox(null)?.Min.X ?? 0):F3}, {cluster.Min(s => s.get_BoundingBox(null)?.Min.Y ?? 0):F3}, {cluster.Min(s => s.get_BoundingBox(null)?.Min.Z ?? 0):F3})");
                        DebugLogger.Log($"[CLUSTER-DEBUG] BoundingBox Max=({cluster.Max(s => s.get_BoundingBox(null)?.Max.X ?? 0):F3}, {cluster.Max(s => s.get_BoundingBox(null)?.Max.Y ?? 0):F3}, {cluster.Max(s => s.get_BoundingBox(null)?.Max.Z ?? 0):F3})");
                        DebugLogger.Log($"[CLUSTER-DEBUG] BoundingBox center=({mid.X:F3}, {mid.Y:F3}, {mid.Z:F3})");
                        var avgPt = new XYZ(placementPoints.Average(p => p.X), placementPoints.Average(p => p.Y), placementPoints.Average(p => p.Z));
                        DebugLogger.Log($"[CLUSTER-DEBUG] Average placement point=({avgPt.X:F3}, {avgPt.Y:F3}, {avgPt.Z:F3})");
                        DebugLogger.Log($"[CLUSTER-DEBUG] Center-Avg delta=({mid.X - avgPt.X:F3}, {mid.Y - avgPt.Y:F3}, {mid.Z - avgPt.Z:F3})");

                        // Delete originals
                        foreach (var s in cluster)
                        {
                            var sid = s.Id;
                            doc.Delete(sid);
                            deletedCount++;
                        }
                    }
                }
                tx.Commit();
            }
            DebugLogger.Log($"All symbol groups processed. Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
            return Result.Succeeded;
        }
    }
}
