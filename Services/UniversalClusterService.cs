using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Universal clustering service - extracted from RectangularSleeveClusterCommandV2
    /// Can be called from any context (ICommand, IExternalCommand, etc.)
    /// </summary>
    public class UniversalClusterService
    {
        // Helper struct for grouping key
        private struct SleeveGroupKey
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

        /// <summary>
        /// Cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="targetCategory">Category to cluster (e.g., "Ducts", "Pipes") or null for all</param>
        /// <param name="uiDoc">Optional UIDocument for section box filtering</param>
        /// <returns>Tuple of (placedCount, deletedCount)</returns>
        public (int placedCount, int deletedCount) ClusterSleeves(Document doc, string targetCategory, UIDocument uiDoc = null)
        {
            int placedCount = 0;
            int deletedCount = 0;

            try
            {
                // Get cluster configuration
                double toleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance;
                double toleranceDist = UnitUtils.ConvertToInternalUnits(toleranceMm, UnitTypeId.Millimeters);
                
                DebugLogger.Log($"[UniversalClusterService] Using JoinOpeningsDistance: {toleranceMm}mm (from {ClusterConfigurationManager.Instance.ConfigurationSource})");
                DebugLogger.Log($"[UniversalClusterService] Internal units: {toleranceDist:F6} feet");
                DebugLogger.Log($"[UniversalClusterService] Target category: {targetCategory ?? "ALL"}");

                // Collect all sleeves using the 4 universal families
                var allSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => 
                    {
                        var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                        return famName == "RectangularOpeningOnWall" ||
                               famName == "CircularOpeningOnWall" ||
                               famName == "RectangularOpeningOnSlab" ||
                               famName == "CircularOpeningOnSlab";
                    })
                    .ToList();
                
                DebugLogger.Log($"[UniversalClusterService] Found {allSleeves.Count} total sleeves (all categories)");
                
                // Filter by MEP_Category parameter if targetCategory is specified
                var rawSleeves = string.IsNullOrEmpty(targetCategory)
                    ? allSleeves
                    : allSleeves.Where(sleeve => 
                    {
                        var categoryParam = sleeve.LookupParameter("MEP_Category");
                        string sleeveCategory = categoryParam?.AsString() ?? "";
                        return sleeveCategory == targetCategory;
                    }).ToList();
                
                DebugLogger.Log($"[UniversalClusterService] Filtered to {rawSleeves.Count} sleeves" + 
                               (string.IsNullOrEmpty(targetCategory) ? " (all categories)" : $" for category '{targetCategory}'"));

                // Use SectionBoxHelper to reduce to only elements visible in the active 3D section box (if UIDocument provided)
                List<FamilyInstance> sleeves;
                if (uiDoc != null)
                {
                    try
                    {
                        var rawElements = rawSleeves.Cast<Element>().Select(e => (element: (Element)e, transform: (Transform?)null)).ToList();
                        var filtered = SectionBoxHelper.FilterElementsBySectionBox(uiDoc, rawElements);
                        sleeves = filtered.Select(t => t.element).OfType<FamilyInstance>().ToList();
                        DebugLogger.Log($"[UniversalClusterService] Raw sleeves={rawSleeves.Count}, Filtered by section box={sleeves.Count}");
                        
                        // Fallback to raw collection if section box filtering yields zero results
                        if (sleeves.Count == 0 && rawSleeves.Count > 0)
                        {
                            DebugLogger.Log($"[UniversalClusterService] Section-box filtering yielded 0 results; falling back to raw collection of {rawSleeves.Count} sleeves.");
                            sleeves = rawSleeves;
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[UniversalClusterService] SectionBox filtering failed: {ex.Message}; falling back to raw collection");
                        sleeves = rawSleeves;
                    }
                }
                else
                {
                    sleeves = rawSleeves;
                    DebugLogger.Log($"[UniversalClusterService] No UIDocument provided, skipping section box filtering");
                }

                if (sleeves.Count == 0)
                {
                    DebugLogger.Log($"[UniversalClusterService] No sleeves to cluster");
                    return (0, 0);
                }

                // ⚠️ CRITICAL: Transaction must be started by caller
                if (!doc.IsModifiable)
                {
                    DebugLogger.Error($"[UniversalClusterService] Document is not modifiable - transaction must be started by caller");
                    throw new InvalidOperationException("Document must be in a transaction before calling ClusterSleeves");
                }

                // Group sleeves by host type, system type, and orientation
                var sleeveGroups = sleeves.GroupBy(sleeve => {
                    var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
                    string effectiveOrientation = hostOrientationParam != null ? hostOrientationParam.AsString() : "";
                    
                    var famName = sleeve.Symbol.Family.Name.ToLower();
                    string hostType = famName.Contains("onwall") ? "Wall" : (famName.Contains("onslab") || famName.Contains("onfloor") ? "Floor" : "Unknown");
                    
                    var categoryParam = sleeve.LookupParameter("MEP_Category");
                    string systemType = categoryParam?.AsString() ?? "Unknown";
                    
                    return new SleeveGroupKey(hostType, systemType, effectiveOrientation);
                });

                // Log group diagnostics
                foreach (var g in sleeveGroups)
                {
                    var list = g.ToList();
                    var ids = list.Select(fi => fi.Id.IntegerValue.ToString()).Take(10).ToList();
                    DebugLogger.Log($"[ClusterService] Group hostType={g.Key.hostType}, systemType={g.Key.systemType}, orientation={g.Key.orientation}, count={list.Count}, sampleIds={string.Join(",", ids)}");
                }

                // Form clusters using spatial hashing
                var clustersByGroup = FormClusters(sleeveGroups, toleranceDist);

                // Process each cluster
                foreach (var groupEntry in clustersByGroup)
                {
                    var groupKey = groupEntry.Key;
                    var clusters = groupEntry.Value;

                    foreach (var cluster in clusters)
                    {
                        if (cluster.Count <= 1) continue; // Skip individual sleeves

                        try
                        {
                            // Place cluster sleeve
                            PlaceClusterSleeve(doc, cluster, groupKey, out int placed1, out int deleted1);
                            placedCount += placed1;
                            deletedCount += deleted1;
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Error($"[UniversalClusterService] Error placing cluster: {ex.Message}");
                        }
                    }
                }

                DebugLogger.Log($"[UniversalClusterService] Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterService] Clustering failed: {ex.Message}");
                DebugLogger.Error($"[UniversalClusterService] Stack trace: {ex.StackTrace}");
                throw;
            }

            return (placedCount, deletedCount);
        }

        private Dictionary<SleeveGroupKey, List<List<FamilyInstance>>> FormClusters(
            IEnumerable<IGrouping<SleeveGroupKey, FamilyInstance>> sleeveGroups, 
            double toleranceDist)
        {
            var clustersByGroup = new Dictionary<SleeveGroupKey, List<List<FamilyInstance>>>();
            double cellSize = toleranceDist > 0.0 ? toleranceDist : UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);

            foreach (var group in sleeveGroups)
            {
                var familySleeves = group.ToList();
                var centers = new Dictionary<FamilyInstance, XYZ>(familySleeves.Count);
                var bboxes = new Dictionary<FamilyInstance, BoundingBoxXYZ>(familySleeves.Count);

                // Precompute centers and bboxes
                foreach (var s in familySleeves)
                {
                    var center = (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin;
                    centers[s] = center;
                    try { var bb = s.get_BoundingBox(null); if (bb != null) bboxes[s] = bb; } catch { }
                }

                // Build spatial hash grid
                var grid = BuildSpatialGrid(familySleeves, bboxes, centers, cellSize);

                // Form clusters using BFS
                var groupClusters = FormClustersFromGrid(familySleeves, grid, bboxes, centers, cellSize, toleranceDist);
                clustersByGroup[group.Key] = groupClusters;
            }

            return clustersByGroup;
        }

        private Dictionary<(int, int, int), List<FamilyInstance>> BuildSpatialGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize)
        {
            var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();

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
                    // Fallback to center-based bucketing
                    var c = centers[s];
                    int ix = (int)Math.Floor(c.X / cellSize);
                    int iy = (int)Math.Floor(c.Y / cellSize);
                    int iz = (int)Math.Floor(c.Z / cellSize);
                    var key = (ix, iy, iz);
                    if (!grid.TryGetValue(key, out var list)) { list = new List<FamilyInstance>(); grid[key] = list; }
                    list.Add(s);
                }
            }

            return grid;
        }

        private List<List<FamilyInstance>> FormClustersFromGrid(
            List<FamilyInstance> familySleeves,
            Dictionary<(int, int, int), List<FamilyInstance>> grid,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            Dictionary<FamilyInstance, XYZ> centers,
            double cellSize,
            double toleranceDist)
        {
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

                    BoundingBoxXYZ o1_bbox = bboxes.ContainsKey(inst) ? bboxes[inst] : inst.get_BoundingBox(null);
                    var candidates = GetCandidatesFromGrid(inst, o1_bbox, centers, grid, cellSize, toleranceDist);
                    var neighbors = FilterNeighborsByBoundingBox(inst, candidates, o1_bbox, bboxes, unprocessedSet, toleranceDist);

                    foreach (var n in neighbors)
                    {
                        if (unprocessedSet.Remove(n)) queue.Enqueue(n);
                    }
                }

                groupClusters.Add(cluster);
            }

            return groupClusters;
        }

        private List<FamilyInstance> GetCandidatesFromGrid(
            FamilyInstance inst,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, XYZ> centers,
            Dictionary<(int, int, int), List<FamilyInstance>> grid,
            double cellSize,
            double toleranceDist)
        {
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
                // Fallback to 3x3x3 neighbor search
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

            return candidates;
        }

        private List<FamilyInstance> FilterNeighborsByBoundingBox(
            FamilyInstance inst,
            List<FamilyInstance> candidates,
            BoundingBoxXYZ o1_bbox,
            Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
            HashSet<FamilyInstance> unprocessedSet,
            double toleranceDist)
        {
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

            return neighbors;
        }

        private void PlaceClusterSleeve(
            Document doc,
            List<FamilyInstance> cluster,
            SleeveGroupKey groupKey,
            out int placed,
            out int deleted)
        {
            placed = 0;
            deleted = 0;

            // Determine if cluster is circular or rectangular
            bool isCircular = cluster.All(s => 
            {
                var fam = s.Symbol?.Family?.Name ?? "";
                return fam.Contains("Circular");
            });
            
            // Select universal family based on host type and shape
            string familyName = "";
            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                familyName = isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
            }
            else if (groupKey.hostType == "Floor")
            {
                familyName = isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
            }
            else
            {
                DebugLogger.Log($"Unknown host type for cluster group, skipping. HostType={groupKey.hostType}");
                return;
            }

            DebugLogger.Log($"[ClusterService] Creating {groupKey.systemType} cluster using family '{familyName}' (shape: {(isCircular ? "Circular" : "Rectangular")})");

            // Find FamilySymbol
            var allClusterSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (allClusterSymbols.Count == 0)
            {
                DebugLogger.Log($"No suitable cluster family found for name '{familyName}'");
                return;
            }

            var clusterSymbol = allClusterSymbols.First();
            if (!clusterSymbol.IsActive) clusterSymbol.Activate();

            // Get bounding box and midpoint
            var (width, height, depth, mid) = ClusterBoundingBoxServices.GetClusterBoundingBox(cluster);

            // Get reference level
            Level? refLevel = HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]);
            if (refLevel == null)
            {
                DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                return;
            }

            // Place cluster sleeve
            FamilyInstance inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, refLevel!, StructuralType.NonStructural);

            // Apply rotation if needed
            double rotationAngle = 0.0;
            if (groupKey.hostType == "Wall" && groupKey.orientation == "Y")
            {
                rotationAngle = Math.PI / 2;
            }

            if (rotationAngle != 0.0)
            {
                XYZ axisOrigin = mid;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);
            }

            // Set size parameters
            SetClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth);

            placed++;

            // Delete originals
            foreach (var s in cluster)
            {
                doc.Delete(s.Id);
                deleted++;
            }
        }

        private void SetClusterSizeParameters(
            Document doc,
            FamilyInstance inst,
            List<FamilyInstance> cluster,
            SleeveGroupKey groupKey,
            double width,
            double height,
            double depth)
        {
            var widthParam = inst.LookupParameter("Width");
            var heightParam = inst.LookupParameter("Height");
            var depthParam = inst.LookupParameter("Depth");

            if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
            {
                // Map to match family created in Left view
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(height);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(depth);

                // Use host thickness for Depth
                double hostThickness = width;
                if (groupKey.hostType == "Wall")
                {
                    var wall = cluster[0].Host as Wall;
                    if (wall != null)
                    {
                        hostThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                        DebugLogger.Log($"[ClusterService] Wall thickness used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
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
                            DebugLogger.Log($"[ClusterService] Framing 'b' parameter used for Depth: {UnitUtils.ConvertFromInternalUnits(hostThickness, UnitTypeId.Millimeters):F1}mm");
                        }
                    }
                }
                if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(hostThickness);
            }
            else
            {
                // For other hosts (Floor), use bounding box values
                if (widthParam != null && !widthParam.IsReadOnly) widthParam.Set(width);
                if (heightParam != null && !heightParam.IsReadOnly) heightParam.Set(height);
                if (depthParam != null && !depthParam.IsReadOnly) depthParam.Set(depth);
            }
        }
    }
}

