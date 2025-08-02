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
            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.EndsWith("OpeningOnWall", StringComparison.OrdinalIgnoreCase)
                          || fi.Symbol.Family.Name.EndsWith("OpeningOnSlab", StringComparison.OrdinalIgnoreCase))
                .ToList();

            using (var tx = new Transaction(doc, "Place Clustered Rectangular Openings V2"))
            {
                tx.Start();

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

                // For each group, form clusters based on edge-to-edge distance (≤ 100mm)
                Dictionary<SleeveGroupKey, List<List<FamilyInstance>>> clustersByGroup = new Dictionary<SleeveGroupKey, List<List<FamilyInstance>>>();
                foreach (var group in sleeveGroups)
                {
                    var familySleeves = group.ToList();
                    var groupLocations = familySleeves.ToDictionary(s => s, s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin);
                    var groupClusters = new List<List<FamilyInstance>>();
                    var unprocessedGroup = new List<FamilyInstance>(familySleeves);
                    while (unprocessedGroup.Any())
                    {
                        var queue = new Queue<FamilyInstance>();
                        var cluster = new List<FamilyInstance>();
                        queue.Enqueue(unprocessedGroup[0]);
                        unprocessedGroup.RemoveAt(0);
                        while (queue.Any())
                        {
                            var inst = queue.Dequeue();
                            cluster.Add(inst);
                            var o1_bbox = inst.get_BoundingBox(null);
                            var neighbors = unprocessedGroup.Where(s =>
                            {
                                if (o1_bbox == null) return false;
                                var o2_bbox = s.get_BoundingBox(null);
                                if (o2_bbox == null) return false;

                                // Check if bounding boxes are within tolerance
                                bool xOverlap = o1_bbox.Max.X >= o2_bbox.Min.X - toleranceDist && o1_bbox.Min.X <= o2_bbox.Max.X + toleranceDist;
                                bool yOverlap = o1_bbox.Max.Y >= o2_bbox.Min.Y - toleranceDist && o1_bbox.Min.Y <= o2_bbox.Max.Y + toleranceDist;
                                bool zOverlap = o1_bbox.Max.Z >= o2_bbox.Min.Z - toleranceDist && o1_bbox.Min.Z <= o2_bbox.Max.Z + toleranceDist;

                                return xOverlap && yOverlap && zOverlap;
                            }).ToList();
                            foreach (var n in neighbors)
                            {
                                queue.Enqueue(n);
                                unprocessedGroup.Remove(n);
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
                        Level refLevel = HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]);

                        


                        // --- Cluster sleeve duplicate suppression ---
                        double clusterSuppressionTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                        if (ClusterSleeveDuplicationService.IsClusterSleeveAtLocation(doc, mid, clusterSuppressionTol))
                        {
                            DebugLogger.Log($"Suppression: Existing cluster sleeve found within {UnitUtils.ConvertFromInternalUnits(clusterSuppressionTol, UnitTypeId.Millimeters):F0}mm at {mid}, skipping placement.");
                            continue;
                        }

                        // Always use bounding box center (mid) for placement, just like X-axis
                        // Place the cluster sleeve family instance at the cluster midpoint
                        FamilyInstance inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, refLevel, StructuralType.NonStructural);

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