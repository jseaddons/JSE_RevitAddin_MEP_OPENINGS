using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class OpeningsPLaceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Initialize logging for the main command
            DebugLogger.InitLogFile("OpeningsPLaceCommand");
            DebugLogger.Log("OpeningsPLaceCommand: Execute started - running all sleeve placement commands");

            // Initialization code
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Call each sleeve placement in sequence
            DebugLogger.Log("Starting duct sleeve placement...");
            PlaceDuctSleeves(commandData, doc);

            DebugLogger.Log("Starting damper sleeve placement...");
            PlaceDamperSleeves(commandData, doc);

            DebugLogger.Log("Starting cable tray sleeve placement...");
            PlaceCableTraySleeves(commandData, doc);

            DebugLogger.Log("Starting pipe sleeve placement...");
            PlacePipeSleeves(commandData, doc);

            DebugLogger.Log("Starting rectangular pipe opening clustering...");
            PlaceRectangularPipeOpenings(commandData, doc);

            DebugLogger.Log("Starting rectangular sleeve clustering (final step)...");
            PlaceRectangularSleeveCluster(commandData, doc);

            DebugLogger.Log("OpeningsPLaceCommand: All sleeve commands completed successfully");
            DebugLogger.Log("Summary of completed operations:");
            DebugLogger.Log("- Duct sleeve placement");
            DebugLogger.Log("- Fire damper sleeve placement");
            DebugLogger.Log("- Cable tray sleeve placement");
            DebugLogger.Log("- Pipe sleeve placement");
            DebugLogger.Log("- Rectangular pipe opening clustering");
            DebugLogger.Log("- Rectangular sleeve clustering (final step)");
            DebugLogger.Log("Check individual log files for detailed results:");
            DebugLogger.Log("- DuctSleeveCommand.log, FireDamperPlaceCommand.log, CableTraySleeveCommand.log");
            DebugLogger.Log("- PipeSleeveCommand.log, PipeOpeningsRectCommand.log, RectangularSleeveClusterCommand.log");
            DebugLogger.Log("- OpeningsPLaceCommand.log");
            DebugLogger.Log("All processes completed automatically without user intervention.");
            return Result.Succeeded;
        }

        private void PlaceDuctSleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing DuctSleeveCommand...");
                var ductCommand = new DuctSleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = ductCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Duct sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Duct sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing duct sleeves: {ex.Message}");
            }
        }

        private void PlaceDamperSleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing FireDamperPlaceCommand...");
                var damperCommand = new FireDamperPlaceCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = damperCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Damper sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Damper sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing damper sleeves: {ex.Message}");
            }
        }

        private void PlaceCableTraySleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing CableTraySleeveCommand...");
                var cableTrayCommand = new CableTraySleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = cableTrayCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Cable tray sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Cable tray sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing cable tray sleeves: {ex.Message}");
            }
        }

        private void PlacePipeSleeves(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing PipeSleeveCommand...");
                var pipeCommand = new PipeSleeveCommand();
                string message = "";
                ElementSet elements = new ElementSet();
                var result = pipeCommand.Execute(commandData, ref message, elements);
                if (result != Result.Succeeded)
                {
                    DebugLogger.Log($"Pipe sleeve placement failed: {message}");
                }
                else
                {
                    DebugLogger.Log("Pipe sleeve placement completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing pipe sleeves: {ex.Message}");
            }
        }

        private void PlaceRectangularPipeOpenings(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing rectangular pipe opening clustering...");

                // Initialize logging for rectangular pipe openings
                DebugLogger.InitCustomLogFile("PipeOpeningsRectCommand");
                DebugLogger.Log("PipeOpeningsRectCommand: Execute started as part of OpeningsPLaceCommand.");

                int placedCount = 0;
                int deletedCount = 0;

                // Collect all placed pipe sleeves (assuming circular PS# family instances)
                double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
                // Only cluster pipes at same height (within 1 mm)
                double zTolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters);
                var sleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("PS#"))
                    .ToList();
                DebugLogger.Log($"Collected {sleeves.Count} pipe sleeves. Tolerance: {toleranceMm:F1} mm");

                // Find rectangular opening symbol
                var rectSymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(sym => sym.Family.Name.Contains("PipeOpeningOnWallRect"));
                if (rectSymbol == null)
                {
                    DebugLogger.Log("PipeOpeningOnWallRect family not found. Skipping rectangular clustering.");
                    return;
                }
                DebugLogger.Log($"Rectangular opening symbol found: {rectSymbol.Name}");

                using (var tx = new Transaction(doc, "Place Rectangular Pipe Openings"))
                {
                    DebugLogger.Log("Starting transaction: Place Rectangular Pipe Openings");
                    tx.Start();
                    // Activate the rectangular opening symbol within a valid transaction
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

                        // ...existing code... (compute extents, midpoint, etc.)
                        ProcessCluster(cluster, sleeveLocations, rectSymbol, doc, ref placedCount, ref deletedCount);
                    }

                    tx.Commit();
                    DebugLogger.Log("Transaction committed for rectangular openings.");
                    DebugLogger.Log($"Summary: {placedCount} rectangular openings placed, {deletedCount} circular sleeves deleted.");
                }

                DebugLogger.Log("Rectangular pipe opening clustering completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing rectangular pipe openings: {ex.Message}");
            }
        }

        private void ProcessCluster(List<FamilyInstance> cluster, Dictionary<FamilyInstance, XYZ> sleeveLocations,
            FamilySymbol rectSymbol, Document doc, ref int placedCount, ref int deletedCount)
        {
            // compute outer-edge extents of sleeves in XY
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

            // center point in 3D (Z from average)
            double midX = (xMinEdge + xMaxEdge) / 2.0;
            double midY = (yMinEdge + yMaxEdge) / 2.0;
            double midZ = cluster.Average(s => sleeveLocations[s].Z);
            var midPoint = new XYZ(midX, midY, midZ);

            // dimensions span outer edges
            double widthInternal = xMaxEdge - xMinEdge;
            double heightInternal = yMaxEdge - yMinEdge;
            double widthMm = UnitUtils.ConvertFromInternalUnits(widthInternal, UnitTypeId.Millimeters);
            double heightMm = UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters);
            DebugLogger.Log($"Placing rectangular opening at {midPoint} with width {widthMm:F1} mm, height {heightMm:F1} mm");

            // Calculate opening depth
            var wallHost = cluster[0].Host as Wall;
            double fallbackDepth = wallHost != null ? wallHost.Width : 0.0;
            if (wallHost != null)
                DebugLogger.Log($"Wall thickness for depth fallback: {UnitUtils.ConvertFromInternalUnits(fallbackDepth, UnitTypeId.Millimeters):F1} mm");
            else
                DebugLogger.Log("No wall host found - families are unhosted Generic Models in linked walls");

            // collect sleeve depths
            var sleeveDepths = cluster.Select(s => s.LookupParameter("Depth")?.AsDouble() ?? 0.0);
            // use max sleeve depth, or fallback if none
            double depthInternal = sleeveDepths.DefaultIfEmpty(fallbackDepth).Max();
            DebugLogger.Log($"Calculated opening depth: {UnitUtils.ConvertFromInternalUnits(depthInternal, UnitTypeId.Millimeters):F1} mm");

            // place on active level as work-plane based Generic Model
            var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            if (level == null)
            {
                DebugLogger.Log("Level not found, skipping cluster.");
                return;
            }

            // Wall orientation detection and dimension adjustment
            double finalWidth = widthInternal;
            double finalHeight = heightInternal;

            // Setup ReferenceIntersector to find walls at the cluster location
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            var view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);

            Wall detectedWall = null;
            if (view3D != null)
            {
                var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
                {
                    FindReferencesInRevitLinks = true
                };

                // Cast rays from cluster center to find intersecting wall
                XYZ[] rayDirections = { XYZ.BasisZ, XYZ.BasisZ.Negate(), XYZ.BasisX, XYZ.BasisY };

                foreach (var rayDirection in rayDirections)
                {
                    var references = refIntersector.Find(midPoint, rayDirection);
                    if (references != null && references.Count > 0)
                    {
                        var firstRefWithContext = references.First();
                        var firstRef = firstRefWithContext.GetReference();
                        if (firstRefWithContext.Proximity <= 0.01) // Very close to the point
                        {
                            if (firstRef.LinkedElementId != ElementId.InvalidElementId)
                            {
                                // This is a linked wall
                                var linkInstance = doc.GetElement(firstRef.ElementId) as RevitLinkInstance;
                                if (linkInstance?.GetLinkDocument() != null)
                                {
                                    detectedWall = linkInstance.GetLinkDocument().GetElement(firstRef.LinkedElementId) as Wall;
                                    DebugLogger.Log($"Found linked wall ID: {detectedWall?.Id.IntegerValue ?? -1}");
                                    break;
                                }
                            }
                            else
                            {
                                // This is a wall in the current document
                                detectedWall = doc.GetElement(firstRef.ElementId) as Wall;
                                DebugLogger.Log($"Found current document wall ID: {detectedWall?.Id.IntegerValue ?? -1}");
                                break;
                            }
                        }
                    }
                }
            }

            DebugLogger.Log($"Wall host found: {wallHost != null}");
            DebugLogger.Log($"ReferenceIntersector detected wall: {detectedWall != null}");

            Wall wallToUse = wallHost ?? detectedWall;
            if (wallToUse != null)
            {
                DebugLogger.Log($"Using wall ID: {wallToUse.Id.IntegerValue}");

                // Use Wall.Orientation property to get wall normal/direction
                XYZ wallOrientation = wallToUse.Orientation;
                DebugLogger.Log($"Wall Orientation (normal): {wallOrientation}");

                // Determine if wall is more aligned with X or Y axis by checking the orientation normal
                bool isWallAlignedWithXAxis = Math.Abs(wallOrientation.Y) > Math.Abs(wallOrientation.X);
                DebugLogger.Log($"Wall orientation analysis: Normal.X={wallOrientation.X:F3}, Normal.Y={wallOrientation.Y:F3}");
                DebugLogger.Log($"Is wall aligned with X-axis (runs horizontally): {isWallAlignedWithXAxis}");

                if (isWallAlignedWithXAxis)
                {
                    // Wall runs horizontally (X-axis) - keep original dimensions
                    DebugLogger.Log($"Horizontal wall detected. KEEPING original dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else
                {
                    // Wall runs vertically (Y-axis) - SWAP width/height for vertical walls
                    finalWidth = heightInternal;
                    finalHeight = widthInternal;
                    DebugLogger.Log($"Vertical wall detected. SWAPPED dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
            }
            else
            {
                // FALLBACK: Use aspect ratio analysis if no wall is found
                double aspectRatio = widthInternal / heightInternal;
                DebugLogger.Log($"No wall found. Using aspect ratio analysis: {aspectRatio:F2}");

                if (aspectRatio > 1.5) // Cluster is wider than tall (horizontal arrangement)
                {
                    // Horizontal arrangement - assume horizontal wall (X-axis) - keep original dimensions
                    DebugLogger.Log($"Horizontal wall assumed from cluster aspect ratio. Keeping original dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else if (aspectRatio < 0.67) // Cluster is taller than wide (vertical arrangement)
                {
                    // Vertical arrangement - assume vertical wall (Y-axis) - swap dimensions
                    finalWidth = heightInternal;
                    finalHeight = widthInternal;
                    DebugLogger.Log($"Vertical wall assumed from cluster aspect ratio. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else
                {
                    // For nearly square clusters (0.67 <= ratio <= 1.5), default to vertical wall behavior
                    finalWidth = heightInternal;
                    finalHeight = widthInternal;
                    DebugLogger.Log($"Square-ish cluster - defaulting to vertical wall behavior. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
            }

            var rectInst = doc.Create.NewFamilyInstance(midPoint, rectSymbol, level, StructuralType.NonStructural);
            // Set rectangle parameters with corrected width/height based on wall orientation
            rectInst.LookupParameter("Depth")?.Set(depthInternal);
            rectInst.LookupParameter("Width")?.Set(finalWidth);
            rectInst.LookupParameter("Height")?.Set(finalHeight);

            // Apply original sleeve orientation
            var origRot = (cluster[0].Location as LocationPoint)?.Rotation ?? 0;
            if (origRot != 0)
            {
                var axis = Line.CreateBound(midPoint, midPoint + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, rectInst.Id, axis, origRot);
            }
            DebugLogger.Log("Rectangular pipe opening placed with computed Width/Height/Depth and original rotation.");

            if (rectInst != null)
            {
                placedCount++;
                DebugLogger.Log($"Rectangular opening created with id {rectInst.Id.IntegerValue} (total placed: {placedCount})");
                // delete cluster sleeves
                foreach (var s in cluster)
                {
                    doc.Delete(s.Id);
                    deletedCount++;
                }
                DebugLogger.Log($"Deleted {cluster.Count} circular sleeves (total deleted: {deletedCount})");
            }
        }

        private void PlaceRectangularSleeveCluster(ExternalCommandData commandData, Document doc)
        {
            try
            {
                DebugLogger.Log("Executing rectangular sleeve clustering...");

                // Initialize logging for rectangular sleeve clustering
                DebugLogger.InitCustomLogFile("RectangularSleeveClusterCommand");
                DebugLogger.Log("RectangularSleeveClusterCommand: Execute started as part of OpeningsPLaceCommand.");

                int placedCount = 0;
                int deletedCount = 0;

                // Collect all placed rectangular sleeves (PS Rectangular family instances)
                double toleranceDist = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                double toleranceMm = UnitUtils.ConvertFromInternalUnits(toleranceDist, UnitTypeId.Millimeters);
                // Collect all placed rectangular sleeves for Duct, Damper, CableTray families
                var sleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.EndsWith("OpeningOnWall")
                        && (fi.Symbol.Name.StartsWith("DS#") || fi.Symbol.Name.StartsWith("DMS#") || fi.Symbol.Name.StartsWith("CT#")))
                    .ToList();
                DebugLogger.Log($"Collected {sleeves.Count} rectangular sleeves. Tolerance: {toleranceMm:F1} mm");

                if (sleeves.Count == 0)
                {
                    DebugLogger.Log("No rectangular sleeves found for clustering. Skipping.");
                    return;
                }

                // Log symbol variants for debug
                DebugLogger.Log("Sleeve types: " + string.Join(", ", sleeves.Select(s => s.Symbol.Name)));

                using (var tx = new Transaction(doc, "Place Clustered Rectangular Openings"))
                {
                    tx.Start();

                    // Build map of sleeve to insertion point
                    var sleeveLocations = sleeves.ToDictionary(
                        s => s,
                        s => (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin);

                    // Process clusters per symbol group (group by type name, not object identity)
                    var symbolGroups = sleeves.GroupBy(s => s.Symbol.Name);
                    foreach (var symbolGroup in symbolGroups)
                    {
                        string typeName = symbolGroup.Key;
                        var symbolSleeves = symbolGroup.ToList();
                        var clusterSymbol = symbolSleeves.First().Symbol;
                        DebugLogger.Log($"=== Begin processing type {typeName} ({symbolSleeves.Count} instances) ===");
                        DebugLogger.Log($"Clustering {symbolSleeves.Count} {clusterSymbol.Name} sleeves...");

                        // build per-group locations
                        var groupLocations = symbolSleeves.ToDictionary(s => s, s => sleeveLocations[s]);
                        var groupClusters = new List<List<FamilyInstance>>();
                        var unprocessedGroup = new List<FamilyInstance>(symbolSleeves);

                        // Build clusters for this symbol group
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
                                var o1 = groupLocations[inst];
                                var neighbors = unprocessedGroup.Where(s =>
                                {
                                    var o2 = groupLocations[s];
                                    double dx = o1.X - o2.X;
                                    double dy = o1.Y - o2.Y;
                                    double w1 = inst.LookupParameter("Width")?.AsDouble() ?? 0;
                                    double h1 = inst.LookupParameter("Height")?.AsDouble() ?? 0;
                                    double w2 = s.LookupParameter("Width")?.AsDouble() ?? 0;
                                    double h2 = s.LookupParameter("Height")?.AsDouble() ?? 0;
                                    double xGap = Math.Max(0, Math.Abs(dx) - (w1 / 2 + w2 / 2));
                                    double yGap = Math.Max(0, Math.Abs(dy) - (h1 / 2 + h2 / 2));
                                    double edgeDist = Math.Sqrt(xGap * xGap + yGap * yGap);
                                    double edgeMm = UnitUtils.ConvertFromInternalUnits(edgeDist, UnitTypeId.Millimeters);
                                    DebugLogger.Log($"[{clusterSymbol.Name}] Compare {inst.Id}->{s.Id}: edge gap {edgeMm:F1} mm");
                                    return edgeDist <= toleranceDist;
                                }).ToList();
                                foreach (var n in neighbors)
                                {
                                    queue.Enqueue(n);
                                    unprocessedGroup.Remove(n);
                                }
                            }
                            groupClusters.Add(cluster);
                            DebugLogger.Log($"Formed cluster of {cluster.Count} sleeves for {clusterSymbol.Name}: IDs {string.Join(",", cluster.Select(s => s.Id))}");
                        }

                        DebugLogger.Log($"{clusterSymbol.Name}: Total clusters formed: {groupClusters.Count}");

                        // Process each cluster for this symbol (only clusters with 2+ sleeves)
                        foreach (var cluster in groupClusters)
                        {
                            if (cluster.Count < 2)
                                continue;

                            // Process the cluster using the helper method
                            ProcessRectangularCluster(cluster, groupLocations, clusterSymbol, doc, ref placedCount, ref deletedCount);
                        }

                        DebugLogger.Log($"=== End processing symbol {clusterSymbol.Name} ===");
                    }

                    DebugLogger.Log($"All symbol groups processed. Summary: {placedCount} openings placed, {deletedCount} sleeves deleted.");
                    tx.Commit();
                }

                DebugLogger.Log("Rectangular sleeve clustering completed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error placing rectangular sleeve clusters: {ex.Message}");
            }
        }

        private void ProcessRectangularCluster(List<FamilyInstance> cluster, Dictionary<FamilyInstance, XYZ> groupLocations,
            FamilySymbol clusterSymbol, Document doc, ref int placedCount, ref int deletedCount)
        {
            // compute extents
            double xMin = double.MaxValue, xMax = double.MinValue;
            double yMin = double.MaxValue, yMax = double.MinValue;
            double zSum = 0;
            foreach (var s in cluster)
            {
                var o = groupLocations[s];
                double w = s.LookupParameter("Width")?.AsDouble() ?? 0;
                double h = s.LookupParameter("Height")?.AsDouble() ?? 0;
                xMin = Math.Min(xMin, o.X - w / 2);
                xMax = Math.Max(xMax, o.X + w / 2);
                yMin = Math.Min(yMin, o.Y - h / 2);
                yMax = Math.Max(yMax, o.Y + h / 2);
                zSum += o.Z;
            }
            var mid = new XYZ((xMin + xMax) / 2, (yMin + yMax) / 2, zSum / cluster.Count);
            double width = xMax - xMin;
            double height = yMax - yMin;
            double depthInternal = cluster.Max(s => s.LookupParameter("Depth")?.AsDouble() ?? 0);

            // place combined opening
            if (!clusterSymbol.IsActive)
                clusterSymbol.Activate();
            var level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().FirstOrDefault();
            var inst = doc.Create.NewFamilyInstance(mid, clusterSymbol, level, StructuralType.NonStructural);

            // Apply original orientation
            var origRot = (cluster[0].Location as LocationPoint)?.Rotation ?? 0;
            if (origRot != 0)
            {
                // rotate around vertical axis through midpoint
                var axis = Line.CreateBound(
                    new XYZ(mid.X, mid.Y, mid.Z - 10),
                    new XYZ(mid.X, mid.Y, mid.Z + 10));
                ElementTransformUtils.RotateElement(doc, inst.Id, axis, origRot);
            }

            // Wall orientation detection and dimension adjustment
            double finalWidth = width;
            double finalHeight = height;

            // Try to get wall host first
            var wallHost = cluster[0].Host as Wall;
            DebugLogger.Log($"Wall host found: {wallHost != null}");

            // Setup ReferenceIntersector to find walls at the cluster location if no host
            Wall detectedWall = null;
            if (wallHost == null)
            {
                ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
                var view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);

                if (view3D != null)
                {
                    var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
                    {
                        FindReferencesInRevitLinks = true
                    };

                    // Cast rays from cluster center to find intersecting wall
                    XYZ[] rayDirections = { XYZ.BasisZ, XYZ.BasisZ.Negate(), XYZ.BasisX, XYZ.BasisY };

                    foreach (var rayDirection in rayDirections)
                    {
                        var references = refIntersector.Find(mid, rayDirection);
                        if (references != null && references.Count > 0)
                        {
                            var firstRefWithContext = references.First();
                            var firstRef = firstRefWithContext.GetReference();
                            if (firstRefWithContext.Proximity <= 0.01) // Very close to the point
                            {
                                if (firstRef.LinkedElementId != ElementId.InvalidElementId)
                                {
                                    // This is a linked wall
                                    var linkInstance = doc.GetElement(firstRef.ElementId) as RevitLinkInstance;
                                    if (linkInstance?.GetLinkDocument() != null)
                                    {
                                        detectedWall = linkInstance.GetLinkDocument().GetElement(firstRef.LinkedElementId) as Wall;
                                        DebugLogger.Log($"Found linked wall ID: {detectedWall?.Id.IntegerValue ?? -1}");
                                        break;
                                    }
                                }
                                else
                                {
                                    // This is a wall in the current document
                                    detectedWall = doc.GetElement(firstRef.ElementId) as Wall;
                                    DebugLogger.Log($"Found current document wall ID: {detectedWall?.Id.IntegerValue ?? -1}");
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            DebugLogger.Log($"ReferenceIntersector detected wall: {detectedWall != null}");

            Wall wallToUse = wallHost ?? detectedWall;
            if (wallToUse != null)
            {
                DebugLogger.Log($"Using wall ID: {wallToUse.Id.IntegerValue}");

                // Use Wall.Orientation property to get wall normal/direction - the proven method from PipeSleeveCommand
                XYZ wallOrientation = wallToUse.Orientation;
                DebugLogger.Log($"Wall Orientation (normal): {wallOrientation}");

                // Determine if wall is more aligned with X or Y axis by checking the orientation normal
                bool isWallAlignedWithXAxis = Math.Abs(wallOrientation.Y) > Math.Abs(wallOrientation.X);
                DebugLogger.Log($"Wall orientation analysis: Normal.X={wallOrientation.X:F3}, Normal.Y={wallOrientation.Y:F3}");
                DebugLogger.Log($"Is wall aligned with X-axis (runs horizontally): {isWallAlignedWithXAxis}");

                if (isWallAlignedWithXAxis)
                {
                    // Wall runs horizontally (X-axis) - keep original dimensions
                    DebugLogger.Log($"Horizontal wall detected. KEEPING original dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else
                {
                    // Wall runs vertically (Y-axis) - SWAP width/height for vertical walls
                    finalWidth = height;
                    finalHeight = width;
                    DebugLogger.Log($"Vertical wall detected. SWAPPED dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
            }
            else
            {
                // FALLBACK: Use aspect ratio analysis if no wall is found
                double aspectRatio = width / height;
                DebugLogger.Log($"No wall found. Using aspect ratio analysis: {aspectRatio:F2}");

                if (aspectRatio > 1.5) // Cluster is wider than tall (horizontal arrangement)
                {
                    // Horizontal arrangement - assume horizontal wall (X-axis) - keep original dimensions
                    DebugLogger.Log($"Horizontal wall assumed from cluster aspect ratio. Keeping original dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else if (aspectRatio < 0.67) // Cluster is taller than wide (vertical arrangement)
                {
                    // Vertical arrangement - assume vertical wall (Y-axis) - swap dimensions
                    finalWidth = height;
                    finalHeight = width;
                    DebugLogger.Log($"Vertical wall assumed from cluster aspect ratio. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
                else
                {
                    // For nearly square clusters (0.67 <= ratio <= 1.5), default to vertical wall behavior
                    // This ensures consistent behavior when wall detection fails
                    finalWidth = height;
                    finalHeight = width;
                    DebugLogger.Log($"Square-ish cluster - defaulting to vertical wall behavior. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                }
            }

            // Use corrected extents: apply wall orientation logic
            inst.LookupParameter("Width")?.Set(finalWidth);
            inst.LookupParameter("Height")?.Set(finalHeight);
            inst.LookupParameter("Depth")?.Set(depthInternal);
            DebugLogger.Log($"Placed cluster opening {inst.Id} of type {clusterSymbol.Name} width {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm height {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm depth {UnitUtils.ConvertFromInternalUnits(depthInternal, UnitTypeId.Millimeters):F1} mm");
            placedCount++;

            // delete originals
            foreach (var s in cluster)
            {
                var sid = s.Id;
                doc.Delete(sid);
                deletedCount++;
                DebugLogger.Log($"Deleted original sleeve {sid}");
            }
        }
    }
}
