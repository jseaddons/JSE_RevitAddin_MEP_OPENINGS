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
            
            var rectSymbol = allFamilySymbols
                .FirstOrDefault(sym => sym.Family.Name.Contains("PipeOpeningOnWallRect"));
            if (rectSymbol == null)
            {
                TaskDialog.Show("Error", "Please load sleeve opening families.");
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
                    DebugLogger.Log($"Processing sleeve {sleeve.Id.IntegerValue} for level assignment");
                    // Try to get reference level from parameter or helper
                    Level refLevel = HostLevelHelper.GetHostReferenceLevel(doc, sleeve);
                    if (refLevel == null)
                    {
                        var pt = (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin;
                        refLevel = GetNearestPositiveZLevel(doc, pt);
                        DebugLogger.Log($"Using nearest positive Z level for sleeve {sleeve.Id.IntegerValue}: {refLevel?.Name ?? "null"}");
                    }
                    else
                    {
                        DebugLogger.Log($"Got reference level from helper for sleeve {sleeve.Id.IntegerValue}: {refLevel.Name}");
                    }
                    if (refLevel != null)
                    {
                        Parameter levelParam = sleeve.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                            ?? sleeve.LookupParameter("Level")
                            ?? sleeve.LookupParameter("Reference Level");
                        if (levelParam != null && !levelParam.IsReadOnly && levelParam.StorageType == StorageType.ElementId)
                        {
                            levelParam.Set(refLevel.Id);
                            DebugLogger.Log($"Set Level parameter for sleeve {sleeve.Id.IntegerValue} to {refLevel.Name}");
                        }
                        // Set Schedule Level if available
                        var schedLevelParam = sleeve.LookupParameter("Schedule Level");
                        if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                        {
                            DebugLogger.Log($"Setting Schedule Level for sleeve {sleeve.Id.IntegerValue}, StorageType: {schedLevelParam.StorageType}");
                            if (schedLevelParam.StorageType == StorageType.ElementId)
                            {
                                schedLevelParam.Set(refLevel.Id);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.IntegerValue} Schedule Level to ElementId: {refLevel.Id.IntegerValue} ({refLevel.Name})");
                            }
                            else if (schedLevelParam.StorageType == StorageType.String)
                            {
                                schedLevelParam.Set(refLevel.Name);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.IntegerValue} Schedule Level to String: '{refLevel.Name}'");
                            }
                            else if (schedLevelParam.StorageType == StorageType.Integer)
                            {
                                schedLevelParam.Set(refLevel.Id.IntegerValue);
                                DebugLogger.Log($"Set sleeve {sleeve.Id.IntegerValue} Schedule Level to Integer: {refLevel.Id.IntegerValue} ({refLevel.Name})");
                            }
                        }
                        else
                        {
                            DebugLogger.Log($"Sleeve {sleeve.Id.IntegerValue} has no Schedule Level parameter or it's read-only");
                        }
                    }
                    else
                    {
                        DebugLogger.Log($"No reference level found for sleeve {sleeve.Id.IntegerValue}");
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
                    DebugLogger.Log($"About to process cluster with first sleeve ID: {cluster[0].Id.IntegerValue}");
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
                    DebugLogger.Log($"Planned rectangular opening at {midPoint} with width {widthMm:F1} mm, height {heightMm:F1} mm");

                    // DUPLICATE SUPPRESSION: Check for existing rectangular opening within 100mm (ignore size)
                    var existingRects = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol.Family.Name.Contains("PipeOpeningOnWallRect"))
                        .ToList();
                    bool duplicateFound = existingRects.Any(rect => {
                        var loc = (rect.Location as LocationPoint)?.Point ?? rect.GetTransform().Origin;
                        double dist = midPoint.DistanceTo(loc);
                        return dist <= suppressionTolerance;
                    });
                    if (duplicateFound)
                    {
                        DebugLogger.Log($"Suppressed duplicate rectangular opening at {midPoint} (existing rectangular opening within 100mm)");
                        continue;
                    }

                    // IMPORTANT: These are UNHOSTED Generic Model families placed in LINKED WALLS
                    // They do not have a direct Host property relationship to walls
                    // The families are placed independently and positioned near/through walls
                    // Wall orientation detection needs to be done through spatial analysis or user input

                    // calculate opening depth: use original sleeve family Depth parameter or fallback to wall thickness
                    // NOTE: wallHost will likely be null since these are unhosted Generic Model families
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
                    
                    DebugLogger.Log("DEBUG CHECKPOINT A: About to start Schedule Level processing...");
                    DebugLogger.Log("Starting Schedule Level processing...");
                    Level refLevel = null;
                    
                    DebugLogger.Log("DEBUG CHECKPOINT B: Entering try block for Schedule Level parameter lookup...");
                    try 
                    {
                        DebugLogger.Log($"About to get Schedule Level from cluster[0] with ID: {cluster[0].Id.IntegerValue}");
                        // Try to get Schedule Level from the first pipe in the cluster
                        var pipeSchedLevelParam = cluster[0].LookupParameter("Schedule Level");
                        DebugLogger.Log($"LookupParameter('Schedule Level') completed. Result: {(pipeSchedLevelParam != null ? "Found" : "NULL")}");
                        if (pipeSchedLevelParam != null)
                        {
                            DebugLogger.Log($"Pipe Schedule Level parameter found. StorageType: {pipeSchedLevelParam.StorageType}");
                            if (pipeSchedLevelParam.StorageType == StorageType.ElementId)
                            {
                                var schedLevelElem = doc.GetElement(pipeSchedLevelParam.AsElementId()) as Level;
                                if (schedLevelElem != null)
                                {
                                    refLevel = schedLevelElem;
                                    DebugLogger.Log($"Got Schedule Level from pipe ElementId: {schedLevelElem.Name} (Id: {schedLevelElem.Id.IntegerValue})");
                                }
                            }
                            else if (pipeSchedLevelParam.StorageType == StorageType.String)
                            {
                                string levelName = pipeSchedLevelParam.AsString();
                                DebugLogger.Log($"Pipe Schedule Level string value: '{levelName}'");
                                if (!string.IsNullOrEmpty(levelName))
                                {
                                    var schedLevelElem = new FilteredElementCollector(doc)
                                        .OfClass(typeof(Level))
                                        .Cast<Level>()
                                        .FirstOrDefault(l => l.Name == levelName);
                                    if (schedLevelElem != null)
                                    {
                                        refLevel = schedLevelElem;
                                        DebugLogger.Log($"Got Schedule Level from pipe string: {schedLevelElem.Name} (Id: {schedLevelElem.Id.IntegerValue})");
                                    }
                                    else
                                    {
                                        DebugLogger.Log($"Could not find level with name '{levelName}' in document");
                                    }
                                }
                            }
                        }
                        else
                        {
                            DebugLogger.Log("No 'Schedule Level' parameter found on pipe sleeve");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        DebugLogger.Log($"Exception during Schedule Level parameter processing: {ex.Message}");
                        DebugLogger.Log($"Exception Stack Trace: {ex.StackTrace}");
                    }
                    
                    DebugLogger.Log("DEBUG CHECKPOINT C: Completed Schedule Level parameter processing, about to determine final refLevel...");
                    if (refLevel == null)
                    {
                        DebugLogger.Log("No Schedule Level found on pipe, using fallback...");
                        var pt = (cluster[0].Location as LocationPoint)?.Point ?? cluster[0].GetTransform().Origin;
                        refLevel = GetNearestLevelBelowZ(doc, pt);
                        DebugLogger.Log($"[FALLBACK] Using nearest level strictly below Z for cluster[0] (ElementId: {cluster[0].Id.IntegerValue}): {refLevel?.Name ?? "null"} (Id: {refLevel?.Id.IntegerValue.ToString() ?? "null"})");
                    }
                    else
                    {
                        DebugLogger.Log($"Reference level for cluster[0] (ElementId: {cluster[0].Id.IntegerValue}) is: {refLevel.Name} (Id: {refLevel.Id.IntegerValue})");
                    }
                    
                    DebugLogger.Log("DEBUG CHECKPOINT D: About to create rectangular instance...");
                    DebugLogger.Log($"Creating rectangular instance with level: {refLevel?.Name ?? "null"}");
                    var rectInst = doc.Create.NewFamilyInstance(midPoint, rectSymbol, refLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    DebugLogger.Log("DEBUG CHECKPOINT E: Rectangular instance created successfully, about to set Level parameter...");
                    // Explicitly set the Level parameter to match the host reference level for schedule consistency
                    Parameter levelParam = rectInst.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    if (levelParam != null && !levelParam.IsReadOnly)
                    {
                        levelParam.Set(refLevel.Id);
                    }
                    else
                    {
                        // Fallback: try by name if built-in parameter is not available
                        var levelByName = rectInst.LookupParameter("Level");
                        if (levelByName != null && !levelByName.IsReadOnly)
                            levelByName.Set(refLevel.Id);
                    }
                    // Set Schedule Level parameter for schedule consistency
                    DebugLogger.Log("DEBUG CHECKPOINT F: About to set Schedule Level on rectangular opening...");
                    try 
                    {
                        DebugLogger.Log("About to set Schedule Level on rectangular opening...");
                        var schedLevelParam = rectInst.LookupParameter("Schedule Level");
                        DebugLogger.Log($"Rectangle LookupParameter('Schedule Level') completed. Result: {(schedLevelParam != null ? "Found" : "NULL")}");
                        if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                        {
                            DebugLogger.Log($"Rectangle Schedule Level parameter found. StorageType: {schedLevelParam.StorageType}");
                            if (schedLevelParam.StorageType == StorageType.ElementId)
                            {
                                schedLevelParam.Set(refLevel.Id);
                                DebugLogger.Log($"Set Rectangle Schedule Level to ElementId: {refLevel.Id.IntegerValue} ({refLevel.Name})");
                            }
                            else if (schedLevelParam.StorageType == StorageType.String)
                            {
                                schedLevelParam.Set(refLevel.Name);
                                DebugLogger.Log($"Set Rectangle Schedule Level to String: '{refLevel.Name}'");
                            }
                            else if (schedLevelParam.StorageType == StorageType.Integer)
                            {
                                schedLevelParam.Set(refLevel.Id.IntegerValue);
                                DebugLogger.Log($"Set Rectangle Schedule Level to Integer: {refLevel.Id.IntegerValue} ({refLevel.Name})");
                            }
                        }
                        else
                        {
                            DebugLogger.Log($"Rectangle Schedule Level parameter not found or is read-only");
                        }
                    }
                    catch (System.Exception ex) 
                    {
                        DebugLogger.Log($"Exception during Rectangle Schedule Level setting: {ex.Message}");
                        DebugLogger.Log($"Exception Stack Trace: {ex.StackTrace}");
                    }
                    // Wall orientation detection and dimension adjustment
                    double finalWidth = widthInternal;
                    double finalHeight = heightInternal;
                    DebugLogger.Log($"[DEBUG] width={UnitUtils.ConvertFromInternalUnits(widthInternal, UnitTypeId.Millimeters):F1} mm, height={UnitUtils.ConvertFromInternalUnits(heightInternal, UnitTypeId.Millimeters):F1} mm BEFORE swap logic");
                    
                    // Detect wall orientation to determine if width/height swap is needed
                    // For Y-axis walls (vertical walls running North-South), we need to swap width/height
                    // This is for sleeves placed for X-axis pipes (East-West pipes)
                    bool isYAxisWall = false;
                    
                    // Try to get wall host first (will be null for linked walls)
                    var wallHostForOrientation = cluster[0].Host as Wall;
                    DebugLogger.Log($"Wall host found: {wallHostForOrientation != null}");
                    
                    // Setup ReferenceIntersector to find walls at the cluster location if no host
                    Wall detectedWall = null;
                    if (wallHostForOrientation == null)
                    {
                        ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
                        var view3D = new FilteredElementCollector(doc)
                            .OfClass(typeof(View3D))
                            .Cast<View3D>()
                            .FirstOrDefault(v => !v.IsTemplate && v.ViewType == ViewType.ThreeD);

                        double bestProximity = double.MaxValue;
                        Wall bestWall = null;
                        double proximityThreshold = 0.1; // Loosened threshold (about 30mm)

                        if (view3D != null)
                        {
                            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
                            {
                                FindReferencesInRevitLinks = true
                            };

                            // Try rays from cluster center and from each sleeve center
                            List<XYZ> testPoints = new List<XYZ> { midPoint };
                            foreach (var s in cluster)
                            {
                                var loc = (s.Location as LocationPoint)?.Point ?? s.GetTransform().Origin;
                                testPoints.Add(loc);
                            }
                            XYZ[] directions = new[] { XYZ.BasisZ, XYZ.BasisZ.Negate(), XYZ.BasisX, XYZ.BasisY };
                            foreach (var pt in testPoints)
                            {
                                foreach (var dir in directions)
                                {
                                    var references = refIntersector.Find(pt, dir);
                                    if (references != null && references.Count > 0)
                                    {
                                        foreach (var refWithContext in references)
                                        {
                                            var refObj = refWithContext.GetReference();
                                            double prox = refWithContext.Proximity;
                                            DebugLogger.Log($"[WallDetect] Found reference at pt {pt} dir {dir} proximity {prox:F4}");
                                            if (prox < bestProximity && prox <= proximityThreshold)
                                            {
                                                Wall candidateWall = null;
                                                if (refObj.LinkedElementId != ElementId.InvalidElementId)
                                                {
                                                    var linkInstance = doc.GetElement(refObj.ElementId) as RevitLinkInstance;
                                                    if (linkInstance?.GetLinkDocument() != null)
                                                    {
                                                        candidateWall = linkInstance.GetLinkDocument().GetElement(refObj.LinkedElementId) as Wall;
                                                        DebugLogger.Log($"[WallDetect] Linked wall candidate: {candidateWall?.Id.IntegerValue ?? -1}");
                                                    }
                                                }
                                                else
                                                {
                                                    candidateWall = doc.GetElement(refObj.ElementId) as Wall;
                                                    DebugLogger.Log($"[WallDetect] Host wall candidate: {candidateWall?.Id.IntegerValue ?? -1}");
                                                }
                                                if (candidateWall != null)
                                                {
                                                    bestWall = candidateWall;
                                                    bestProximity = prox;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        detectedWall = bestWall;
                        DebugLogger.Log($"[WallDetect] Best wall found: {(detectedWall != null ? detectedWall.Id.IntegerValue.ToString() : "NONE")} with proximity {bestProximity:F4}");
                    }

                    DebugLogger.Log($"ReferenceIntersector detected wall: {detectedWall != null}");

                    Wall wallToUse = wallHostForOrientation ?? detectedWall;
                    if (wallToUse != null)
                    {
                        DebugLogger.Log($"Using wall ID: {wallToUse.Id.IntegerValue}");

                        // Use Wall.Orientation property to get wall normal/direction - the proven method from PipeSleeveCommand
                        XYZ wallOrientation = wallToUse.Orientation;
                        DebugLogger.Log($"Wall Orientation (normal): {wallOrientation}");

                        // Determine if wall is more aligned with X or Y axis by checking the orientation normal
                        // CORRECTED LOGIC: Y-axis walls (vertical, North-South) need width/height swapped
                        // Wall.Orientation gives the wall normal vector
                        bool isWallAlignedWithYAxis = Math.Abs(wallOrientation.X) > Math.Abs(wallOrientation.Y);
                        DebugLogger.Log($"Wall orientation analysis: Normal.X={wallOrientation.X:F3}, Normal.Y={wallOrientation.Y:F3}");
                        DebugLogger.Log($"Is wall aligned with Y-axis (runs vertically): {isWallAlignedWithYAxis}");

                        if (isWallAlignedWithYAxis)
                        {
                            // Wall runs vertically (Y-axis, North-South) - SWAP width/height for vertical walls
                            finalWidth = heightInternal;
                            finalHeight = widthInternal;
                            isYAxisWall = true;
                            DebugLogger.Log($"[DEBUG] Vertical wall detected. SWAPPED dimensions: finalWidth={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, finalHeight={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                        }
                        else
                        {
                            // Wall runs horizontally (X-axis, East-West) - keep original dimensions (these are working correctly)
                            DebugLogger.Log($"[DEBUG] Horizontal wall detected. KEEPING original dimensions: finalWidth={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, finalHeight={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                        }
                        DebugLogger.Log($"[DEBUG] Setting parameters: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
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
                            isYAxisWall = true;
                            DebugLogger.Log($"Vertical wall assumed from cluster aspect ratio. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                        }
                        else
                        {
                            // For nearly square clusters (0.67 <= ratio <= 1.5), default to vertical wall behavior
                            // This ensures consistent behavior when wall detection fails
                            finalWidth = heightInternal;
                            finalHeight = widthInternal;
                            isYAxisWall = true;
                            DebugLogger.Log($"Square-ish cluster - defaulting to vertical wall behavior. Swapped dimensions: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm");
                        }
                    }
                    
                    // Set rectangle parameters with final computed dimensions
                    rectInst.LookupParameter("Depth")?.Set(depthInternal);
                    rectInst.LookupParameter("Width")?.Set(finalWidth);
                    rectInst.LookupParameter("Height")?.Set(finalHeight);
                    DebugLogger.Log($"Set rectangular opening parameters: Width={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1} mm, Height={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1} mm, Depth={UnitUtils.ConvertFromInternalUnits(depthInternal, UnitTypeId.Millimeters):F1} mm");
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
