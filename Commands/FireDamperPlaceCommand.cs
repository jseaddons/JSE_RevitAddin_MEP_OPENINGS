using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;


namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class FireDamperPlaceCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument? uiDoc = null;
            Document? doc = null;
            try
            {
                DebugLogger.Log("=== [FireDamperPlaceCommand] DEBUG START ===");
                // TEMP: Comment out DamperLogger.InitLogFile() to fix logger conflict
                // DamperLogger.InitLogFile();
                DebugLogger.Log("[FireDamperPlaceCommand] After InitLogFile");

                uiDoc = commandData.Application.ActiveUIDocument;
                doc = uiDoc.Document;
                DebugLogger.Log("[FireDamperPlaceCommand] Got UIDocument and Document");
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[FireDamperPlaceCommand] EXCEPTION: {ex.Message}\n{ex.StackTrace}");
                Autodesk.Revit.UI.TaskDialog.Show("FireDamperPlaceCommand Exception", ex.ToString());
                throw;
            }

            // Auto-select the opening family symbol
            var openingSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("DamperOpeningOnWall"));
            DebugLogger.Log(openingSymbol == null ? "[FireDamperPlaceCommand] openingSymbol is NULL" : $"[FireDamperPlaceCommand] Found openingSymbol: {openingSymbol.Family.Name} - {openingSymbol.Name}");

            if (openingSymbol == null)
            {
                DebugLogger.Log("[FireDamperPlaceCommand] openingSymbol is null, showing error dialog and exiting");
                TaskDialog.Show("Error", "Please load sleeve opening families.");
                return Result.Failed;
            }

            // Activate the symbol
            using (var txActivate = new Transaction(doc, "Activate Opening Symbol"))
            {
                txActivate.Start();
                if (!openingSymbol.IsActive)
                {
                    DebugLogger.Log("[FireDamperPlaceCommand] Activating openingSymbol");
                    openingSymbol.Activate();
                }
                txActivate.Commit();
            }
            DebugLogger.Log("[FireDamperPlaceCommand] Opening symbol activated");

            // Collect dampers (host and linked) - use direct collection instead of MEP-only helper
            var damperTuples = new List<(FamilyInstance, Transform?)>();
            
            // Host dampers
            var hostDampers = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("Damper"))
                .ToList();
            damperTuples.AddRange(hostDampers.Select(d => (d, (Transform?)null)));
            DebugLogger.Log($"[FireDamperPlaceCommand] hostDampers count: {hostDampers.Count}");
            
            // Linked dampers
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();
            DebugLogger.Log($"[FireDamperPlaceCommand] linkInstances count: {linkInstances.Count}");
            foreach (var linkInstance in linkInstances)
            {
                // Skip hidden links (category or instance hidden in active view)
                if (doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(doc.ActiveView))
                {
                    DebugLogger.Log($"[FireDamperPlaceCommand] Skipping hidden linkInstance: {linkInstance.Name}");
                    continue;
                }
                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) continue;
                var linkTransform = linkInstance.GetTotalTransform();
                var linkedDampers = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("Damper"))
                    .ToList();
                damperTuples.AddRange(linkedDampers.Select(d => (d, (Transform?)linkTransform)));
                DebugLogger.Log($"[FireDamperPlaceCommand] linkedDampers count: {linkedDampers.Count} for linkInstance {linkInstance.Name}");
            }
            
            DebugLogger.Log($"[FireDamperPlaceCommand] damperTuples total count: {damperTuples.Count}");

            // Collect walls (host and linked) as tuples (wall, transform)
            var wallTuples = new List<(Wall, Transform?)>();
            // Host walls
            var hostWalls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();
            wallTuples.AddRange(hostWalls.Select(w => (w, (Transform?)null)));
            DebugLogger.Log($"[FireDamperPlaceCommand] hostWalls count: {hostWalls.Count}");
            // Linked walls
            foreach (var linkInstance in linkInstances)
            {
                // Skip hidden links (category or instance hidden in active view)
                if (doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(doc.ActiveView))
                {
                    DebugLogger.Log($"[FireDamperPlaceCommand] Skipping hidden linkInstance for walls: {linkInstance.Name}");
                    continue;
                }
                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) { DebugLogger.Log("[FireDamperPlaceCommand] linkInstance with null linkedDoc"); continue; }
                var linkTransform = linkInstance.GetTotalTransform();
                var linkedWalls = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
                wallTuples.AddRange(linkedWalls.Select(w => (w, (Transform?)linkTransform)));
                DebugLogger.Log($"[FireDamperPlaceCommand] linkedWalls count: {linkedWalls.Count} for linkInstance {linkInstance.Name}");
            }

            // SUPPRESSION: Collect existing damper sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("DamperOpeningOnWall") ||
                           (fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("DMS#")))
                .ToList();
            DebugLogger.Log($"[FireDamperPlaceCommand] existingSleeves count: {existingSleeves.Count}");

            // Log existing sleeve details for debugging
            foreach (var sleeve in existingSleeves.Take(5)) // Log first 5 for debugging
            {
                var location = (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin;
                DebugLogger.Log($"[FireDamperPlaceCommand] Existing sleeve: Family={sleeve.Symbol.Family.Name}, Symbol={sleeve.Symbol.Name}, Location={location}");
            }

            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin
            );

            int placedCount = 0;
            int skippedExistingCount = 0; // Counter for dampers with existing sleeves
            int totalDampers = damperTuples.Count;

            DebugLogger.Log($"[FireDamperPlaceCommand] Entered Execute. damperTuples count: {damperTuples.Count}");

            using (var tx = new Transaction(doc, "Place Fire Damper Sleeves"))
            {
                tx.Start();
                foreach (var tuple in damperTuples)
                {
                    var damper = tuple.Item1;
                    var damperTransform = tuple.Item2;

                    DebugLogger.Log($"[FireDamperPlaceCommand] Processing damper id={damper.Id}, Family={damper.Symbol.Family.Name}, Symbol={damper.Symbol.Name}");

                    // 1. Get damper location (host coords)
                    XYZ damperLoc = (damper.Location as LocationPoint)?.Point ?? damper.GetTransform().Origin;

                    // 2. Find the connector side (world-axis logic)
                    string side = JSE_RevitAddin_MEP_OPENINGS.Services.FireDamperSleevePlacerService.GetConnectorSideWorld(damper, out var connector);
                    DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Detected connector side: {side}");
                    
                    // Always use 25mm offset toward connector side
                    double offset25 = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
                    XYZ offsetVec = OffsetVector4Way(side, damper.GetTotalTransform());
                    XYZ sleevePos = damperLoc + offsetVec * offset25;
                    DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Offset vector: {offsetVec}, Sleeve position: {sleevePos}");

                    // 4. Transform sleevePos if damper is in a link
                    if (damperTransform != null)
                        sleevePos = damperTransform.OfPoint(sleevePos);

                    // 5. Duplication suppression
                    double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    if (OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(doc, sleevePos, sleeveCheckRadius))
                    {
                        skippedExistingCount++;
                        DebugLogger.Log($"[FireDamperPlaceCommand] Skipping damper id={damper.Id} due to existing sleeve at location.");
                        continue;
                    }

                    // 6. Wall intersection using offset point
                    foreach (var wallTuple in wallTuples)
                    {
                        var wall = wallTuple.Item1;
                        var wallTransform = wallTuple.Item2;

                        BoundingBoxXYZ wallBBox = wall.get_BoundingBox(null);
                        if (wallBBox == null) continue;
                        if (wallTransform != null)
                            wallBBox = TransformBoundingBox(wallBBox, wallTransform);

                        // Use helper to check if point is inside bounding box
                        if (PointInBoundingBox(sleevePos, wallBBox))
                        {
                            // Use the unified service with debug logging enabled
                            // Enable debug logging to diagnose placement issues
                            var sleevePlacer = new JSE_RevitAddin_MEP_OPENINGS.Services.FireDamperSleevePlacerService(doc, enableDebugLogging: true);
                            bool placed = sleevePlacer.PlaceFireDamperSleeve(damper, openingSymbol, wall);
                            
                            if (placed)
                            {
                                placedCount++;
                                DebugLogger.Log($"[FireDamperPlaceCommand] Successfully placed sleeve for damper id={damper.Id} in wall id={wall.Id}");
                                
                                // Enhanced connector debugging ONLY for successfully placed sleeves
                                if (damper.Id.Value == 1876234)
                                {
                                    var cm = damper.MEPModel?.ConnectorManager;
                                    if (cm != null)
                                    {
                                        DebugLogger.Log($"[FireDamperPlaceCommand] === CONNECTOR DETAILS FOR DAMPER {damper.Id} ===");
                                        DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Total connectors: {cm.Connectors.Size}");
                                        int connIndex = 0;
                                        foreach (Connector c in cm.Connectors)
                                        {
                                            var dir = c.CoordinateSystem.BasisX;
                                            var origin = c.Origin;
                                            DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Connector[{connIndex}]: Origin={origin}, BasisX={dir}");
                                            DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Connector[{connIndex}]: X={dir.X:F3}, Y={dir.Y:F3}, Z={dir.Z:F3}");
                                            DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Connector[{connIndex}]: AbsX={Math.Abs(dir.X):F3}, AbsY={Math.Abs(dir.Y):F3}, AbsZ={Math.Abs(dir.Z):F3}");
                                            connIndex++;
                                        }
                                        if (connector != null)
                                        {
                                            var selectedDir = connector.CoordinateSystem.BasisX;
                                            DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - SELECTED Connector BasisX: {selectedDir}");
                                            DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Direction analysis: Z={selectedDir.Z:F3} (>0.9=Top, <-0.9=Bottom), X={selectedDir.X:F3} (>0.9=Right, <-0.9=Left)");
                                        }
                                        DebugLogger.Log($"[FireDamperPlaceCommand] === END CONNECTOR DETAILS ===");
                                    }
                                }
                            }
                            else
                            {
                                DebugLogger.Log($"[FireDamperPlaceCommand] Failed to place sleeve for damper id={damper.Id} in wall id={wall.Id}");
                            }
                            
                            // If you want to process only one wall per damper, keep the break; otherwise, comment out the next line to process all possible walls.
                            break;
                        }
                    }
                }
                tx.Commit();
            }

            // Summary logging
            DebugLogger.Log($"FireDamperPlaceCommand summary: Total={totalDampers}, Placed={placedCount}, SkippedExisting={skippedExistingCount}");
            return Autodesk.Revit.UI.Result.Succeeded;
        }

        /// <summary>
        /// Returns true if point is inside bounding box (inclusive).
        /// </summary>
        private static bool PointInBoundingBox(XYZ pt, BoundingBoxXYZ bbox)
        {
            return pt.X >= bbox.Min.X - 1e-6 && pt.X <= bbox.Max.X + 1e-6 &&
                   pt.Y >= bbox.Min.Y - 1e-6 && pt.Y <= bbox.Max.Y + 1e-6 &&
                   pt.Z >= bbox.Min.Z - 1e-6 && pt.Z <= bbox.Max.Z + 1e-6;
        }

        /// <summary>
        /// Returns "Left", "Right", "Top", "Bottom" for MSFD family
        /// (connectors are front/back in family but we map them to Left/Right).
        /// </summary>
        
        /// <summary>
        /// Converts side string to unit vector in WORLD coordinates (not damper local).
        /// </summary>
        private static XYZ OffsetVector4Way(string side, Transform damperT)
        {
            switch (side)
            {
                case "Right":  return  XYZ.BasisX;   // World +X direction
                case "Left":   return -XYZ.BasisX;   // World -X direction  
                case "Top":    return  XYZ.BasisZ;   // World +Z direction
                case "Bottom": return -XYZ.BasisZ;   // World -Z direction
                default:       return  XYZ.BasisX;   // fallback to world +X
            }
        }

        // Helper to transform a bounding box by a transform
        private BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            var corners = new[] {
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
            };
            var transformed = corners.Select(pt => transform.OfPoint(pt)).ToList();
            var min = new XYZ(transformed.Min(p => p.X), transformed.Min(p => p.Y), transformed.Min(p => p.Z));
            var max = new XYZ(transformed.Max(p => p.X), transformed.Max(p => p.Y), transformed.Max(p => p.Z));
            var newBox = new BoundingBoxXYZ { Min = min, Max = max };
            return newBox;
        }     
    }
}
