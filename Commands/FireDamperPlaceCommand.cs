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
                // Route command-level logging to the damper-specific log file so
                // all messages produced by this command appear in the damper log
                // instead of the default cable-tray log.
                DebugLogger.SetDamperLogFile();
                DebugLogger.InitLogFile("dampersleeveplacer");
                DamperLogger.InitLogFile();

                DebugLogger.Log("=== [FireDamperPlaceCommand] DEBUG START ===");
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


            // Collect visible dampers (host + visible links) using helper
            var damperTuples = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectFamilyInstancesVisibleOnly(doc, "Damper");
            DebugLogger.Log($"[FireDamperPlaceCommand] damperTuples total count: {damperTuples.Count}");

            // Apply section-box filtering to dampers so we only process dampers inside the active 3D view's section box
            // This mirrors the proven duct collection logic and prevents scanning the entire model.
            try
            {
                var uiDocLocal = new UIDocument(doc);
                var rawList = damperTuples.Select(d => (element: (Element)d.instance, transform: d.transform)).ToList();
                var filtered = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.FilterElementsBySectionBox(uiDocLocal, rawList);
                var filteredDampers = filtered.Select(t => ((FamilyInstance)t.element, t.transform)).ToList();
                DebugLogger.Log($"[FireDamperPlaceCommand] damperTuples filtered by section box: before={damperTuples.Count}, after={filteredDampers.Count}");
                damperTuples = filteredDampers;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[FireDamperPlaceCommand] Section-box filtering failed: {ex.Message}");
            }

            // Collect visible walls (host + visible links) using helper
            var wallTuples = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectWallsVisibleOnly(doc);
            DebugLogger.Log($"[FireDamperPlaceCommand] wallTuples total count: {wallTuples.Count}");

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
                    XYZ offsetVec = UtilityClass.OffsetVector4Way(side, damper.GetTotalTransform());
                    XYZ sleevePos = damperLoc + offsetVec * offset25;
                    DebugLogger.Log($"[FireDamperPlaceCommand] Damper id={damper.Id} - Offset vector: {offsetVec}, Sleeve position: {sleevePos}");

                    // 4. Transform sleevePos if damper is in a link
                    if (damperTransform != null)
                        sleevePos = damperTransform.OfPoint(sleevePos);

                    // 5. Duplication suppression
                    double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    BoundingBoxXYZ? sectionBox = null;
                    if (uiDoc.ActiveView is View3D vb) sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb);
                    // Require same family for damper duplication checks to avoid skipping due to nearby duct openings
                    if (OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(doc, sleevePos, sleeveCheckRadius, clusterExpansion: 0.0, ignoreIds: null, hostType: "OpeningOnWall", sectionBox: sectionBox, requireSameFamily: true, familyName: damper.Symbol?.Family?.Name))
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
                            wallBBox = UtilityClass.TransformBoundingBox(wallBBox, wallTransform);

                        // Use helper to check if point is inside bounding box
                        if (UtilityClass.PointInBoundingBox(sleevePos, wallBBox))
                        {
                            // Use the unified service with debug logging enabled
                            // Enable debug logging to diagnose placement issues
                            var sleevePlacer = new JSE_RevitAddin_MEP_OPENINGS.Services.FireDamperSleevePlacerService(doc, enableDebugLogging: true);
                            // Find original transforms from collected tuples
                            Transform? damperTr = tuple.Item2;
                            Transform? wallTr = wallTuple.Item2;
                            bool placed = sleevePlacer.PlaceFireDamperSleeve(damper, openingSymbol, wall, damperTr, wallTr);
                            
                            if (placed)
                            {
                                placedCount++;
                                DebugLogger.Log($"[FireDamperPlaceCommand] Successfully placed sleeve for damper id={damper.Id} in wall id={wall.Id}");
                                
                                // Enhanced connector debugging ONLY for successfully placed sleeves
                                if (damper.Id.IntegerValue == 1876234)
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
            
            // Show status prompt to user
            string summary = $"FIRE DAMPER SUMMARY: Total={totalDampers}, Placed={placedCount}, Skipped={skippedExistingCount}";
            TaskDialog.Show("Fire Damper Placement", summary);
            
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}
