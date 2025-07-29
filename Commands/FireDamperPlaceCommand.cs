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
            DamperLogger.InitLogFile();
//             DamperLogger.Log("FireDamperPlaceCommand: Execute started.");

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Auto-select the opening family symbol
            var openingSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("DamperOpeningOnWall"));

            if (openingSymbol == null)
            {
                TaskDialog.Show("Error", "Please load sleeve opening families.");
                return Result.Failed;
            }

            // Activate the symbol
            using (var txActivate = new Transaction(doc, "Activate Opening Symbol"))
            {
                txActivate.Start();
                if (!openingSymbol.IsActive)
                    openingSymbol.Activate();
                txActivate.Commit();
            }

            // Get all dampers in the active document
            var dampers = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("Damper"))
                .ToList();

            // Get all walls in the active document
            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            // Include dampers and walls from linked files
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(li => li.GetLinkDocument() != null)
                .ToList();

            foreach (var linkInstance in linkInstances)
            {
                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) continue;

                var linkedDampers = new FilteredElementCollector(linkedDoc)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("Damper"))
                    .ToList();

                dampers.AddRange(linkedDampers);

                var linkedWalls = new FilteredElementCollector(linkedDoc)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();

                walls.AddRange(linkedWalls);
            }

            // SUPPRESSION: Collect existing damper sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("DamperOpeningOnWall") ||
                           (fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("DMS#")))
                .ToList();

//             DamperLogger.Log($"Found {existingSleeves.Count} existing damper sleeves in the model");

            // Log existing sleeve details for debugging
            foreach (var sleeve in existingSleeves.Take(5)) // Log first 5 for debugging
            {
                var location = (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin;
//                 DamperLogger.Log($"Existing sleeve: Family={sleeve.Symbol.Family.Name}, Symbol={sleeve.Symbol.Name}, Location={location}");
            }

            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin
            );

            var debugPlacer = new FireDamperSleevePlacerDebug(doc);
            int placedCount = 0;
            int skippedExistingCount = 0; // Counter for dampers with existing sleeves
            int totalDampers = dampers.Count;

//             DamperLogger.Log($"Starting damper processing: Total dampers={totalDampers}, Existing sleeves={existingSleeves.Count}");

            using (var tx = new Transaction(doc, "Place Fire Damper Sleeves"))
            {
                tx.Start();

                foreach (var damper in dampers)
                {
                    var damperBBox = damper.get_BoundingBox(null);
                    if (damperBBox == null) continue;

                    // COMPREHENSIVE SUPPRESSION CHECK: Skip if ANY sleeve (individual OR cluster) already exists at this placement point
                    XYZ damperLocation = (damper.Location as LocationPoint)?.Point ?? damper.GetTransform().Origin;
                    // Use 100mm tolerance for robust duplicate detection (was 1mm which was too precise)
                    double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    double clusterExpansion = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // 50mm additional expansion for cluster bounds

                    // Log damper location for debugging
//                     DamperLogger.Log($"Damper ID={damper.Id.IntegerValue}: location = {damperLocation}");

                    // Check if any existing sleeve (individual OR cluster) is within tolerance of this damper
                    bool hasExistingSleeveAtPlacement = OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(
                        doc, damperLocation, sleeveCheckRadius, clusterExpansion);

                    if (hasExistingSleeveAtPlacement)
                    {
//                         DamperLogger.Log($"Damper ID={damper.Id.IntegerValue}: existing sleeve found within {UnitUtils.ConvertFromInternalUnits(sleeveCheckRadius, UnitTypeId.Millimeters):F0}mm, skipping");
                        skippedExistingCount++;
                        continue;
                    }

//                     DamperLogger.Log($"Damper ID={damper.Id.IntegerValue}: no existing sleeve within {UnitUtils.ConvertFromInternalUnits(sleeveCheckRadius, UnitTypeId.Millimeters):F0}mm, proceeding with placement check");

                    foreach (var wall in walls)
                    {
                        var wallBBox = wall.get_BoundingBox(null);
                        if (wallBBox == null) continue;

                        if (BoundingBoxesIntersect(damperBBox, wallBBox))
                        {
                            bool placed = debugPlacer.PlaceFireDamperSleeveDebug(damper, openingSymbol, wall);
                            if (placed) placedCount++;
                            break;
                        }
                    }
                }

                tx.Commit();
            }

//             DamperLogger.Log($"FireDamperPlaceCommand summary: Total={dampers.Count}, Placed={placedCount}, SkippedExisting={skippedExistingCount}");
//             DamperLogger.Log($"Fire damper sleeves placement completed: Placed {placedCount}, Skipped {skippedExistingCount} with existing sleeves.");
            return Result.Succeeded;
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            return !(box1.Max.X < box2.Min.X || box1.Min.X > box2.Max.X ||
                     box1.Max.Y < box2.Min.Y || box1.Min.Y > box2.Max.Y ||
                     box1.Max.Z < box2.Min.Z || box1.Min.Z > box2.Max.Z);
        }
    }
}
