using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DuctFittingOpeningCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Don't reinitialize log file - it should already be initialized by the main command
            DebugLogger.Log("=== DuctFittingOpeningCommand: Execute started ===");
            DebugLogger.Log($"Build Version: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");
            DebugLogger.Log($"Timestamp: {DateTime.Now}");

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            DebugLogger.Log($"Document: {doc.Title}");

            var openingSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Name.StartsWith("DS#") && sym.Family.Name == "DuctOpeningOnWall");

            if (openingSymbol == null)
            {
                DebugLogger.Log("DuctOpeningOnWall DS# family symbol not found. Aborting command.");
                TaskDialog.Show("Error", "DuctOpeningOnWall DS# family symbol not found.");
                return Result.Failed;
            }

            DebugLogger.Log($"Family symbol found: {openingSymbol.Name} in family {openingSymbol.Family.Name}");

            // Activate the symbol if needed
            using (var txActivate = new Transaction(doc, "Activate Duct Fitting Opening Symbol"))
            {
                txActivate.Start();
                if (!openingSymbol.IsActive)
                {
                    openingSymbol.Activate();
                    DebugLogger.Log("Activated opening symbol");
                }
                txActivate.Commit();
            }

            // Find a non-template 3D view for intersection calculations
            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                DebugLogger.Log("No non-template 3D view found. Creating intersection logic without 3D view.");
                TaskDialog.Show("Warning", "No 3D view found. Intersection detection may be limited.");
            }
            else
            {
                DebugLogger.Log($"Using 3D view: {view3D.Name}");
            }

            // Wall filter and intersector for robust wall detection
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            ReferenceIntersector refIntersector = null;
            if (view3D != null)
            {
                refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Face, view3D)
                {
                    FindReferencesInRevitLinks = true
                };
                DebugLogger.Log("ReferenceIntersector initialized for robust wall detection");
            }

            int elementsPlaced = 0;
            int noFittingsFound = 0;
            int noIntersectionsFound = 0;
            int failedPlacements = 0;

            using (var transaction = new Transaction(doc, "Place Duct Fitting Openings"))
            {
                transaction.Start();

                // Get all duct fittings using proper category filters
                DebugLogger.Log("=== Collecting Duct Fittings ===");
                var ductFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();

                DebugLogger.Log($"Found {ductFittings.Count} duct fittings using OST_DuctFitting category");

                // If no fittings found, try broader search
                if (ductFittings.Count == 0)
                {
                    DebugLogger.Log("No duct fittings found with OST_DuctFitting. Trying broader search...");
                    var allInstances = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Category != null &&
                               (fi.Category.Name.Contains("Duct") ||
                                fi.Category.Name.Contains("Fitting") ||
                                fi.Category.Name.Contains("Accessory")))
                        .ToList();

                    ductFittings = allInstances;
                    DebugLogger.Log($"Broader search found {ductFittings.Count} potential duct fittings");

                    // Log categories found
                    var categories = ductFittings.Select(f => f.Category?.Name ?? "Unknown").Distinct().ToList();
                    DebugLogger.Log($"Categories found: {string.Join(", ", categories)}");
                }

                if (ductFittings.Count == 0)
                {
                    DebugLogger.Log("No duct fittings found. Aborting command.");
                    TaskDialog.Show("Info", "No duct fittings found. Command aborted.");
                    transaction.RollBack();
                    return Result.Cancelled;
                }

                foreach (var fitting in ductFittings)
                {
                    DebugLogger.Log($"=== Processing Fitting ID: {fitting.Id.IntegerValue} ===");
                    DebugLogger.Log($"Family: {fitting.Symbol.Family.Name}");
                    DebugLogger.Log($"Type: {fitting.Symbol.Name}");
                    DebugLogger.Log($"Category: {fitting.Category?.Name ?? "None"}");

                    // Get fitting location and dimensions
                    var location = fitting.Location;
                    XYZ fittingCenter = null;
                    XYZ fittingDirection = XYZ.BasisX; // Default direction

                    if (location is LocationPoint locPt)
                    {
                        fittingCenter = locPt.Point;
                        DebugLogger.Log($"Fitting center (LocationPoint): {fittingCenter}");

                        // Try to get direction from fitting transform or connectors
                        var transform = fitting.GetTransform();
                        fittingDirection = transform.BasisX.Normalize();
                        DebugLogger.Log($"Fitting direction from transform: {fittingDirection}");
                    }
                    else if (location is LocationCurve locCurve && locCurve.Curve != null)
                    {
                        fittingCenter = locCurve.Curve.Evaluate(0.5, true);
                        if (locCurve.Curve is Line line)
                        {
                            fittingDirection = line.Direction.Normalize();
                        }
                        DebugLogger.Log($"Fitting center (LocationCurve): {fittingCenter}");
                        DebugLogger.Log($"Fitting direction from curve: {fittingDirection}");
                    }
                    else
                    {
                        DebugLogger.Log("Fitting has no valid location. Skipping.");
                        continue;
                    }

                    // Get fitting dimensions
                    double width = GetFittingDimension(fitting, "Width");
                    double height = GetFittingDimension(fitting, "Height");
                    double depth = GetFittingDimension(fitting, "Depth");

                    DebugLogger.Log($"Fitting dimensions: W={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, D={UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm");

                    if (width <= 0 || height <= 0)
                    {
                        DebugLogger.Log($"Invalid fitting dimensions. Using bounding box fallback.");
                        var bbox = fitting.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            var size = bbox.Max - bbox.Min;
                            width = Math.Max(size.X, size.Y);
                            height = size.Z;
                            DebugLogger.Log($"Bounding box dimensions: W={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm");
                        }

                        if (width <= 0 || height <= 0)
                        {
                            DebugLogger.Log("Could not determine fitting dimensions. Skipping.");
                            continue;
                        }
                    }

                    // Use robust wall intersection detection
                    bool placedOpening = false;
                    if (refIntersector != null)
                    {
                        placedOpening = FindWallIntersectionAndPlaceOpening(fitting, fittingCenter, fittingDirection,
                            width, height, openingSymbol, refIntersector, doc);
                    }
                    else
                    {
                        // Fallback to simple bounding box intersection
                        placedOpening = FindWallIntersectionSimple(fitting, fittingCenter, fittingDirection,
                            width, height, openingSymbol, doc);
                    }

                    if (placedOpening)
                    {
                        elementsPlaced++;
                        DebugLogger.Log($"Successfully placed opening for fitting {fitting.Id.IntegerValue}");
                    }
                    else
                    {
                        noIntersectionsFound++;
                        DebugLogger.Log($"No wall intersection found for fitting {fitting.Id.IntegerValue}");
                    }
                }

                transaction.Commit();
            }

            DebugLogger.Log("Command execution summary:");
            DebugLogger.Log("Placed openings: " + elementsPlaced);
            DebugLogger.Log("No intersections found: " + noIntersectionsFound);
            DebugLogger.Log("Failed placements: " + failedPlacements);
            DebugLogger.Log("No fittings found: " + noFittingsFound);

            if (elementsPlaced == 0)
            {
                TaskDialog.Show("Info", "No openings placed. See logs for details.");
            }
            else
            {
                TaskDialog.Show("Success", elementsPlaced + " openings placed successfully.");
            }

            return Result.Succeeded;
        }

        private double GetFittingDimension(FamilyInstance fitting, string parameterName)
        {
            try
            {
                var param = fitting.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    double value = param.AsDouble();
                    DebugLogger.Log($"Parameter '{parameterName}': {UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters):F1}mm");
                    return value;
                }

                // Try symbol parameters
                param = fitting.Symbol.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    double value = param.AsDouble();
                    DebugLogger.Log($"Symbol parameter '{parameterName}': {UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters):F1}mm");
                    return value;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error getting parameter '{parameterName}': {ex.Message}");
            }

            DebugLogger.Log($"Parameter '{parameterName}' not found or has no value");
            return 0.0;
        }

        private bool FindWallIntersectionAndPlaceOpening(FamilyInstance fitting, XYZ center, XYZ direction,
            double width, double height, FamilySymbol openingSymbol, ReferenceIntersector refIntersector, Document doc)
        {
            DebugLogger.Log($"=== Bounding Box Intersection Detection for Fitting {fitting.Id.IntegerValue} ===");

            // Get fitting bounding box
            BoundingBoxXYZ fittingBounds = fitting.get_BoundingBox(null);
            if (fittingBounds == null)
            {
                DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: no bounding box available");
                return false;
            }

            DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: bounding box Min=({fittingBounds.Min.X:F3},{fittingBounds.Min.Y:F3},{fittingBounds.Min.Z:F3}), Max=({fittingBounds.Max.X:F3},{fittingBounds.Max.Y:F3},{fittingBounds.Max.Z:F3})");

            // Find walls that intersect with the fitting bounding box
            var intersectingWalls = GetWallsIntersectingBounds(fittingBounds, doc);

            if (!intersectingWalls.Any())
            {
                DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: no walls intersect with bounding box");
                return false;
            }

            DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: found {intersectingWalls.Count} intersecting walls");

            // For now, use the first intersecting wall (we can add prioritization logic later)
            var wall = intersectingWalls.First();
            DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: using wall ID {wall.Id.IntegerValue}");

            // Calculate intersection point as center of fitting (simpler approach)
            XYZ intersectionPoint = center;
            DebugLogger.Log($"Fitting {fitting.Id.IntegerValue}: placing opening at fitting center {intersectionPoint}");

            // Add some clearance to dimensions
            double clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            double finalWidth = width + clearance * 2;
            double finalHeight = height + clearance * 2;

            try
            {
                var placer = new DuctFittingOpeningPlacer(doc);
                placer.PlaceDuctFittingOpening(fitting, intersectionPoint, finalWidth, finalHeight, direction, openingSymbol, wall);
                DebugLogger.Log($"Successfully placed opening for fitting {fitting.Id.IntegerValue}");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Exception during opening placement: {ex.Message}");
                return false;
            }
        }

        private bool FindWallIntersectionSimple(FamilyInstance fitting, XYZ center, XYZ direction,
            double width, double height, FamilySymbol openingSymbol, Document doc)
        {
            DebugLogger.Log($"=== Simple Wall Intersection Detection for Fitting {fitting.Id.IntegerValue} ===");

            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            DebugLogger.Log($"Found {walls.Count} walls to check");
            DebugLogger.Log($"DEBUG: Total walls found in simple detection: {walls.Count}");

            var fittingBBox = fitting.get_BoundingBox(null);
            if (fittingBBox == null)
            {
                DebugLogger.Log("Could not get fitting bounding box");
                return false;
            }

            bool placedAny = false;
            foreach (var wall in walls)
            {
                var wallBBox = wall.get_BoundingBox(null);
                if (wallBBox == null) continue;

                if (BoundingBoxesIntersect(wallBBox, fittingBBox))
                {
                    DebugLogger.Log($"Bounding box intersection detected with wall {wall.Id.IntegerValue}");

                    // Add clearance
                    double clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                    double finalWidth = width + clearance * 2;
                    double finalHeight = height + clearance * 2;

                    try
                    {
                        var placer = new DuctFittingOpeningPlacer(doc);
                        placer.PlaceDuctFittingOpening(fitting, center, finalWidth, finalHeight, direction, openingSymbol, wall);
                        placedAny = true;
                        DebugLogger.Log($"Successfully placed opening using simple detection");
                        break;
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Exception during simple placement: {ex.Message}");
                    }
                }
            }

            return placedAny;
        }

        private List<Wall> GetWallsIntersectingBounds(BoundingBoxXYZ fittingBounds, Document doc)
        {
            var intersectingWalls = new List<Wall>();

            try
            {
                // Get all walls in the document
                var allWalls = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();

                DebugLogger.Log($"Checking {allWalls.Count} walls for bounding box intersection");
                DebugLogger.Log($"DEBUG: Total walls found in document: {allWalls.Count}");

                if (allWalls.Count == 0)
                {
                    DebugLogger.Log("WARNING: No walls found in document! This may be a linked model issue.");
                    // Try to get walls from linked models
                    var linkCollector = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .ToList();
                    DebugLogger.Log($"Found {linkCollector.Count} linked models");
                }

                foreach (var wall in allWalls)
                {
                    // Get wall bounding box
                    BoundingBoxXYZ wallBounds = wall.get_BoundingBox(null);
                    if (wallBounds == null) continue;

                    // Check if bounding boxes intersect
                    if (BoundingBoxesIntersect(fittingBounds, wallBounds))
                    {
                        intersectingWalls.Add(wall);
                        DebugLogger.Log($"Wall {wall.Id.IntegerValue} intersects with fitting bounds");
                    }
                }

                DebugLogger.Log($"Found {intersectingWalls.Count} walls intersecting with fitting bounds");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error finding intersecting walls: {ex.Message}");
            }

            return intersectingWalls;
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            // Check if two 3D bounding boxes intersect
            // They intersect if they overlap in all three dimensions (X, Y, Z)

            bool xOverlap = box1.Min.X <= box2.Max.X && box1.Max.X >= box2.Min.X;
            bool yOverlap = box1.Min.Y <= box2.Max.Y && box1.Max.Y >= box2.Min.Y;
            bool zOverlap = box1.Min.Z <= box2.Max.Z && box1.Max.Z >= box2.Min.Z;

            return xOverlap && yOverlap && zOverlap;
        }
    }
}
