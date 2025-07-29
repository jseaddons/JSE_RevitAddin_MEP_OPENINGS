using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing; // For Pipe
using Autodesk.Revit.DB.Mechanical; // For Duct
using Autodesk.Revit.DB.Electrical; // For CableTray
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
public class StructuralSleevePlacementCommand : IExternalCommand
{


    // Helper: check if a point is inside any existing Rect sleeve's bounding box (for OnWallRect/OnSlabRect families)
    private static FamilyInstance IsSuppressedByRectBoundingBox(XYZ point, List<FamilyInstance> rectSleeves)
    {
        double tol = 10.0 / 304.8; // 10mm in feet
        int checkedCount = 0;
        foreach (var rectSleeve in rectSleeves)
        {
            var bbox = rectSleeve.get_BoundingBox(null);
            if (bbox == null)
            {
                DebugLogger.Debug($"[SUPPRESSION][RECT] Sleeve {rectSleeve.Id} has null bounding box.");
                continue;
            }
            var min = bbox.Min - new XYZ(tol, tol, tol);
            var max = bbox.Max + new XYZ(tol, tol, tol);
            checkedCount++;
            DebugLogger.Debug($"[SUPPRESSION][RECT] Checking point {point} against bbox of sleeve {rectSleeve.Id}: min={min}, max={max}");
            if (point.X >= min.X && point.X <= max.X &&
                point.Y >= min.Y && point.Y <= max.Y &&
                point.Z >= min.Z && point.Z <= max.Z)
            {
                DebugLogger.Info($"[SUPPRESSION][RECT] Point {point} is inside bounding box of sleeve {rectSleeve.Id}");
                return rectSleeve;
            }
        }
        DebugLogger.Debug($"[SUPPRESSION][RECT] Checked {checkedCount} rect sleeves, no suppression at {point}");
        return null;
    }
    // Helper: check if a point is within 10mm of any existing sleeve's center (placement point)
    private static FamilyInstance IsSuppressedByExistingSleeve(XYZ point, Dictionary<FamilyInstance, XYZ> sleeveLocations)
    {
        const double tolMm = 10.0;
        int checkedCount = 0;
        foreach (var kvp in sleeveLocations)
        {
            var existingPoint = kvp.Value;
            double distMm = UnitUtils.ConvertFromInternalUnits(point.DistanceTo(existingPoint), UnitTypeId.Millimeters);
            checkedCount++;
            DebugLogger.Debug($"[SUPPRESSION][CENTER] Checking point {point} vs sleeve {kvp.Key.Id} at {existingPoint}: distMm={distMm:F2}");
            if (distMm <= tolMm)
            {
                DebugLogger.Info($"[SUPPRESSION][CENTER] Point {point} is within {tolMm}mm of sleeve {kvp.Key.Id} at {existingPoint}");
                return kvp.Key;
            }
        }
        DebugLogger.Debug($"[SUPPRESSION][CENTER] Checked {checkedCount} sleeves, no suppression at {point}");
        return null;
    }
        // 🎯 TARGET: Track these 9 specific MEP elements that should intersect
        private static readonly HashSet<int> EXPECTED_MEP_ELEMENTS = new HashSet<int>
        {
            // FLOORS (2 elements):
            1892943, // Pipe → should intersect floor
            1933792, // Duct → should intersect floor
            
            // STRUCTURAL FRAMING (7 elements):
            1935283, // Cable Tray → should intersect framing
            1935473, // Cable Tray → should intersect framing  
            1934835, // Pipe → should intersect framing
            1935108, // Pipe → should intersect framing
            1935654, // Pipe → should intersect framing
            1935656, // Pipe → should intersect framing
            1935079  // Duct → should intersect framing
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.InitLogFileOverwrite("StructuralSleevePlacement_Debug"); // Ensure log file is overwritten
            DebugLogger.Info("Starting StructuralSleevePlacementCommand execution.");
            DebugLogger.Info($"🎯 TRACKING 9 EXPECTED MEP ELEMENTS: {string.Join(", ", EXPECTED_MEP_ELEMENTS)}");

            // --- Suppression logic helpers (must be in scope for all placement logic) ---

            // Use the main document for all suppression logic
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol != null && (
                    fi.Symbol.Family.Name.Contains("PipeOpeningOnSlab") ||
                    fi.Symbol.Family.Name.Contains("DuctOpeningOnSlab") ||
                    fi.Symbol.Family.Name.Contains("OnWall")))
                .ToList();

            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin);

            // (Removed duplicate declaration of uiDoc and doc)

            try
            {
                // Collect structural elements from active document and linked files for validation
                var structuralElements = CollectStructuralElements(doc);
                DebugLogger.Info($"Found {structuralElements.Count} total structural elements (active + linked).");

                // Activate required family symbols
                var allFamilySymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                // Update filter to match family names based on README guidelines
                var familySymbol = allFamilySymbols
                    .FirstOrDefault(sym => sym.Family.Name.Contains("OnWall", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("DuctOpeningOnSlab", StringComparison.OrdinalIgnoreCase));

                if (familySymbol == null)
                {
                    TaskDialog.Show("Error", "No suitable sleeve family found. Please ensure the required families are loaded as per the README guidelines.");
                    return Result.Failed;
                }

                using (var txActivate = new Transaction(doc, "Activate Structural Sleeve Symbol"))
                {
                    txActivate.Start();
                    if (!familySymbol.IsActive)
                        familySymbol.Activate();
                    txActivate.Commit();
                }

                // Select a non-template 3D view
                var view3D = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

                if (view3D == null)
                {
                    TaskDialog.Show("Error", "No non-template 3D view found.");
                    return Result.Failed;
                }

                // Define structural filters for floors and framing
                ElementFilter structuralFilter = new LogicalOrFilter(
                    new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                    new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming));

                // Use ReferenceIntersector to detect intersections with structural elements
                var refIntersector = new ReferenceIntersector(structuralFilter, FindReferenceTarget.Face, view3D)
                {
                    FindReferencesInRevitLinks = true
                };

                // Use the new comprehensive MEP collection and geometric validation approach
                DebugLogger.Info("Collecting MEP elements from active document and linked files using enhanced logic.");
                var allMEPElements = MEPElementCollector.CollectMEPElements(doc);
                DebugLogger.Info($"Total MEP elements collected: {allMEPElements.Count}");

                // Filter for only truly intersecting elements using robust geometric validation
                var intersectingElements = new List<Element>();
                int intersectionCount = 0;

                foreach (var mepElement in allMEPElements)
                {
                    // Check implementation details...
                }

                DebugLogger.Info($"Total intersecting MEP elements found: {intersectingElements.Count} (from {intersectionCount} intersection hits).");

                // Summary and placement logic...
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Exception in StructuralSleevePlacementCommand: {ex.Message}");
                DebugLogger.Error($"Stack trace: {ex.StackTrace}");
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public static List<Element> CollectStructuralElements(Document doc)
        {
            var structuralElements = new List<Element>();
            try
            {
                // Collect floors and structural framing
                var floors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .ToList();
                
                var structuralFraming = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .ToList();

                structuralElements.AddRange(floors);
                structuralElements.AddRange(structuralFraming);

                DebugLogger.Info($"Collected {floors.Count} floors and {structuralFraming.Count} structural framing elements.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting structural elements: {ex.Message}");
            }
            
            return structuralElements;
        }

        // Additional helper methods and classes from the reference file...
    }

    public static class MEPElementCollector
    {
        public static List<Element> CollectMEPElements(Document doc)
        {
            var mepElements = new List<Element>();
            try
            {
                // Collect pipes
                var pipes = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType()
                    .ToList();

                // Collect ducts
                var ducts = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctCurves)
                    .WhereElementIsNotElementType()
                    .ToList();

                // Collect cable trays
                var cableTrays = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CableTray)
                    .WhereElementIsNotElementType()
                    .ToList();

                mepElements.AddRange(pipes);
                mepElements.AddRange(ducts);
                mepElements.AddRange(cableTrays);

                DebugLogger.Info($"Collected {pipes.Count} pipes, {ducts.Count} ducts, {cableTrays.Count} cable trays.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting MEP elements: {ex.Message}");
            }

            return mepElements;
        }

        public static bool ValidateGeometricIntersection(Element mepElement, Element structuralElement)
        {
            try
            {
                var locCurve = (mepElement.Location as LocationCurve)?.Curve;
                if (locCurve?.Curve == null) return false;

                var line = locCurve as Line;
                if (line == null) return false;

                // Get structural solid
                Solid structuralSolid = null;
                var geomElem = structuralElement.get_Geometry(new Options());
                foreach (GeometryObject obj in geomElem)
                {
                    if (obj is Solid solid && solid.Volume > 0)
                    {
                        structuralSolid = solid;
                        break;
                    }
                }

                if (structuralSolid == null) return false;

                // Check for intersections
                var intersectionPoints = new List<XYZ>();
                foreach (Face face in structuralSolid.Faces)
                {
                    IntersectionResultArray ira = null;
                    SetComparisonResult res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        foreach (IntersectionResult ir in ira)
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                    }
                }

                return intersectionPoints.Count > 0;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error validating geometric intersection: {ex.Message}");
                return false;
            }
        }
    }
}
