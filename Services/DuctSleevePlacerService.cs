using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class DuctSleevePlacerService
    {
        private readonly Document _doc;
        private readonly List<(Duct, Transform?)> _ductTuples;
        private readonly List<(Element, Transform?)> _structuralElements;
        private readonly FamilySymbol _ductWallSymbol;
        private readonly FamilySymbol _ductSlabSymbol;
        private readonly Action<string> _log;

        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public DuctSleevePlacerService(
            Document doc,
            List<(Duct, Transform?)> ductTuples,
            List<(Element, Transform?)> structuralElements,
            FamilySymbol ductWallSymbol,
            FamilySymbol ductSlabSymbol,
            Action<string> log)
        {
            _doc = doc;
            _ductTuples = ductTuples;
            _structuralElements = structuralElements;
            _ductWallSymbol = ductWallSymbol;
            _ductSlabSymbol = ductSlabSymbol;
            _log = log;
        }

        public void PlaceAllDuctSleeves()
        {
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;

            // Collect all sleeves and filter by section box
            var allSleeves = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => (fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab")))
                .ToList();

            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (_doc.ActiveView is View3D vb)
                    sectionBox = SectionBoxHelper.GetSectionBoxBounds(vb);
            }
            catch { /* ignore */ }

            if (sectionBox != null)
            {
                allSleeves = allSleeves.Where(s =>
                {
                    var bb = s.get_BoundingBox(null);
                    return bb != null && BoundingBoxesIntersect(bb, sectionBox);
                }).ToList();
            }

            var sleeveGrid = new SleeveSpatialGrid(allSleeves);
            var spatialService = new SpatialPartitioningService(_structuralElements);

            foreach (var tuple in _ductTuples)
            {
                var duct = tuple.Item1;
                var transform = tuple.Item2;
                if (duct == null) { SkippedCount++; continue; }
                var locCurve = duct.Location as LocationCurve;
                var ductLine = locCurve?.Curve as Line;
                    if (ductLine == null)
                    {
                        _log?.Invoke($"SKIP: Duct {duct?.Id} is not a line");
                        SkippedCount++;
                        continue;
                    }
                
                Line hostLine = ductLine;
                if (transform != null)
                {
                    hostLine = Line.CreateBound(
                        transform.OfPoint(ductLine.GetEndPoint(0)),
                        transform.OfPoint(ductLine.GetEndPoint(1))
                    );
                }
                _log?.Invoke($"PROCESSING: Duct {duct.Id} Line Start={hostLine.GetEndPoint(0)}, End={hostLine.GetEndPoint(1)}");

                var nearbyStructuralElements = spatialService.GetNearbyElements(duct);
                if (!nearbyStructuralElements.Any())
                {
                    _log?.Invoke($"SKIP: Duct {duct.Id} no nearby structural elements found.");
                    SkippedCount++;
                    continue;
                }

                List<(Element, BoundingBoxXYZ, XYZ)> intersections;
                if (transform != null)
                {
                    // We have a linked-source duct: transform its line into host coordinates and
                    // call the host-line FindIntersections overload to avoid comparing link-space
                    // bounding boxes with host-space geometry.
                    var p0 = hostLine.GetEndPoint(0);
                    var p1 = hostLine.GetEndPoint(1);
                    var mepBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(p0.X, p1.X), Math.Min(p0.Y, p1.Y), Math.Min(p0.Z, p1.Z)),
                        Max = new XYZ(Math.Max(p0.X, p1.X), Math.Max(p0.Y, p1.Y), Math.Max(p0.Z, p1.Z))
                    };
                    intersections = MepIntersectionService.FindIntersections(hostLine, mepBBox, nearbyStructuralElements, _log ?? (_ => {}));
                }
                else
                {
                    intersections = MepIntersectionService.FindIntersections(duct, nearbyStructuralElements, _log ?? (_ => {}));
                }
                if (intersections.Count > 0)
                {
                    foreach (var intersectionTuple in intersections)
                    {
                        Element hostElem = intersectionTuple.Item1;
                        BoundingBoxXYZ bbox = intersectionTuple.Item2;
                        XYZ placePt = intersectionTuple.Item3;
                        // Early null-guard: ensure host element exists before dereferencing
                        if (hostElem == null)
                        {
                            _log?.Invoke($"SKIP: Duct {duct.Id} intersection host element is null (early skip).");
                            SkippedCount++;
                            continue;
                        }
                        // NOTE: `MepIntersectionService` returns intersection points in the active
                        // document coordinate space (linked structural solids are transformed before
                        // intersection). Therefore do NOT re-apply the source-element tuple transform
                        // to `placePt` — doing so double-transforms the point and produces large
                        // coordinate deltas. Use the intersection-derived point as-is for placement.
                        XYZ placePtToUse = placePt;
                        if (transform != null)
                        {
                            _log?.Invoke($"[TransformDebug] Info: duct source tuple contains a transform; ignoring it for intersection-derived point (placePt={placePt}).");
                        }
                        bool hostIsLinked = hostElem.Document != null && hostElem.Document != _doc;
                        double distMM = UnitUtils.ConvertFromInternalUnits(placePtToUse.DistanceTo(placePt), UnitTypeId.Millimeters);
                        double dzMM = UnitUtils.ConvertFromInternalUnits(placePtToUse.Z - placePt.Z, UnitTypeId.Millimeters);
                        string hostType = hostElem.GetType().Name;
                        string hostId = hostElem.Id.IntegerValue.ToString();
                        _log?.Invoke($"HOST: Duct {duct.Id} intersects {hostType} {hostId} BBox=({bbox.Min},{bbox.Max})");

                        _log?.Invoke($"[PlacementDebug] Duct {duct.Id} host {hostType} (ID:{hostId}) hostIsLinked={hostIsLinked}, transformProvided={(transform!=null)}");
                        _log?.Invoke($"[PlacementDebug]   - intersection center: {placePt}");
                        _log?.Invoke($"[PlacementDebug]   - placePtToUse (for active doc): {placePtToUse} (delta={distMM:F1}mm, dz={dzMM:F1}mm)");

                        if (hostElem is Floor floor)
                        {
                            var isStructuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                            bool isStructural = isStructuralParam != null && isStructuralParam.AsInteger() == 1;
                            if (!isStructural)
                            {
                                _log?.Invoke($"SKIP: Duct {duct.Id} host Floor {floor.Id.IntegerValue} is NON-STRUCTURAL. Sleeve will NOT be placed.");
                                SkippedCount++;
                                continue;
                            }
                        }
                        
                        double indivTol = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
                        double clusterTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                        // Use optimized duplication suppressor which checks individual sleeves first and
                        // then cluster bounding boxes. Pass hostType and sectionBox to reduce scanning.
                        string hostTypeFilter = hostElem is Wall ? "DuctOpeningOnWall" : (hostElem is Floor ? "DuctOpeningOnSlab" : "DuctOpeningOnWall");

                        _log?.Invoke($"[DuplicationCheck Optimized] Duct {duct.Id}: using enhanced duplication checker hostType={hostTypeFilter}");

                        var nearbySleeves = sleeveGrid.GetNearbySleeves(placePt, indivTol > clusterTol ? indivTol : clusterTol);
                        bool duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationOptimized(placePt, indivTol, clusterTol, nearbySleeves, hostTypeFilter);

                        if (duplicateExists)
                        {
                            _log?.Invoke($"SKIP: Duct {duct.Id} host {hostType} {hostId} duplicate sleeve (individual or cluster) exists near {placePt} (optimized check)");
                            SkippedCount++;
                            continue;
                        }
                        double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                        double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                        // Support round ducts: use diameter when width/height are not provided
                        double diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                        double clearance = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveClearanceHelper.GetClearance(duct);

                        if ((w <= 0.0 || h2 <= 0.0) && diameter > 0.0)
                        {
                            // Round duct detected - use diameter for both width and height
                            _log?.Invoke($"INFO: Duct {duct.Id} appears round - using diameter for sleeve: diameter={UnitUtils.ConvertFromInternalUnits(diameter, UnitTypeId.Millimeters):F1}mm");
                            w = diameter;
                            h2 = diameter;
                        }

                        // Apply per-side clearance (clearance is per-side, so add twice)
                        w = w + 2 * clearance;
                        h2 = h2 + 2 * clearance;
                        // hostElem was null-guarded earlier; no need to check again here.
                        try
                        {
                            FamilySymbol? symbolToUse = null;
                            if (hostElem is Floor)
                                symbolToUse = _ductSlabSymbol;
                            else if (hostElem is Wall)
                                symbolToUse = _ductWallSymbol;
                            else if (hostElem is FamilyInstance fi && fi.StructuralType == StructuralType.Beam)
                                symbolToUse = _ductWallSymbol;
                            else
                                symbolToUse = _ductWallSymbol; 
                            if (symbolToUse == null)
                            {
                                _log?.Invoke($"ERROR: Duct {duct.Id} host {hostType} {hostId} no suitable family symbol found.");
                                ErrorCount++;
                                continue;
                            }
                            var placer = new DuctSleevePlacer(_doc);
                            if (hostElem != null)
                            {
                                LocationCurve? ductLocation = duct.Location as LocationCurve;
                                XYZ ductWidthDirection = XYZ.BasisY; 
                                try
                                {
                                    ConnectorManager connectorManager = duct.ConnectorManager;
                                    if (connectorManager != null)
                                    {
                                        foreach (Connector connector in connectorManager.Connectors)
                                        {
                                            if (connector.ConnectorType == ConnectorType.End)
                                            {
                                                Transform connectorTransform = connector.CoordinateSystem;
                                                if (connectorTransform != null)
                                                {
                                                    ductWidthDirection = connectorTransform.BasisX;
                                                    break; 
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _log?.Invoke($"Error getting connector orientation: {ex.Message}");
                                    if (ductLocation?.Curve is Line ductLocationLine)
                                    {
                                        XYZ ductFlowDirection = ductLocationLine.Direction;
                                        if (Math.Abs(ductFlowDirection.Z) < 0.9) 
                                        {
                                            ductWidthDirection = new XYZ(-ductFlowDirection.Y, ductFlowDirection.X, 0);
                                            if (ductWidthDirection.GetLength() > 0.001)
                                                ductWidthDirection = ductWidthDirection.Normalize();
                                            else
                                                ductWidthDirection = XYZ.BasisY;
                                        }
                                    }
                                }
                                
                                double dotY = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisY));
                                double dotX = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisX));
                                string orientationStatus = dotY > dotX ? "Y-ORIENTED" : "X-ORIENTED";

                                if (orientationStatus == "Y-ORIENTED")
                                {
                                    placer.PlaceDuctSleeveWithOrientation(duct, placePtToUse, w, h2, hostLine.Direction, ductWidthDirection, symbolToUse, hostElem);
                                }
                                else
                                {
                                    placer.PlaceDuctSleeve(duct, placePtToUse, w, h2, hostLine.Direction, symbolToUse, hostElem);
                                }
                                _log?.Invoke($"PLACED: Duct {duct.Id} host {hostType} {hostId} at {placePtToUse} (original={placePt}) size=({w},{h2})");
                                PlacedCount++;
                            }
                            else
                            {
                                _log?.Invoke($"SKIP: Duct {duct.Id} intersection host element is null (not placing sleeve)");
                                SkippedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            _log?.Invoke($"ERROR: Duct {duct.Id} host {hostType} {hostId} failed to place sleeve: {ex.Message}");
                            ErrorCount++;
                        }
                    }
                }
                else
                {
                    _log?.Invoke($"SKIP: Duct {duct.Id} no intersection with any structural element");
                    SkippedCount++;
                }
            }
        }

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }
    }
}
