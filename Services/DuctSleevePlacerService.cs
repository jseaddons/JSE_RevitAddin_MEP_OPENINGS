using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using JSE_RevitAddin_MEP_OPENINGS.Models;

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
            // Set logging context for DuctSleevePlacerService debugging
            DebugLogger.SetServiceContext("SleevePlacers");
            
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
            
            DebugLogger.Info($"[DuctSleevePlacerService] Starting PlaceAllDuctSleeves with {_ductTuples.Count} ducts and {_structuralElements.Count} structural elements");

            var settings = ApplicationProfileService.Instance.GetCurrentSettings();

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
                    _log?.Invoke($"WARNING: Duct {duct.Id} no nearby structural elements found via spatial partitioning. Falling back to all structural elements.");
                    
                    // FALLBACK: Use all structural elements when spatial partitioning fails
                    // This ensures the system works regardless of coordinate/precision issues
                    nearbyStructuralElements = _structuralElements;
                    _log?.Invoke($"Fallback: Using all {nearbyStructuralElements.Count} structural elements for duct {duct.Id}");
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
                    var intersections4Tuple = MepIntersectionService.FindIntersections(duct, nearbyStructuralElements, _log ?? (_ => {}));
                    intersections = intersections4Tuple.Select(x => (x.Item2, x.Item3, x.Item4)).ToList();
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

                        // IMPLEMENTED: Check for existing cluster openings before placing individual sleeves
                        // This prevents individual sleeves from being placed where cluster openings already exist
                        if (!duplicateExists)
                        {
                            try
                            {
                                // Check for existing ClusterOpeningOnWallX cluster openings
                                var existingClusterOpenings = new FilteredElementCollector(_doc)
                                    .OfClass(typeof(FamilyInstance))
                                    .Cast<FamilyInstance>()
                                    .Where(fi => fi.Symbol?.Family?.Name != null &&
                                           fi.Symbol.Family.Name.Contains("ClusterOpeningOnWallX"))
                                    .Where(fi => 
                                    {
                                        var fiLocation = (fi.Location as LocationPoint)?.Point;
                                        if (fiLocation == null) return false;
                                        
                                        // Check if placement point is within the cluster opening's bounding box
                                        var clusterBBox = fi.get_BoundingBox(null);
                                        if (clusterBBox == null) return false;
                                        
                                        // Use 2D XY check for cluster membership (clusters are typically planar in XY)
                                        bool insideXY = placePt.X >= clusterBBox.Min.X && placePt.X <= clusterBBox.Max.X &&
                                                       placePt.Y >= clusterBBox.Min.Y && placePt.Y <= clusterBBox.Max.Y;
                                        
                                        if (insideXY)
                                        {
                                            _log?.Invoke($"[ClusterCheck] SKIP: Duct {duct.Id} placement point {placePt} is INSIDE existing cluster opening {fi.Symbol.Family.Name} (ID:{fi.Id.IntegerValue}) bounds min=({clusterBBox.Min.X:F3},{clusterBBox.Min.Y:F3}) max=({clusterBBox.Max.X:F3},{clusterBBox.Max.Y:F3})");
                                        }
                                        
                                        return insideXY;
                                    })
                                    .ToList();

                                if (existingClusterOpenings.Any())
                                {
                                    _log?.Invoke($"SKIP: Duct {duct.Id} host {hostType} {hostId} suppressed by existing cluster opening at {placePt} (cluster opening check)");
                                    SkippedCount++;
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                _log?.Invoke($"[ClusterCheck] Cluster opening check failed: {ex.Message}");
                            }
                        }

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
                        
                        // DEBUG: Log original duct dimensions
                        DebugLogger.Log($"[DUCT_SIZE_DEBUG] Duct {duct.Id} original dimensions:");
                        DebugLogger.Log($"  - Width: {w} internal units ({UnitUtils.ConvertFromInternalUnits(w, UnitTypeId.Millimeters):F1}mm)");
                        DebugLogger.Log($"  - Height: {h2} internal units ({UnitUtils.ConvertFromInternalUnits(h2, UnitTypeId.Millimeters):F1}mm)");
                        DebugLogger.Log($"  - Diameter: {diameter} internal units ({UnitUtils.ConvertFromInternalUnits(diameter, UnitTypeId.Millimeters):F1}mm)");
                        double clearance = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveClearanceHelper.GetClearance(duct);

                        if ((w <= 0.0 || h2 <= 0.0) && diameter > 0.0)
                        {
                            DebugLogger.Log($"[DUCT_SIZE_DEBUG] Duct {duct.Id} detected as round duct - using diameter");
                            if (diameter > settings.RoundOpeningsRectangular)
                            {
                                _log?.Invoke($"INFO: Duct {duct.Id} is round and its diameter is greater than {settings.RoundOpeningsRectangular}. Creating a rectangular sleeve.");
                                w = diameter;
                                h2 = diameter;
                            }
                            else
                            {
                                // Round duct detected - use diameter for both width and height
                                _log?.Invoke($"INFO: Duct {duct.Id} appears round - using diameter for sleeve: diameter={UnitUtils.ConvertFromInternalUnits(diameter, UnitTypeId.Millimeters):F1}mm");
                                w = diameter;
                                h2 = diameter;
                            }
                            DebugLogger.Log($"[DUCT_SIZE_DEBUG] After round duct processing - Width: {w}, Height: {h2}");
                        }

                        // Apply per-side clearance (clearance is per-side, so add twice)
                        w = w + 2 * clearance;
                        h2 = h2 + 2 * clearance;
                        
                        // DEBUG: Log dimensions after clearance application
                        DebugLogger.Log($"[DUCT_SIZE_DEBUG] Duct {duct.Id} after clearance application:");
                        DebugLogger.Log($"  - Clearance: {UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F1}mm ({clearance} internal units)");
                        DebugLogger.Log($"  - Final Width: {w} internal units ({UnitUtils.ConvertFromInternalUnits(w, UnitTypeId.Millimeters):F1}mm)");
                        DebugLogger.Log($"  - Final Height: {h2} internal units ({UnitUtils.ConvertFromInternalUnits(h2, UnitTypeId.Millimeters):F1}mm)");
                        DebugLogger.Log($"  - Opening Area: {w * h2} internal units²");
                        DebugLogger.Log($"  - Threshold: {settings.IgnoreOpeningsSmallerThan} internal units²");

                        if (w * h2 < settings.IgnoreOpeningsSmallerThan)
                        {
                            _log?.Invoke($"SKIP: Duct {duct.Id} opening is smaller than {settings.IgnoreOpeningsSmallerThan}. Skipping placement.");
                            SkippedCount++;
                            continue;
                        }

                        if (settings.RoundUpDimensions != "Do not round up")
                        {
                            double roundValue = 0;
                            if(double.TryParse(settings.RoundUpDimensions, out roundValue) && roundValue > 0)
                            {
                                w = Math.Ceiling(w / roundValue) * roundValue;
                                h2 = Math.Ceiling(h2 / roundValue) * roundValue;
                            }
                        }

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
                                _log?.Invoke($"SKIP: Duct {duct.Id} intersection host element is null (not placing sleeve).");
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
            
            DebugLogger.Info($"[DuctSleevePlacerService] PlaceAllDuctSleeves completed: Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}");
        }

        /// <summary>
        /// OPTIMIZED: Place duct sleeves using pre-detected clash zones (avoids re-finding intersections)
        /// </summary>
        public void PlaceAllDuctSleevesWithClashZones(List<ClashZone> clashZones)
        {
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;
            
            DebugLogger.Info($"[DuctSleevePlacerService] Starting PlaceAllDuctSleevesWithClashZones with {_ductTuples.Count} ducts and {clashZones.Count} clash zones");

            var settings = ApplicationProfileService.Instance.GetCurrentSettings();

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

            // OPTIMIZATION: Group clash zones by MEP element for efficient processing
            var clashZonesByMep = clashZones
                .GroupBy(cz => cz.MepElementId)
                .ToDictionary(g => g.Key, g => g.ToList());

            DebugLogger.Info($"[DuctSleevePlacerService] Grouped clash zones by {clashZonesByMep.Count} MEP elements");

            foreach (var tuple in _ductTuples)
            {
                var duct = tuple.Item1;
                var transform = tuple.Item2;
                if (duct == null) { SkippedCount++; continue; }

                // OPTIMIZATION: Get clash zones for this specific duct
                if (!clashZonesByMep.TryGetValue(duct.Id, out var ductClashZones))
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] No clash zones found for duct {duct.Id} - skipping");
                    SkippedCount++;
                    continue;
                }

                DebugLogger.Info($"[DuctSleevePlacerService] Processing duct {duct.Id} with {ductClashZones.Count} clash zones");

                // OPTIMIZATION: Use clash zone intersection points directly
                foreach (var clashZone in ductClashZones)
                {
                    try
                    {
                        // OPTIMIZATION: Check if clash zone is already resolved (individual sleeve placed)
                        if (clashZone.IsResolved)
                        {
                            DebugLogger.Info($"[DuctSleevePlacerService] SKIP: Clash zone {clashZone.Id} already resolved - individual sleeve already placed");
                            SkippedCount++;
                            continue;
                        }
                        
                        // OPTIMIZATION: Check if clash zone is cluster resolved
                        if (clashZone.IsClusterResolved)
                        {
                            DebugLogger.Info($"[DuctSleevePlacerService] SKIP: Clash zone {clashZone.Id} cluster resolved - cluster sleeve already placed");
                            SkippedCount++;
                            continue;
                        }

                        DebugLogger.Info($"[DuctSleevePlacerService] Processing clash zone {clashZone.Id}: MEP={clashZone.MepElementId}, Structural={clashZone.StructuralElementId}");
                        
                        var structuralElement = _doc.GetElement(clashZone.StructuralElementId);
                        if (structuralElement == null)
                        {
                            DebugLogger.Warning($"[DuctSleevePlacerService] Structural element {clashZone.StructuralElementId} not found for clash zone {clashZone.Id}");
                            ErrorCount++;
                            continue;
                        }

                        // Use the pre-calculated intersection point from clash zone
                        var intersectionPoint = clashZone.IntersectionPoint;
                        var clashBoundingBox = clashZone.ClashBoundingBox;

                        DebugLogger.Info($"[DuctSleevePlacerService] Using clash zone intersection point: ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})");

                        // OPTIMIZATION: Place sleeve directly using clash zone data (skip intersection detection)
                        PlaceSleeveFromClashZone(duct, structuralElement, intersectionPoint, clashBoundingBox, transform, sleeveGrid, clashZone);
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[DuctSleevePlacerService] Error processing clash zone {clashZone.Id}: {ex.Message}");
                        ErrorCount++;
                    }
                }
            }
            
            DebugLogger.Info($"[DuctSleevePlacerService] PlaceAllDuctSleevesWithClashZones completed: Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}");
        }

        /// <summary>
        /// OPTIMIZED: Place sleeve directly using clash zone data (avoids re-finding intersections)
        /// </summary>
        private void PlaceSleeveFromClashZone(Duct duct, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ clashBoundingBox, Transform? transform, SleeveSpatialGrid sleeveGrid, ClashZone clashZone)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                // OPTIMIZATION: Check if clash zone is already resolved (sleeve placed)
                if (clashZone.IsResolved)
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] SKIP: Clash zone {clashZone.Id} already resolved - sleeve already placed");
                    SkippedCount++;
                    return;
                }
                
                // Use intersection point directly from clash zone
                XYZ placePtToUse = intersectionPoint;
                
                // OPTIMIZATION: No expensive spatial duplicate detection needed
                // The IsResolved flag check above already prevents duplicates
                
                // Get duct dimensions and apply clearance
                double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                double h2 = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                double diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                
                double clearance = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveClearanceHelper.GetClearance(duct);
                
                // Handle round ducts
                if ((w <= 0.0 || h2 <= 0.0) && diameter > 0.0)
                {
                    if (diameter > settings.RoundOpeningsRectangular)
                    {
                        w = diameter;
                        h2 = diameter;
                    }
                    else
                    {
                        w = diameter;
                        h2 = diameter;
                    }
                }
                
                // Apply clearance
                w = w + 2 * clearance;
                h2 = h2 + 2 * clearance;
                
                // Check minimum size
                if (w * h2 < settings.IgnoreOpeningsSmallerThan)
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] SKIP: Duct {duct.Id} opening too small");
                    SkippedCount++;
                    return;
                }
                
                // Round up dimensions if needed
                if (settings.RoundUpDimensions != "Do not round up")
                {
                    if (double.TryParse(settings.RoundUpDimensions, out double roundValue) && roundValue > 0)
                    {
                        w = Math.Ceiling(w / roundValue) * roundValue;
                        h2 = Math.Ceiling(h2 / roundValue) * roundValue;
                    }
                }
                
                // Select appropriate symbol
                FamilySymbol? symbolToUse = null;
                if (structuralElement is Floor)
                    symbolToUse = _ductSlabSymbol;
                else if (structuralElement is Wall)
                    symbolToUse = _ductWallSymbol;
                else if (structuralElement is FamilyInstance fi && fi.StructuralType == StructuralType.Beam)
                    symbolToUse = _ductWallSymbol;
                else
                    symbolToUse = _ductWallSymbol;
                
                if (symbolToUse == null)
                {
                    DebugLogger.Error($"[DuctSleevePlacerService] ERROR: No suitable family symbol found for duct {duct.Id}");
                    ErrorCount++;
                    return;
                }
                
                // Get duct orientation
                var ductLocation = duct.Location as LocationCurve;
                var ductLine = ductLocation?.Curve as Line;
                if (ductLine == null)
                {
                    DebugLogger.Error($"[DuctSleevePlacerService] ERROR: Duct {duct.Id} is not a line");
                    ErrorCount++;
                    return;
                }
                
                XYZ ductWidthDirection = XYZ.BasisY;
                try
                {
                    var connectorManager = duct.ConnectorManager;
                    if (connectorManager != null)
                    {
                        foreach (Connector connector in connectorManager.Connectors)
                        {
                            if (connector.ConnectorType == ConnectorType.End)
                            {
                                var connectorTransform = connector.CoordinateSystem;
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
                    DebugLogger.Warning($"[DuctSleevePlacerService] Error getting connector orientation: {ex.Message}");
                    var ductFlowDirection = ductLine.Direction;
                    if (Math.Abs(ductFlowDirection.Z) < 0.9)
                    {
                        ductWidthDirection = new XYZ(-ductFlowDirection.Y, ductFlowDirection.X, 0);
                        if (ductWidthDirection.GetLength() > 0.001)
                            ductWidthDirection = ductWidthDirection.Normalize();
                        else
                            ductWidthDirection = XYZ.BasisY;
                    }
                }
                
                // Place the sleeve
                var placer = new DuctSleevePlacer(_doc);
                double dotY = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisY));
                double dotX = Math.Abs(ductWidthDirection.DotProduct(XYZ.BasisX));
                
                bool placementSuccessful = false;
                try
                {
                    if (dotY > dotX)
                    {
                        placer.PlaceDuctSleeveWithOrientation(duct, placePtToUse, w, h2, ductLine.Direction, ductWidthDirection, symbolToUse, structuralElement);
                    }
                    else
                    {
                        placer.PlaceDuctSleeve(duct, placePtToUse, w, h2, ductLine.Direction, symbolToUse, structuralElement);
                    }
                    
                    placementSuccessful = true;
                    DebugLogger.Info($"[DuctSleevePlacerService] PLACED: Duct {duct.Id} sleeve at {placePtToUse} size=({w},{h2})");
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[DuctSleevePlacerService] FAILED: Duct {duct.Id} sleeve placement failed: {ex.Message}");
                    placementSuccessful = false;
                }
                
                // CRITICAL FIX: Only mark as resolved if placement was actually successful
                if (placementSuccessful)
                {
                    clashZone.IsResolved = true;
                    clashZone.LastUpdated = DateTime.Now;
                    DebugLogger.Info($"[DuctSleevePlacerService] Marked clash zone {clashZone.Id} as resolved - sleeve placement successful");
                }
                else
                {
                    clashZone.IsResolved = false;
                    clashZone.LastUpdated = DateTime.Now;
                    DebugLogger.Warning($"[DuctSleevePlacerService] Clash zone {clashZone.Id} remains unresolved - sleeve placement failed");
                }
                
                // ENHANCEMENT: Apply parameter transfer settings if available
                try
                {
                    var parameterTransferService = new ParameterTransferService();
                    var currentConfig = parameterTransferService.GetCurrentParameterTransferConfiguration();
                    
                    if (currentConfig != null && currentConfig.Mappings.Count > 0)
                    {
                        // Find the placed sleeve element (would need to be captured from the placer)
                        // For now, we'll get it by finding the most recently placed opening at this location
                        var recentSleeves = GetRecentSleeveElementsAtLocation(placePtToUse);
                        
                        if (recentSleeves.Count > 0)
                        {
                            var targetSleeveIds = recentSleeves.Select(s => s.Id).ToList();
                            var transferResult = parameterTransferService.ExecuteTransferConfiguration(_doc, targetSleeveIds, currentConfig);
                            
                            if (transferResult.Success)
                            {
                                DebugLogger.Info($"[DuctSleevePlacerService] Applied parameter transfer to {targetSleeveIds.Count} sleeve(s): {transferResult.Message}");
                            }
                            else
                            {
                                DebugLogger.Warning($"[DuctSleevePlacerService] Parameter transfer failed: {transferResult.Message}");
                            }
                        }
                        else
                        {
                            DebugLogger.Info($"[DuctSleevePlacerService] No recent sleeves found at location for parameter transfer");
                        }
                    }
                }
                catch (Exception paramEx)
                {
                    DebugLogger.Warning($"[DuctSleevePlacerService] Parameter transfer error: {paramEx.Message}");
                }
                
                PlacedCount++;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacerService] Error placing sleeve for duct {duct.Id}: {ex.Message}");
                ErrorCount++;
            }
        }

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }

        #region Parameter Transfer Support

        /// <summary>
        /// Gets recent sleeve elements placed at or near the specified location
        /// </summary>
        private List<Element> GetRecentSleeveElementsAtLocation(XYZ location)
        {
            try
            {
                var recentSleeves = new List<Element>();
                var tolerance = 0.1; // 0.1 feet tolerance for finding sleeves
                
                // Create a bounding box around the location
                var min = new XYZ(location.X - tolerance, location.Y - tolerance, location.Z - tolerance);
                var max = new XYZ(location.X + tolerance, location.Y + tolerance, location.Z + tolerance);
                var boundingBox = new BoundingBoxXYZ { Min = min, Max = max };
                
                // Create a filter for opening elements (Generic Models, typically sleeves)
                var categoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
                var boundingBoxFilter = new BoundingBoxIntersectsFilter(new Outline(min, max));
                var combinedFilter = new LogicalAndFilter(categoryFilter, boundingBoxFilter);
                
                // Collect elements
                var collector = new FilteredElementCollector(_doc)
                    .WherePasses(combinedFilter)
                    .WhereElementIsNotElementType()
                    .ToElements();
                
                // Filter for elements that are likely sleeves (have appropriate parameters)
                foreach (var element in collector)
                {
                    try
                    {
                        // Check if element has sleeve-like parameters (Width, Height, etc.)
                        var hasWidthParam = element.LookupParameter("Width") != null;
                        var hasHeightParam = element.LookupParameter("Height") != null;
                        
                        if (hasWidthParam || hasHeightParam)
                        {
                            recentSleeves.Add(element);
                        }
                    }
                    catch
                    {
                        // Continue if parameter check fails
                    }
                }
                
                return recentSleeves;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[DuctSleevePlacerService] Error finding recent sleeves at location: {ex.Message}");
                return new List<Element>();
            }
        }

        #endregion

    }
}

                
                // ENHANCEMENT: Apply parameter transfer settings if available
                try
                {
                    var parameterTransferService = new ParameterTransferService();
                    var currentConfig = parameterTransferService.GetCurrentParameterTransferConfiguration();
                    
                    if (currentConfig != null && currentConfig.Mappings.Count > 0)
                    {
                        // Find the placed sleeve element (would need to be captured from the placer)
                        // For now, we'll get it by finding the most recently placed opening at this location
                        var recentSleeves = GetRecentSleeveElementsAtLocation(placePtToUse);
                        
                        if (recentSleeves.Count > 0)
                        {
                            var targetSleeveIds = recentSleeves.Select(s => s.Id).ToList();
                            var transferResult = parameterTransferService.ExecuteTransferConfiguration(_doc, targetSleeveIds, currentConfig);
                            
                            if (transferResult.Success)
                            {
                                DebugLogger.Info($"[DuctSleevePlacerService] Applied parameter transfer to {targetSleeveIds.Count} sleeve(s): {transferResult.Message}");
                            }
                            else
                            {
                                DebugLogger.Warning($"[DuctSleevePlacerService] Parameter transfer failed: {transferResult.Message}");
                            }
                        }
                        else
                        {
                            DebugLogger.Info($"[DuctSleevePlacerService] No recent sleeves found at location for parameter transfer");
                        }
                    }
                }
                catch (Exception paramEx)
                {
                    DebugLogger.Warning($"[DuctSleevePlacerService] Parameter transfer error: {paramEx.Message}");
                }
                
                PlacedCount++;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacerService] Error placing sleeve for duct {duct.Id}: {ex.Message}");
                ErrorCount++;
            }
        }

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }

        #region Parameter Transfer Support

        /// <summary>
        /// Gets recent sleeve elements placed at or near the specified location
        /// </summary>
        private List<Element> GetRecentSleeveElementsAtLocation(XYZ location)
        {
            try
            {
                var recentSleeves = new List<Element>();
                var tolerance = 0.1; // 0.1 feet tolerance for finding sleeves
                
                // Create a bounding box around the location
                var min = new XYZ(location.X - tolerance, location.Y - tolerance, location.Z - tolerance);
                var max = new XYZ(location.X + tolerance, location.Y + tolerance, location.Z + tolerance);
                var boundingBox = new BoundingBoxXYZ { Min = min, Max = max };
                
                // Create a filter for opening elements (Generic Models, typically sleeves)
                var categoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);
                var boundingBoxFilter = new BoundingBoxIntersectsFilter(new Outline(min, max));
                var combinedFilter = new LogicalAndFilter(categoryFilter, boundingBoxFilter);
                
                // Collect elements
                var collector = new FilteredElementCollector(_doc)
                    .WherePasses(combinedFilter)
                    .WhereElementIsNotElementType()
                    .ToElements();
                
                // Filter for elements that are likely sleeves (have appropriate parameters)
                foreach (var element in collector)
                {
                    try
                    {
                        // Check if element has sleeve-like parameters (Width, Height, etc.)
                        var hasWidthParam = element.LookupParameter("Width") != null;
                        var hasHeightParam = element.LookupParameter("Height") != null;
                        
                        if (hasWidthParam || hasHeightParam)
                        {
                            recentSleeves.Add(element);
                        }
                    }
                    catch
                    {
                        // Continue if parameter check fails
                    }
                }
                
                return recentSleeves;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[DuctSleevePlacerService] Error finding recent sleeves at location: {ex.Message}");
                return new List<Element>();
            }
        }

        #endregion

    }
}
