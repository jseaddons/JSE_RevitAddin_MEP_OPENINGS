using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class PipeSleevePlacerService
    {
        private readonly Document _doc;
        private readonly List<(Pipe, Transform?)> _pipeTuples;
        private readonly List<(Element, Transform?)> _structuralElements;
        private readonly FamilySymbol _pipeWallSymbol;
        private readonly FamilySymbol _pipeSlabSymbol;
        private readonly List<FamilyInstance> _existingSleeves;
        private readonly Action<string> _log;
        private readonly PipeSleevePlacer _placer;

        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public PipeSleevePlacerService(
            Document doc,
            List<(Pipe, Transform?)> pipeTuples,
            List<(Element, Transform?)> structuralElements,
            FamilySymbol pipeWallSymbol,
            FamilySymbol pipeSlabSymbol,
            List<FamilyInstance> existingSleeves,
            Action<string> log)
        {
            _doc = doc;
            _pipeTuples = pipeTuples;
            _structuralElements = structuralElements;
            _pipeWallSymbol = pipeWallSymbol;
            _pipeSlabSymbol = pipeSlabSymbol;
            _existingSleeves = existingSleeves;
            _log = log;
            // Fix ambiguous constructor: explicitly cast to the correct type if needed
            _placer = (PipeSleevePlacer)Activator.CreateInstance(typeof(PipeSleevePlacer), _doc)!;
        }

        public void PlaceAllPipeSleeves()
        {
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;
            _log($"PipeSleevePlacerService: Starting. Pipe count = {_pipeTuples.Count}, Structural host count = {_structuralElements.Count}");
            _log("Pipe IDs: " + string.Join(", ", _pipeTuples.Select(t => t.Item1?.Id.Value.ToString() ?? "null")));
            _log("Host types: " + string.Join(", ", _structuralElements.Select(e => e.Item1?.GetType().FullName ?? "null")));
            foreach (var tuple in _pipeTuples)
            {
                var pipe = tuple.Item1;
                var transform = tuple.Item2;
                int pipeIdValue = (int)(pipe?.Id?.Value ?? -1);
                _log($"Processing pipe {pipeIdValue}");
                if (pipe == null) { SkippedCount++; _log("Skipped: pipe is null"); continue; }
                var locCurve = pipe.Location as LocationCurve;
                var pipeLine = locCurve?.Curve as Line;
                if (pipeLine == null)
                {
                    SkippedCount++;
                    _log($"Skipped: pipe {pipeIdValue} has no valid Line geometry");
                    continue;
                }
                // --- CLUSTERING DATA PREP FOR SOIL/WASTE SYSTEMS ---
                var system = pipe.MEPSystem;
                string systemName = system?.Name ?? "NULL";
                _log($"Pipe {pipeIdValue}: MEP System = {systemName}");
                string sysName = system?.Name?.ToLowerInvariant() ?? "";
                bool isSoilWasteSanitary = !string.IsNullOrEmpty(sysName) && (sysName.Contains("soil") || sysName.Contains("waste") || sysName.Contains("sp ") || sysName.Contains("sanitary"));
                List<FamilyInstance>? allFittings = null;
                double clusterRadius = 0;
                if (isSoilWasteSanitary)
                {
                    _log($"Pipe {pipeIdValue}: SOIL/WASTE system detected - clustering data will be checked per intersection host");
                    clusterRadius = UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters);
                    var pipeCurve = pipeLine;
                    if (pipeCurve != null && system != null)
                    {
                        var hostFittings = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(fi =>
                                fi.MEPModel != null &&
                                fi.MEPModel.ConnectorManager != null &&
                                fi.MEPModel.ConnectorManager.Connectors != null &&
                                fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().Any(c => c.MEPSystem != null && c.MEPSystem.Id == system.Id)
                            )
                            .ToList();
                        _log($"Pipe {pipeIdValue}: Found {hostFittings.Count} host fittings in system {systemName}");
                        allFittings = new List<FamilyInstance>(hostFittings);
                        var linkInstances = new FilteredElementCollector(_doc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .Where(li => li.GetLinkDocument() != null)
                            .ToList();
                        foreach (var linkInstance in linkInstances)
                        {
                            var linkedDoc = linkInstance.GetLinkDocument();
                            if (linkedDoc == null) continue;
                            var linkedFittings = new FilteredElementCollector(linkedDoc)
                                .OfClass(typeof(FamilyInstance))
                                .WhereElementIsNotElementType()
                                .Cast<FamilyInstance>()
                                .Where(fi =>
                                    fi.MEPModel != null &&
                                    fi.MEPModel.ConnectorManager != null &&
                                    fi.MEPModel.ConnectorManager.Connectors != null &&
                                    fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().Any(c => c.MEPSystem != null && c.MEPSystem.Name == systemName)
                                )
                                .ToList();
                            allFittings.AddRange(linkedFittings);
                            _log($"Pipe {pipeIdValue}: Found {linkedFittings.Count} linked fittings in {linkInstance.Name}");
                        }
                        _log($"Pipe {pipeIdValue}: Total fittings in system = {allFittings.Count}");
                    }
                }
                // --- END CLUSTERING DATA PREP ---

                // Transform geometry if from a linked model
                Line hostLine = pipeLine;
                if (transform != null)
                {
                    hostLine = Line.CreateBound(
                        transform.OfPoint(pipeLine.GetEndPoint(0)),
                        transform.OfPoint(pipeLine.GetEndPoint(1))
                    );
                }
                // Accept intersections as List<(Element, object?, XYZ)> for robust transform handling
                var intersections = PipeSleeveIntersectionService.FindDirectStructuralIntersectionBoundingBoxesVisibleOnly(pipe, _structuralElements, hostLine);
                _log($"Pipe {pipe.Id.Value}: Found {intersections?.Count ?? 0} intersections");
                if (intersections != null)
                {
                    foreach (var it in intersections)
                    {
                        var h = it.Item1;
                        _log($"  Intersection host: {h?.Id.Value ?? -1}, type: {h?.GetType().FullName ?? "null"}");
                    }
                }
                if (intersections != null && intersections.Count > 0)
                {
                    foreach (var intersectionTuple in intersections)
                    {
                        Element hostElem = intersectionTuple.Item1;
                        object hostTransform = intersectionTuple.Item2;
                        XYZ intersectionPoint = intersectionTuple.Item3;
                        if (hostElem == null)
                        {
                            SkippedCount++;
                            _log($"Skipped: intersection host element is null");
                            continue;
                        }
                        _log($"Checking intersection at {intersectionPoint} with host {hostElem.Id.Value}");
                        string hostId = hostElem.Id.Value.ToString();
                        string hostType = hostElem.GetType().FullName;
                        string hostMsg = $"HOST: Pipe {pipe.Id} intersects {hostType} {hostId}";
                        _log($"[PipeSleevePlacerService] {hostMsg}");
                        try
                        {
                            bool isWall = hostElem.GetType().Name.Contains("Wall", StringComparison.OrdinalIgnoreCase);
                            bool isFraming = false;
                            bool isFloor = hostElem is Floor;
                            if (hostElem is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                            {
                                isFraming = true;
                            }
                            // --- CLUSTERING SUPPRESSION FOR WALLS AND FLOORS (PER INTERSECTION) ---
                            int intersectionClusterCount = 0;
                            if (isSoilWasteSanitary && allFittings != null && (isWall || isFloor))
                            {
                                foreach (var fi in allFittings)
                                {
                                    var fittingLoc = (fi.Location as LocationPoint)?.Point;
                                    if (fittingLoc != null)
                                    {
                                        double distance = fittingLoc.DistanceTo(intersectionPoint);
                                        double distanceMM = UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters);
                                        _log($"Pipe {pipe.Id.Value}: Fitting {fi.Id.Value} at {fittingLoc}, distance to intersection = {distanceMM:F1}mm");
                                        if (distance <= clusterRadius)
                                        {
                                            intersectionClusterCount++;
                                            _log($"Pipe {pipe.Id.Value}: Fitting {fi.Id.Value} is WITHIN cluster radius of intersection ({distanceMM:F1}mm <= 300mm)");
                                        }
                                    }
                                }
                                _log($"Pipe {pipe.Id.Value}: Cluster count at intersection = {intersectionClusterCount} (threshold = 0)");
                                if (intersectionClusterCount > 0)
                                {
                                    _log($"*** CLUSTERING SUPPRESSION *** Pipe ID={pipe.Id.Value} - sleeve placement SKIPPED for host type {(isWall ? "WALL" : "FLOOR")} due to fitting clustering for SOIL/WASTE system. Found {intersectionClusterCount} fittings within {UnitUtils.ConvertFromInternalUnits(clusterRadius, UnitTypeId.Millimeters):F0}mm radius at intersection.");
                                    SkippedCount++;
                                    continue;
                                }
                            }
                            XYZ placePoint = intersectionPoint;
                            var xform = hostTransform as Transform;
                            if (xform != null)
                            {
                                placePoint = xform.OfPoint(intersectionPoint);
                                _log($"[Transform] Applied link transform to intersectionPoint: {intersectionPoint} -> {placePoint}");
                            }
                            if (isWall)
                            {
                                var wall = hostElem as Wall;
                                if (wall != null)
                                {
                                    var wallLocCurve = wall.Location as LocationCurve;
                                    if (wallLocCurve != null && wallLocCurve.Curve != null)
                                    {
                                        XYZ wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                                        XYZ wallNormal = wall.Orientation.Normalize();
                                        double distToCenter = (placePoint - wallCenter).DotProduct(wallNormal);
                                        placePoint = placePoint - wallNormal.Multiply(distToCenter);
                                        _log($"[CenterlineProjection] Projected to wall centerline: wallCenter={wallCenter}, wallNormal={wallNormal}, result={placePoint}");
                                    }
                                }
                            }
                            _log($"Placing sleeve: hostElem type = {hostElem.GetType().FullName}, isWall = {isWall}, isFraming = {isFraming}, placePoint = {placePoint}");
                            double indivTol = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters); // 10mm for individual sleeves
                            double clusterTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm for clusters
                            bool duplicate = OpeningDuplicationChecker.IsAnySleeveAtLocation(_doc, placePoint, indivTol)
                                || OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(_doc, placePoint, clusterTol).Any();
                            if (duplicate)
                            {
                                SkippedCount++;
                                _log($"Suppressed: existing individual or cluster sleeve found at {placePoint}");
                                continue;
                            }
                            FamilySymbol symbolToUse = (isWall || isFraming) ? _pipeWallSymbol : _pipeSlabSymbol;
                            _placer.PlaceSleeve(
                                pipe,
                                placePoint,
                                hostLine.Direction,
                                symbolToUse,
                                hostElem
                            );
                            PlacedCount++;
                        }
                        catch (Exception ex)
                        {
                            ErrorCount++;
                            _log($"ERROR: Failed to place sleeve for pipe {pipe.Id} at {intersectionPoint}: {ex.Message}");
                        }
                    }
                }
                else
                {
                    SkippedCount++;
                    _log($"Skipped: pipe {pipe.Id} has no intersection with structural host");
                }
            }
        }
    }
}
