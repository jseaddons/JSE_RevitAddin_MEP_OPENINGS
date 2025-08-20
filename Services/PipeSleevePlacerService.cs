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
            _placer = new PipeSleevePlacer(doc);
        }

        public void PlaceAllPipeSleeves()
        {
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;
            _log($"PipeSleevePlacerService: Starting. Pipe count = {_pipeTuples.Count}, Structural host count = {_structuralElements.Count}");
            _log("Pipe IDs: " + string.Join(", ", _pipeTuples.Select(t => t.Item1?.Id.IntegerValue.ToString() ?? "null")));
            _log("Host types: " + string.Join(", ", _structuralElements.Select(e => e.Item1?.GetType().FullName ?? "null")));

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

            foreach (var tuple in _pipeTuples)
            {
                var pipe = tuple.Item1;
                var transform = tuple.Item2;
                int pipeIdValue = (int)(pipe?.Id?.IntegerValue ?? -1);
                _log($"Processing pipe {pipeIdValue}");
                if (pipe == null) { SkippedCount++; _log("Skipped: pipe is null"); continue; }

                var locationCurve = pipe.Location as LocationCurve;
                if (locationCurve == null)
                {
                    SkippedCount++;
                    _log($"Skipped: pipe {pipeIdValue} has no LocationCurve");
                    continue;
                }

                var pipeLine = locationCurve.Curve as Line;
                if (pipeLine == null)
                {
                    SkippedCount++;
                    _log($"Skipped: pipe {pipeIdValue} has no valid Line geometry");
                    continue;
                }

                var nearbyStructuralElements = spatialService.GetNearbyElements(pipe);
                if (!nearbyStructuralElements.Any())
                {
                    _log($"SKIP: Pipe {pipe.Id} no nearby structural elements found.");
                    SkippedCount++;
                    continue;
                }

                Line hostLine = pipeLine;
                BoundingBoxXYZ? pipeBBox = null;
                if (transform != null)
                {
                    // transform pipe end points into host coords
                    hostLine = Line.CreateBound(
                        transform.OfPoint(pipeLine.GetEndPoint(0)),
                        transform.OfPoint(pipeLine.GetEndPoint(1)));
                    // derive bbox in host coords
                    var origBbox = pipe.get_BoundingBox(null);
                    if (origBbox != null)
                    {
                        var min = transform.OfPoint(origBbox.Min);
                        var max = transform.OfPoint(origBbox.Max);
                        pipeBBox = new BoundingBoxXYZ { Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)), Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z)) };
                    }
                }

                List<(Element, BoundingBoxXYZ, XYZ)> intersections = new List<(Element, BoundingBoxXYZ, XYZ)>();
                if (transform == null)
                {
                    intersections = MepIntersectionService.FindIntersections(pipe, nearbyStructuralElements, _log);
                }
                else
                {
                    _log($"[TransformDebug] Pipe {pipe.Id.IntegerValue} transformed to host line Start={hostLine.GetEndPoint(0)}, End={hostLine.GetEndPoint(1)}; pipeBBox={pipeBBox}");
                    intersections = MepIntersectionService.FindIntersections(hostLine, pipeBBox, nearbyStructuralElements, _log);
                }
                
                _log($"Pipe {pipe.Id.IntegerValue}: Found {intersections?.Count ?? 0} intersections");
                if (intersections != null)
                {
                    foreach (var it in intersections)
                    {
                        var h = it.Item1;
                        _log($"  Intersection host: {h?.Id.IntegerValue ?? -1}, type: {h?.GetType().FullName ?? "null"}");
                    }
                }
                if (intersections != null && intersections.Count > 0)
                {
                    foreach (var intersectionTuple in intersections)
                    {
                        Element hostElem = intersectionTuple.Item1;
                        XYZ intersectionPoint = intersectionTuple.Item3;
                        if (hostElem == null)
                        {
                            SkippedCount++;
                            _log("Skipped: intersection host element is null");
                            continue;
                        }
                        // Determine if this pipe is essentially vertical (Z-aligned)
                        bool isVerticalPipe = false;
                        try
                        {
                            var dir = pipeLine.Direction;
                            if (dir != null)
                            {
                                var nd = dir.Normalize();
                                isVerticalPipe = Math.Abs(nd.Z) > 0.9; // mostly vertical
                            }
                        }
                        catch { /* ignore */ }
                        _log($"Checking intersection at {intersectionPoint} with host {hostElem.Id.IntegerValue}");
                        string hostId = hostElem.Id.IntegerValue.ToString();
                        string hostType = hostElem.GetType().FullName;
                        string hostMsg = $"HOST: Pipe {pipe.Id} intersects {hostType} {hostId}";
                        _log($"[PipeSleevePlacerService] {hostMsg}");
                        try
                        {
                            bool isWall = hostElem.GetType().Name.Contains("Wall", StringComparison.OrdinalIgnoreCase);
                            bool isFraming = false;
                            if (hostElem is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                            {
                                isFraming = true;
                            }
                            // If this is a vertical pipe embedded in a wall or structural framing, skip processing
                            if (isVerticalPipe && (isWall || isFraming))
                            {
                                SkippedCount++;
                                _log($"SKIP: Pipe {pipe.Id} is vertical and hosted by {(isWall ? "Wall" : "Structural Framing")}; skipping sleeve placement.");
                                continue;
                            }
                            
                            XYZ placePoint = intersectionPoint;

                            // If host is a wall, project the intersection point to the wall location curve (nearest point)
                            if (isWall && hostElem is Wall wall)
                            {
                                try
                                {
                                    var loc = wall.Location as LocationCurve;
                                    var curve = loc?.Curve;
                                    if (curve != null)
                                    {
                                        // If the wall element lives in a linked document, do the projection in the
                                        // link's coordinate space and then transform the projected point back to host.
                                        bool hostElemIsLinked = hostElem.Document != null && hostElem.Document != _doc;
                                        Transform? hostTransform = null;
                                        if (hostElemIsLinked)
                                        {
                                            hostTransform = FindTransformForLinkedElement(hostElem);
                                        }

                                        if (hostElemIsLinked && hostTransform != null)
                                        {
                                            // intersectionPoint is in host coords (intersection calculated after transforming
                                            // link geometry). Convert it to link coords, project, then convert projected point
                                            // back to host coords for placement. Preserve original intersection Z.
                                            try
                                            {
                                                var linkIntersection = hostTransform.Inverse.OfPoint(intersectionPoint);
                                                var proj = curve.Project(linkIntersection);
                                                if (proj != null)
                                                {
                                                    var cp = proj.XYZPoint;
                                                    var cpHost = hostTransform.OfPoint(cp);
                                                    placePoint = new XYZ(cpHost.X, cpHost.Y, intersectionPoint.Z);
                                                    _log($"[CenterlineProjection] (link-doc) Projected in link coords: wallCenter=({cp.X:F6},{cp.Y:F6},{cp.Z:F6}), cp->host={cpHost}, result={placePoint}");
                                                }
                                                else
                                                {
                                                    _log($"[CenterlineProjection] (link-doc) Project returned null; falling back to intersectionPoint");
                                                    placePoint = intersectionPoint;
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                _log($"[CenterlineProjection] (link-doc) Exception projecting to wall centerline: {ex.Message}");
                                                placePoint = intersectionPoint;
                                            }
                                        }
                                        else
                                        {
                                            // Not a linked host or no transform available: perform projection in the element's
                                            // own coordinate space (this is safe when the element belongs to the active doc),
                                            // otherwise fall back to helper or intersection.
                                            if (!hostElemIsLinked)
                                            {
                                                var proj = curve.Project(intersectionPoint);
                                                if (proj != null)
                                                {
                                                    var cp = proj.XYZPoint;
                                                    placePoint = new XYZ(cp.X, cp.Y, intersectionPoint.Z);
                                                    _log($"[CenterlineProjection] Projected to wall centerline using XYZPoint: wallCenter=({cp.X:F6},{cp.Y:F6},{cp.Z:F6}), result={placePoint}");
                                                }
                                                else
                                                {
                                                    placePoint = WallCenterlineHelper.GetElementCenterlinePoint(wall, intersectionPoint);
                                                    _log($"[CenterlineProjection] Project returned null; using WallCenterlineHelper result={placePoint}");
                                                }
                                            }
                                            else
                                            {
                                                // Linked host but we couldn't locate the transform. Safer to place at the intersection
                                                // than to project with mismatched coordinates.
                                                _log($"[CenterlineProjection] Host is linked but no transform found; using intersectionPoint as placePoint");
                                                placePoint = intersectionPoint;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        placePoint = WallCenterlineHelper.GetElementCenterlinePoint(wall, intersectionPoint);
                                        _log($"[CenterlineProjection] No wall location curve; using WallCenterlineHelper result={placePoint}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _log($"[CenterlineProjection] Exception projecting to wall centerline: {ex.Message}");
                                    placePoint = intersectionPoint;
                                }
                            }

                            _log($"Placing sleeve: hostElem type = {hostElem.GetType().FullName}, isWall = {isWall}, isFraming = {isFraming}, placePoint = {placePoint}");

                            // DIAGNOSTIC: Check if host element is from linked document and needs transform
                            bool hostIsLinked = hostElem?.Document != null && hostElem.Document != _doc;
                            string hostIdStr = hostElem?.Id.IntegerValue.ToString() ?? "null";
                            string hostTypeStr = hostElem?.GetType().Name ?? "null";
                            // Read IFC GUID parameter (diagnostic) - try a few common names
                            string hostIfcGuid = "NULL";
                            try
                            {
                                if (hostElem != null)
                                {
                                    var ifcParam = hostElem.LookupParameter("IfcGUID") ?? hostElem.LookupParameter("IfcGuid") ?? hostElem.LookupParameter("Ifc GUID");
                                    if (ifcParam != null)
                                    {
                                        hostIfcGuid = ifcParam.AsString() ?? ifcParam.AsValueString() ?? "NULL";
                                    }
                                }
                            }
                            catch { /* ignore */ }
                            
                            _log($"[TransformDebug] Pipe {pipe.Id} analysis:");
                            _log($"[TransformDebug]   - Pipe document: '{pipe.Document?.Title ?? "NULL"}' vs Active: '{_doc?.Title ?? "NULL"}'");
                            _log($"[TransformDebug]   - Pipe from linked doc: {(pipe.Document != _doc)}");
                            _log($"[TransformDebug]   - Host {hostTypeStr} (ID:{hostIdStr}) document: '{hostElem?.Document?.Title ?? "NULL"}'");
                            _log($"[TransformDebug]   - Host {hostTypeStr} (ID:{hostIdStr}) from linked doc: {hostIsLinked}");
                            _log($"[TransformDebug]   - Host IfcGUID: '{hostIfcGuid}'");
                            _log($"[TransformDebug]   - Transform provided: {(transform != null)}");
                            _log($"[TransformDebug]   - Original intersection center: {intersectionPoint}");
                            _log($"[TransformDebug]   - Original placePoint: {placePoint}");
                            
                            if (transform != null)
                            {
                                _log($"[TransformDebug]   - Transform origin: {transform.Origin}");
                                _log($"[TransformDebug]   - Transform basis X: {transform.BasisX}");
                                _log($"[TransformDebug]   - Transform basis Y: {transform.BasisY}");
                                _log($"[TransformDebug]   - Transform basis Z: {transform.BasisZ}");
                            }

                            // IMPORTANT: `MepIntersectionService` returns intersection points in the active
                            // document coordinate space (it transforms linked structural solids before
                            // computing intersections). DO NOT re-apply the source-element transform to
                            // these intersection-derived points — that causes double-transforms and large
                            // coordinate deltas. We therefore use the intersection-derived placePoint as
                            // the final placement point unless we have a strong reason to prefer the
                            // projected centerline result.
                            XYZ placePtToUse = placePoint;
                            if (transform != null)
                            {
                                // There is a tuple transform (source element from a link). This transform
                                // maps link->active coordinates and should only be used for transforming
                                // source-element-local coordinates. Intersection points are already
                                // active-doc coordinates, so ignore `transform` here (log for diagnostics).
                                _log($"[TransformDebug] Info: source element has a tuple transform; ignoring it for intersection-derived point (placePoint={placePoint}).");
                            }
                            else if (hostIsLinked)
                            {
                                // Host lives in a linked document and we didn't receive a tuple transform.
                                // We may still locate a RevitLinkInstance for diagnostics, but we must
                                // NOT apply hostTransform.OfPoint to an active-doc intersection point
                                // (that would double-transform). Instead, compare the projected and
                                // intersection points in active-doc coords directly and choose the
                                // most faithful one.
                                Transform? hostTransform = FindTransformForLinkedElement(hostElem);
                                if (hostTransform != null)
                                {
                                    try
                                    {
                                        double diff = placePoint.DistanceTo(intersectionPoint);
                                        double diffMM = UnitUtils.ConvertFromInternalUnits(diff, UnitTypeId.Millimeters);
                                        _log($"[TransformDebug]   - projected (active): {placePoint}");
                                        _log($"[TransformDebug]   - intersection (active): {intersectionPoint}");
                                        _log($"[TransformDebug]   - projected vs intersection difference = {diffMM:F1}mm");

                                        double thresholdMM = 50.0;
                                        if (diffMM > thresholdMM)
                                        {
                                            placePtToUse = intersectionPoint;
                                            _log($"[TransformDebug] NOTE: projected and intersection differ by {diffMM:F1}mm; using intersection point for placement.");
                                        }
                                        else
                                        {
                                            placePtToUse = placePoint;
                                        }

                                        _log($"[TransformDebug] SUCCESS: Selected placePtToUse: {placePtToUse}");
                                    }
                                    catch (Exception ex)
                                    {
                                        _log($"[TransformDebug] ERROR while comparing projected/intersection: {ex.Message}");
                                    }
                                }
                                else
                                {
                                    _log($"[TransformDebug] WARNING: Could not find RevitLinkInstance for host document '{hostElem?.Document?.Title}'. Using intersection point.");
                                    placePtToUse = intersectionPoint;
                                }
                            }

                            double distMM = UnitUtils.ConvertFromInternalUnits(placePtToUse.DistanceTo(placePoint), UnitTypeId.Millimeters);
                            double dzMM = UnitUtils.ConvertFromInternalUnits(placePtToUse.Z - placePoint.Z, UnitTypeId.Millimeters);
                            _log($"[PlacementDebug] Final placement decision:");
                            _log($"[PlacementDebug]   - Will place at: {placePtToUse} (delta={distMM:F1}mm, dz={dzMM:F1}mm)");

                            double indivTol = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters); // 10mm for individual sleeves
                            double clusterTol = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm for clusters

                            // OPTIMIZATION: Use the enhanced duplication suppressor which first checks nearby
                            // individual sleeves and then cluster bounding boxes, while accepting a hostType
                            // and sectionBox to avoid scanning unrelated cluster families.
                            _log($"[DuplicationCheck] Starting optimized checks for location {placePtToUse} (individual tol={UnitUtils.ConvertFromInternalUnits(indivTol, UnitTypeId.Millimeters):F0}mm, cluster tol={UnitUtils.ConvertFromInternalUnits(clusterTol, UnitTypeId.Millimeters):F0}mm)");
                            // Correct hostTypeFilter for pipes (was incorrectly using Duct names due to copy-paste)
                            string hostTypeFilter = hostElem is Wall ? "PipeOpeningOnWall" : (hostElem is Floor ? "PipeOpeningOnSlab" : "PipeOpeningOnWall");
                            _log($"[DuplicationCheck] using hostTypeFilter={hostTypeFilter}");

                            var nearbySleeves = sleeveGrid.GetNearbySleeves(placePtToUse, indivTol > clusterTol ? indivTol : clusterTol);
                            bool duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationOptimized(placePtToUse, indivTol, clusterTol, nearbySleeves, hostTypeFilter);

                            // Optimized check pre-filters by nearby sleeve center points which can miss
                            // large rectangular cluster sleeves whose center point is outside the
                            // search radius but whose bounding box still covers the placement point.
                            // As a defensive fallback, if the optimized check reports no duplicate,
                            // perform a document-level cluster bounding-box check to be certain.
                            if (!duplicateExists)
                            {
                                try
                                {
                                    // Use the active view's section box if available to limit the doc scan
                                    BoundingBoxXYZ? sectionBoxForDoc = null;
                                    if (_doc != null)
                                    {
                                        try { if (_doc.ActiveView is View3D vb2) sectionBoxForDoc = SectionBoxHelper.GetSectionBoxBounds(vb2); } catch { }
                                        bool clusterBBoxHit = OpeningDuplicationChecker.IsLocationWithinClusterBounds(_doc, placePtToUse, clusterTol, hostTypeFilter, sectionBoxForDoc);
                                        if (clusterBBoxHit)
                                        {
                                            _log($"SKIP: Pipe {pipe.Id} host {hostTypeStr} {hostIdStr} suppressed by existing cluster bounding box at {placePtToUse} (fallback doc-level check)");
                                            SkippedCount++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        _log("[DuplicationCheck] Warning: _doc is null; skipping doc-level fallback check");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _log($"[DuplicationCheck] Fallback cluster bbox check failed: {ex.Message}");
                                }
                            }

                            if (duplicateExists)
                            {
                                _log($"SKIP: Pipe {pipe.Id} host {hostTypeStr} {hostIdStr} duplicate sleeve (individual or cluster) exists near {placePtToUse} (optimized check)");
                                SkippedCount++;
                                continue;
                            }
                            FamilySymbol symbolToUse = (isWall || isFraming) ? _pipeWallSymbol : _pipeSlabSymbol;
                            
                            _placer.PlaceSleeve(
                                pipe,
                                placePtToUse,
                                pipeLine.Direction,
                                symbolToUse,
                                hostElem!
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

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }

        /// <summary>
        /// Finds the RevitLinkInstance transform for a linked element
        /// </summary>
        private Transform? FindTransformForLinkedElement(Element? linkedElement)
        {
            if (linkedElement?.Document == null || _doc == null)
                return null;
            
            try
            {
                // Find all RevitLinkInstances in the active document
                var linkInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                // Find the link instance that points to the same document as the linked element
                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null && 
                        string.Equals(linkDoc.Title, linkedElement.Document.Title, StringComparison.OrdinalIgnoreCase))
                    {
                        var transform = linkInstance.GetTotalTransform();
                        _log($"[TransformDebug] Found matching RevitLinkInstance for document '{linkDoc.Title}'");
                        _log($"[TransformDebug]   - Transform origin: {transform.Origin}");
                        return transform;
                    }
                }
                
                _log($"[TransformDebug] No matching RevitLinkInstance found for document '{linkedElement.Document.Title}'");
                _log($"[TransformDebug] Available linked documents:");
                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    _log($"[TransformDebug]   - '{linkDoc?.Title ?? "NULL"}'");
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _log($"[TransformDebug] ERROR in FindTransformForLinkedElement: {ex.Message}");
                return null;
            }
        }
    }
}

// Local helpers copied from ProgressiveMepSleeveService to keep PipeSleevePlacerService self-contained
namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static partial class PipeSleevePlacerServiceHelpers
    {
        public static bool PointInBoundingBox(XYZ pt, BoundingBoxXYZ bbox)
        {
            return pt.X >= bbox.Min.X - 1e-6 && pt.X <= bbox.Max.X + 1e-6 &&
                   pt.Y >= bbox.Min.Y - 1e-6 && pt.Y <= bbox.Max.Y + 1e-6 &&
                   pt.Z >= bbox.Min.Z - 1e-6 && pt.Z <= bbox.Max.Z + 1e-6;
        }

        public static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
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
            return new BoundingBoxXYZ { Min = min, Max = max };
        }

        public static XYZ OffsetVector4Way(string side, Transform damperT)
        {
            return side switch
            {
                "Right" => XYZ.BasisX,
                "Left" => -XYZ.BasisX,
                "Top" => XYZ.BasisZ,
                "Bottom" => -XYZ.BasisZ,
                _ => XYZ.BasisX,
            };
        }
    }
}
