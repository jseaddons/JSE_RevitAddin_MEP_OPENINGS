using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Provides methods to check for duplicate or overlapping cluster openings in the model.
    /// </summary>
    public static class OpeningDuplicationChecker
    {
    // Simple in-memory cache keyed by document path + active view id to reduce repeated enumeration
    // of RevitLinkInstance and noisy logging during tight loops. Store ElementId integers to allow
    // re-resolving instances against the current Document on cache hit.
    private static readonly Dictionary<string, List<int>> _visibleLinksCache = new Dictionary<string, List<int>>();

        /// <summary>
        /// Returns all cluster openings of the same type within tolerance of the given location.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="clusterSymbol">The FamilySymbol of the cluster being placed</param>
        /// <returns>List of FamilyInstance clusters within tolerance</returns>
        public static List<FamilyInstance> FindClustersAtLocation(Document doc, XYZ location, double tolerance, FamilySymbol clusterSymbol)
        {
            if (doc == null)
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] FindClustersAtLocation: called with null Document");
                return new List<FamilyInstance>();
            }
            if (clusterSymbol == null || clusterSymbol.Family == null)
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] FindClustersAtLocation: called with null clusterSymbol or clusterSymbol.Family");
                return new List<FamilyInstance>();
            }
            DebugLogger.Log($"[OpeningDuplicationChecker] FindClustersAtLocation: doc='{doc.Title}' location={location} tolerance={UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F1}mm family='{clusterSymbol.Family.Name}'");

            // Restrict search to active 3D section box if available
            var sectionBox = GetActiveSectionBoxForDocument(doc);
            var baseCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));
            var collector = baseCollector
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name == clusterSymbol.Family.Name);

            var found = new List<FamilyInstance>();
            foreach (var fi in collector)
            {
                var fiLocation = (fi.Location as LocationPoint)?.Point;
                if (fiLocation == null) continue;
                // If a section box is provided, skip instances whose bounding box does not intersect it
                if (sectionBox != null)
                {
                    var fiBb = fi.get_BoundingBox(null);
                    if (fiBb == null) continue;
                    if (!BoundingBoxesIntersect(fiBb, sectionBox)) continue;
                }
                var dist = location.DistanceTo(fiLocation);
                var distMm = UnitUtils.ConvertFromInternalUnits(dist, UnitTypeId.Millimeters);
                DebugLogger.Log($"[OpeningDuplicationChecker] Candidate cluster: {fi.Symbol.Family.Name} - {fi.Symbol.Name} (ID:{fi.Id.IntegerValue}) at {fiLocation} (dist {distMm:F1}mm)");
                if (dist <= tolerance)
                    found.Add(fi);
            }
            // Log diagnostic counts
            if (found.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] FindIndividualSleevesAtLocation: Found {found.Count} sleeves near {location}");
            }
            return found;
        }

        // Helper: return active view section box in world coords if available for the document's UIDocument
        private static BoundingBoxXYZ? GetActiveSectionBoxForDocument(Document doc)
        {
            try
            {
                // Attempt to get UIDocument from the active application for this document
                UIApplication uiapp = new UIApplication(doc.Application); // fallback, may not be ideal in all environments
                UIDocument? uidoc = uiapp?.ActiveUIDocument;
                if (uidoc == null) return null;

                if (!(uidoc.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
                    return null;

                return SectionBoxHelper.GetSectionBoxBounds(view3D);
            }
            catch
            {
                return null;
            }
        }

        // Helper: transform a host-section bounding box into the coordinate space of another doc using an inverse transform
        private static BoundingBoxXYZ? TransformBoundingBoxForDoc(BoundingBoxXYZ? box, Transform? inverseTransform)
        {
            if (box == null || inverseTransform == null) return box;
            try
            {
                var newMin = inverseTransform.OfPoint(box.Min);
                var newMax = inverseTransform.OfPoint(box.Max);
                return new BoundingBoxXYZ
                {
                    Min = new XYZ(Math.Min(newMin.X, newMax.X), Math.Min(newMin.Y, newMax.Y), Math.Min(newMin.Z, newMax.Z)),
                    Max = new XYZ(Math.Max(newMin.X, newMax.X), Math.Max(newMin.Y, newMax.Y), Math.Max(newMin.Z, newMax.Z)),
                    Transform = Transform.Identity
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Returns true if any cluster (of the given type) exists at the location, excluding those in the ignore list.
        /// </summary>
        public static bool IsClusterAtLocation(Document doc, XYZ location, double tolerance, FamilySymbol clusterSymbol, IEnumerable<ElementId>? ignoreIds = null)
        {
            var clusters = FindClustersAtLocation(doc, location, tolerance, clusterSymbol);
            if (ignoreIds != null)
                clusters = clusters.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
            return clusters.Any();
        }

        /// <summary>
        /// Returns all individual sleeves (non-cluster) within tolerance of the given location.
        /// Individual sleeves are families containing "OpeningOnWall" or "OpeningOnSlab" but NOT containing "cluster" (case-insensitive).
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="sleeveType">Type of sleeve to check for (e.g., "DS#", "PS#", "CTS#")</param>
        /// <returns>List of FamilyInstance individual sleeves within tolerance</returns>
        public static List<FamilyInstance> FindIndividualSleevesAtLocation(Document doc, XYZ location, double tolerance, string? sleeveType = null)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] FindIndividualSleevesAtLocation: called with null Document");
                return new List<FamilyInstance>();
            }
            // If tolerance is not explicitly set (or set too small), default to 10mm for individual-to-individual suppression
            double minTolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);
            // The original logic incorrectly reduced large tolerances to the minimum. Instead,
            // ensure tolerance is at least the minimum (with small floating allowance).
            if (tolerance < minTolerance * 1.1)
                tolerance = minTolerance;

            var allSleeves = new List<FamilyInstance>();

            // Host document
            DebugLogger.Log($"[OpeningDuplicationChecker] FindIndividualSleevesAtLocation: host doc='{doc.Title}' location={location} tolerance={UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F1}mm sleeveType='{sleeveType}'");
            DebugLogger.Log($"[OpeningDuplicationChecker]   -> Host family filter rules: individual families end with OnWall/OnSlab; cluster families start with 'Cluster'");
            var hostSectionBox = GetActiveSectionBoxForDocument(doc);
            var hostSleeves = FindIndividualSleevesInDoc(doc!, location, tolerance, sleeveType, null, hostSectionBox);
            allSleeves.AddRange(hostSleeves);

            // NOTE: Sleeves live only in the active document for our workflows. Do not search linked documents.

            return allSleeves;
        }

        private static IEnumerable<FamilyInstance> FindIndividualSleevesInDoc(Document doc, XYZ location, double tolerance, string? sleeveType, Transform? transform = null, BoundingBoxXYZ? sectionBox = null)
        {
            var baseCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));

            var collector = baseCollector
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    var fam = fi.Symbol?.Family?.Name ?? string.Empty;
                    // Individual sleeves: family names ending with OnWall, OnWALLx or OnSlab (case-insensitive)
                    bool isIndividual = fam.EndsWith("OnWall", StringComparison.OrdinalIgnoreCase)
                                        || fam.EndsWith("OnSlab", StringComparison.OrdinalIgnoreCase);
                    // Exclude cluster families that start with "Cluster"
                    bool isClusterNamed = fam.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase);
                    return isIndividual && !isClusterNamed;
                });

            if (!string.IsNullOrEmpty(sleeveType))
            {
                collector = collector.Where(fi => fi.Symbol.Name.StartsWith(sleeveType, StringComparison.OrdinalIgnoreCase));
            }

            var found = new List<FamilyInstance>();
            DebugLogger.Log($"[OpeningDuplicationChecker] FindIndividualSleevesInDoc: doc='{doc?.Title}' transformProvided={(transform != null)} sleeveType='{sleeveType}' sectionBoxProvided={(sectionBox != null)}");
            foreach (var fi in collector)
            {
                // If a section box is provided, skip instances whose bounding box does not intersect it
                if (sectionBox != null)
                {
                    var fiBbCheck = fi.get_BoundingBox(null);
                    if (fiBbCheck == null) continue;
                    if (!BoundingBoxesIntersect(fiBbCheck, sectionBox)) continue;
                }

                var originalLocation = (fi.Location as LocationPoint)?.Point;
                if (originalLocation == null) continue;
                var fiLocation = originalLocation;
                if (transform != null)
                {
                    var transformed = transform.OfPoint(originalLocation);
                    fiLocation = transformed;
                }

                var dist = location.DistanceTo(fiLocation);
                var distMm = UnitUtils.ConvertFromInternalUnits(dist, UnitTypeId.Millimeters);
                
                // OPTIMIZATION: Only log candidates that are reasonably close to reduce log noise
                // Reduce the logging window to 2x tolerance to avoid many far-away candidates
                bool isClose = dist <= tolerance * 2;
                
                if (transform != null && isClose)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Candidate sleeve: {fi.Symbol.Family.Name} - {fi.Symbol.Name} (ID:{fi.Id.IntegerValue}) original={originalLocation} transformed={fiLocation}");
                }
                else if (transform == null && isClose)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Candidate sleeve: {fi.Symbol.Family.Name} - {fi.Symbol.Name} (ID:{fi.Id.IntegerValue}) at {fiLocation}");
                }
                
                if (isClose)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker]   - distance to location: {distMm:F1}mm (tolerance {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F1}mm)");
                }

                if (dist <= tolerance)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker]   -> Within tolerance, will be considered a duplicate (ID:{fi.Id.IntegerValue})");
                    found.Add(fi);
                }
            }
            return found;
        }

        /// <summary>
        /// Returns all cluster sleeves (rectangular/cluster openings) within tolerance of the given location.
        /// Cluster sleeves are families ending with "Rect" (e.g., PipeOpeningOnWallRect, DuctOpeningOnWallRect).
        /// Uses bounding box intersection for accurate detection since cluster sleeves can be large.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <returns>List of FamilyInstance cluster sleeves within tolerance</returns>
    public static List<FamilyInstance> FindAllClusterSleevesAtLocation(Document doc, XYZ location, double tolerance, string? hostType = null)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] FindAllClusterSleevesAtLocation: called with null Document");
                return new List<FamilyInstance>();
            }
            var allSleeves = new List<FamilyInstance>();

            // Host document (restrict to active section box if present)
            var hostSectionBox = GetActiveSectionBoxForDocument(doc);
            allSleeves.AddRange(FindAllClusterSleevesInDoc(doc, location, tolerance, null, hostSectionBox, hostType));

            // NOTE: For performance and correctness our workflows only consider sleeves in the active document.

            return allSleeves;
        }

    private static IEnumerable<FamilyInstance> FindAllClusterSleevesInDoc(Document doc, XYZ location, double tolerance, Transform? transform = null, BoundingBoxXYZ? sectionBox = null, string? hostType = null)
        {
            var baseCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance));

            var collector = baseCollector
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    var fam = fi.Symbol?.Family?.Name ?? string.Empty;
                    // Cluster families are those starting with "Cluster" (case-insensitive)
                    // or rectangular cluster families that end with "Rect" (legacy behavior)
                    return fam.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase)
                        || fam.EndsWith("Rect", StringComparison.OrdinalIgnoreCase);
                });

            var found = new List<FamilyInstance>();
            DebugLogger.Log($"[OpeningDuplicationChecker] FindAllClusterSleevesInDoc: doc='{doc?.Title}' transformProvided={(transform != null)} collectorCount={collector.Count()} sectionBoxProvided={(sectionBox != null)} hostType={(hostType ?? "<any>")}");
            foreach (var fi in collector)
            {
                // If a section box is provided, skip instances whose bounding box does not intersect it
                if (sectionBox != null)
                {
                    var fiBb = fi.get_BoundingBox(null);
                    if (fiBb == null) continue;
                    if (!BoundingBoxesIntersect(fiBb, sectionBox)) continue;
                }
                // If hostType provided, skip non-matching families (wall vs slab clusters)
                if (!string.IsNullOrEmpty(hostType))
                {
                    var fam = fi.Symbol?.Family?.Name ?? "";
                    if (!fam.Contains(hostType, StringComparison.OrdinalIgnoreCase))
                        continue;
                }
                var originalLocation = (fi.Location as LocationPoint)?.Point;
                if (originalLocation == null) continue;
                var fiLocation = originalLocation;
                if (transform != null)
                {
                    fiLocation = transform.OfPoint(originalLocation);
                    var famName = fi.Symbol?.Family?.Name ?? "<unknown family>";
                    var symName = fi.Symbol?.Name ?? "<unknown symbol>";
                    DebugLogger.Log($"[OpeningDuplicationChecker] Candidate cluster: {famName} - {symName} (ID:{fi.Id.IntegerValue}) original={originalLocation} transformed={fiLocation}");
                }
                else
                {
                    var famName = fi.Symbol?.Family?.Name ?? "<unknown family>";
                    var symName = fi.Symbol?.Name ?? "<unknown symbol>";
                    DebugLogger.Log($"[OpeningDuplicationChecker] Candidate cluster: {famName} - {symName} (ID:{fi.Id.IntegerValue}) at {fiLocation}");
                }

                var dist = location.DistanceTo(fiLocation);
                var distMm = UnitUtils.ConvertFromInternalUnits(dist, UnitTypeId.Millimeters);
                DebugLogger.Log($"[OpeningDuplicationChecker]   - distance to location: {distMm:F1}mm (tolerance {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F1}mm)");

                if (dist <= tolerance)
                    found.Add(fi);
            }
            return found;
        }

        private static IEnumerable<RevitLinkInstance> GetVisibleLinkedDocuments(Document doc)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] GetVisibleLinkedDocuments: called with null Document");
                return Enumerable.Empty<RevitLinkInstance>();
            }
            
            if (doc.ActiveView == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] GetVisibleLinkedDocuments: ActiveView is null");
                return Enumerable.Empty<RevitLinkInstance>();
            }
            // Build a cache key from document path (falls back to title) and active view id so we refresh when the view changes
            var cacheKey = string.Concat(doc.PathName ?? doc.Title, "|", doc.ActiveView?.Id.IntegerValue.ToString() ?? "0");

            // If we have a cached list of link instance ids for this doc+view, try to re-resolve them and return quickly
            if (_visibleLinksCache.TryGetValue(cacheKey, out var cachedIds))
            {
                var resolved = new List<RevitLinkInstance>();
                foreach (var idInt in cachedIds)
                {
                    try
                    {
                        var el = doc.GetElement(new ElementId(idInt)) as RevitLinkInstance;
                        if (el != null)
                            resolved.Add(el);
                    }
                    catch { }
                }
                DebugLogger.Log($"[OpeningDuplicationChecker] GetVisibleLinkedDocuments: returning {resolved.Count} cached visible links for view '{doc.ActiveView?.Name}'");
                return resolved;
            }

            // Cache miss - enumerate and compute visible links, then store their ids
            var allLinks = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            DebugLogger.Log($"[OpeningDuplicationChecker] GetVisibleLinkedDocuments: Found {allLinks.Count} total RevitLinkInstances in doc");
            var activeViewName = doc.ActiveView?.Name ?? "(null)";
            var activeViewType = doc.ActiveView?.GetType().Name ?? "(null)";
            DebugLogger.Log($"[OpeningDuplicationChecker] Active view: '{activeViewName}' (Type: {activeViewType})");

            var visibleLinks = allLinks.Where(link =>
            {
                var linkDoc = link.GetLinkDocument();
                bool hasDoc = linkDoc != null;
                bool categoryVisible = true;
                bool elementVisible = true;
                try
                {
                    if (doc.ActiveView != null)
                        categoryVisible = !doc.ActiveView.GetCategoryHidden(link.Category.Id);
                }
                catch { }
                try
                {
                    if (doc.ActiveView != null)
                        elementVisible = !link.IsHidden(doc.ActiveView);
                }
                catch { }

                DebugLogger.Log($"[OpeningDuplicationChecker] Link '{linkDoc?.Title}': hasDoc={hasDoc}, categoryVisible={categoryVisible}, elementVisible={elementVisible}");

                return hasDoc && categoryVisible && elementVisible;
            }).ToList();

            DebugLogger.Log($"[OpeningDuplicationChecker] Returning {visibleLinks.Count} visible linked documents");

            try
            {
                _visibleLinksCache[cacheKey] = visibleLinks.Select(l => l.Id.IntegerValue).ToList();
            }
            catch
            {
                // Ignore caching failures
            }

            return visibleLinks;
        }

        /// <summary>
        /// Checks if a point is within the physical bounds of any cluster sleeve (using bounding box intersection).
        /// This is more accurate than center-point distance for large rectangular cluster openings.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="expansionTolerance">Additional tolerance to expand cluster bounding boxes (internal units)</param>
        /// <returns>True if location is within any cluster sleeve's bounding box</returns>
    public static bool IsLocationWithinClusterBounds(Document doc, XYZ location, double expansionTolerance = 0.0, string? hostType = null, BoundingBoxXYZ? sectionBox = null)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] IsLocationWithinClusterBounds: called with null Document");
                return false;
            }
            // Respect caller-supplied section box when provided, otherwise fall back to active view's section box
            if (sectionBox == null)
                sectionBox = GetActiveSectionBoxForDocument(doc);
            var baseCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));

            var clusterSleeves = baseCollector
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    var fam = fi.Symbol?.Family?.Name ?? string.Empty;
                    // Only treat families starting with 'Cluster' as cluster families (case-insensitive)
                    // and preserve rectangular cluster families that end with 'Rect'. Avoid generic contains('cluster')
                    return fam.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase)
                        || fam.EndsWith("Rect", StringComparison.OrdinalIgnoreCase);
                });

            foreach (var cluster in clusterSleeves)
            {
                try
                {
                    // If a hostType filter is provided, skip clusters that don't match (e.g. wall vs slab clusters)
                    if (!string.IsNullOrEmpty(hostType))
                    {
                        var familyName = cluster.Symbol?.Family?.Name ?? "";
                        if (!familyName.Contains(hostType, StringComparison.OrdinalIgnoreCase))
                        {
                            // skip unrelated cluster (reduces noise and cost)
                            continue;
                        }
                    }
                    var clusterFamilyName = cluster.Symbol?.Family?.Name ?? "<unknown family>";
                    DebugLogger.Log($"[OpeningDuplicationChecker] Checking cluster bounding box: {clusterFamilyName} (ID:{cluster.Id.IntegerValue}) expansionTolerance={UnitUtils.ConvertFromInternalUnits(expansionTolerance, UnitTypeId.Millimeters):F1}mm");
                    var boundingBox = cluster.get_BoundingBox(null);
                        if (boundingBox != null)
                        {
                            // Expand bounding box by tolerance if specified
                            var expandedMin = boundingBox.Min - new XYZ(expansionTolerance, expansionTolerance, expansionTolerance);
                            var expandedMax = boundingBox.Max + new XYZ(expansionTolerance, expansionTolerance, expansionTolerance);
                            // Diagnostic logging for bounding box vs location
                            DebugLogger.Log($"[OpeningDuplicationChecker] Cluster bbox (expanded) for ID:{cluster.Id.IntegerValue} min={expandedMin} max={expandedMax}; checking location={location}");

                            // Use a 2D XY check for cluster membership (clusters are typically planar in XY).
                            // Z can vary due to thin family instances or placement offsets; using XY avoids false negatives
                            // when Z differs slightly between the placement point and the cluster family.
                            bool insideXY = location.X >= expandedMin.X && location.X <= expandedMax.X &&
                                            location.Y >= expandedMin.Y && location.Y <= expandedMax.Y;

                            if (insideXY)
                            {
                                var clusterFamilyName2 = cluster.Symbol?.Family?.Name ?? "<unknown family>";
                                DebugLogger.Log($"[OpeningDuplicationChecker] Location {location} is within cluster {clusterFamilyName2} (ID:{cluster.Id.IntegerValue}) XY bounds (min={expandedMin} max={expandedMax})");
                                return true;
                            }
                        }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Warning: Could not check bounding box for cluster {cluster.Id.IntegerValue}: {ex.Message}");
                }
            }
            return false;
        }

        // Axis-aligned bounding-box intersection test
        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }

        /// <summary>
        /// Comprehensive check: Returns true if ANY sleeve (individual OR cluster) exists at the location.
        /// This is the method that should be used in sleeve placement commands to prevent conflicts.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="ignoreIds">ElementIds to ignore in the check</param>
        /// <returns>True if any sleeve (individual or cluster) exists at the location</returns>
        public static bool IsAnySleeveAtLocation(Document doc, XYZ location, double tolerance, IEnumerable<ElementId>? ignoreIds = null)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] IsAnySleeveAtLocation: called with null Document");
                return false;
            }
            // Check for individual sleeves
            var individualSleeves = FindIndividualSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                individualSleeves = individualSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();

            if (individualSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {individualSleeves.Count} individual sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                return true;
            }

            // Check for cluster sleeves
            var clusterSleeves = FindAllClusterSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                clusterSleeves = clusterSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();

            if (clusterSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {clusterSleeves.Count} cluster sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                foreach (var cluster in clusterSleeves)
                {
                    DebugLogger.Log($"  - Cluster: {cluster.Symbol.Family.Name} - {cluster.Symbol.Name} (ID:{cluster.Id.IntegerValue})");
                }
                return true;
            }

            return false;
        }

        /// <summary>
        /// Enhanced comprehensive check: Returns true if ANY sleeve (individual OR cluster) exists at the location.
        /// Uses bounding box intersection for cluster sleeves and distance checking for individual sleeves.
        /// This is the most accurate method for sleeve placement duplication checking.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance for individual sleeves (internal units)</param>
        /// <param name="clusterExpansion">Additional expansion for cluster bounding boxes (internal units)</param>
        /// <param name="ignoreIds">ElementIds to ignore in the check</param>
        /// <returns>True if any sleeve (individual or cluster) exists at the location</returns>
    public static bool IsAnySleeveAtLocationEnhanced(Document doc, XYZ location, double tolerance, double clusterExpansion = 0.0, IEnumerable<ElementId>? ignoreIds = null, string? hostType = null, BoundingBoxXYZ? sectionBox = null, bool requireSameFamily = false, string? familyName = null)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] IsAnySleeveAtLocationEnhanced: called with null Document");
                return false;
            }
            // Diagnostic: log the hostType and whether a sectionBox was provided so callers can see why sleeveType may be empty
            DebugLogger.Log($"[OpeningDuplicationChecker] IsAnySleeveAtLocationEnhanced: hostType={(hostType ?? "<none>")}, sectionBoxProvided={(sectionBox != null)}, requireSameFamily={requireSameFamily}, familyName={(familyName ?? "<none>")}");
            // Check for individual sleeves using distance
            var individualSleeves = FindIndividualSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                individualSleeves = individualSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();

            if (individualSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {individualSleeves.Count} individual sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                if (requireSameFamily && !string.IsNullOrEmpty(familyName))
                {
                    var matching = individualSleeves.Where(fi => string.Equals(fi.Symbol?.Family?.Name ?? "", familyName, StringComparison.OrdinalIgnoreCase)).ToList();
                    if (matching.Any())
                    {
                        DebugLogger.Log($"[OpeningDuplicationChecker] {matching.Count} individual sleeves match required family '{familyName}'");
                        return true;
                    }
                    else
                    {
                        DebugLogger.Log($"[OpeningDuplicationChecker] No individual sleeves matched required family '{familyName}' (found {individualSleeves.Count} other family sleeve(s))");
                        // continue to cluster check
                    }
                }
                else
                {
                    return true;
                }
            }

            // Check for cluster sleeves using bounding box intersection (more accurate)
            if (!requireSameFamily || string.IsNullOrEmpty(familyName))
            {
                if (IsLocationWithinClusterBounds(doc, location, clusterExpansion, hostType, sectionBox))
                {
                    return true;
                }
            }
            else
            {
                // When requiring same family, fetch cluster sleeves and verify family equality
                var clusterList = FindAllClusterSleevesAtLocation(doc, location, clusterExpansion, hostType);
                if (ignoreIds != null)
                    clusterList = clusterList.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
                var matchingClusters = clusterList.Where(fi => string.Equals(fi.Symbol?.Family?.Name ?? "", familyName, StringComparison.OrdinalIgnoreCase)).ToList();
                if (matchingClusters.Any())
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Found {matchingClusters.Count} cluster sleeve(s) matching required family '{familyName}'");
                    return true;
                }
                else
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] No cluster sleeves matched required family '{familyName}'");
                }
            }

            return false;
        }

        /// <summary>
        /// Returns a summary of what sleeve types exist at the given location for debugging purposes.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <returns>Description of sleeves found at location</returns>
        public static string GetSleevesSummaryAtLocation(Document doc, XYZ location, double tolerance)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] GetSleevesSummaryAtLocation: called with null Document");
                return "Document not loaded";
            }
            var individual = FindIndividualSleevesAtLocation(doc, location, tolerance);
            var clusters = FindAllClusterSleevesAtLocation(doc, location, tolerance);

            var summary = new StringBuilder();
            if (individual.Any())
            {
                summary.AppendLine($"Individual sleeves: {string.Join(", ", individual.Select(s => $"{s.Symbol.Name} (ID:{s.Id.IntegerValue})"))}");
            }
            if (clusters.Any())
            {
                summary.AppendLine($"Cluster sleeves: {string.Join(", ", clusters.Select(s => $"{s.Symbol.Name} (ID:{s.Id.IntegerValue})"))}");
            }
            if (!individual.Any() && !clusters.Any())
            {
                summary.AppendLine("No sleeves found at location");
            }
            return summary.ToString().TrimEnd();
        }

        /// <summary>
        /// Finds all cable tray sleeves in the document.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <returns>A list of FamilyInstance objects representing cable tray sleeves.</returns>
        public static List<FamilyInstance> FindCableTraySleeves(Document doc)
        {
            if (doc == null)
            {
                DebugLogger.Log("[OpeningDuplicationChecker] FindCableTraySleeves: called with null Document");
                return new List<FamilyInstance>();
            }
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                    (fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab"))
                    && fi.Symbol.Name.StartsWith("CT#", StringComparison.OrdinalIgnoreCase)
                )
                .ToList();
        }

    }
}

