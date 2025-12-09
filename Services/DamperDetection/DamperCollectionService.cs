using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for collecting damper elements from documents.
    /// Single Responsibility: Collect dampers from Duct Accessories category without expensive intersection calculations.
    /// 
    /// This is much cheaper than using MepIntersectionService because:
    /// - No intersection calculations needed
    /// - Just filter Duct Accessories by family name
    /// - For Standard dampers: Use bounding box centroid as placement point
    /// - For Non-standard dampers (MSFD, MSD, MD, Motorized): Use connector Origin as placement point
    /// </summary>
    public class DamperCollectionService
    {
        private readonly IDamperTypeDetector _damperTypeDetector;
        private readonly IDamperConnectorDetector _connectorDetector;

        public DamperCollectionService(IDamperTypeDetector damperTypeDetector, IDamperConnectorDetector connectorDetector)
        {
            _damperTypeDetector = damperTypeDetector ?? throw new ArgumentNullException(nameof(damperTypeDetector));
            _connectorDetector = connectorDetector ?? throw new ArgumentNullException(nameof(connectorDetector));
        }

        /// <summary>
        /// Collects Standard and Motorized dampers from host and linked documents.
        /// Uses Duct Accessories category filter - much cheaper than intersection calculations.
        /// </summary>
        /// <param name="doc">The document to collect from</param>
        /// <param name="log">Optional logging action</param>
        /// <returns>List of (damper element, transform, placement point as bbox centroid)</returns>
        public List<(Element damper, Transform? transform, XYZ placementPoint)> CollectDampers(
            Document doc, 
            Action<string>? log = null)
        {
            var results = new List<(Element, Transform?, XYZ)>();
            log?.Invoke("[DamperCollection] Starting cheap damper collection from Duct Accessories category.");

            // Collect from host document
            CollectDampersFromDocument(doc, null, results, log);

            // Collect from linked documents
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var link in links)
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                var linkTransform = link.GetTotalTransform();
                CollectDampersFromDocument(linkDoc, linkTransform, results, log);
            }

            log?.Invoke($"[DamperCollection] Finished. Total dampers found: {results.Count}");
            return results;
        }

        /// <summary>
        /// Collects dampers from a single document (host or linked).
        /// Filters by Duct Accessories category and Standard/Motorized damper types.
        /// </summary>
        private void CollectDampersFromDocument(
            Document document,
            Transform? linkTransform,
            List<(Element, Transform?, XYZ)> results,
            Action<string>? log)
        {
            try
            {
                // ✅ CHEAP: Filter by Duct Accessories category only
                var ductAccessories = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => IsStandardOrMotorizedDamper(fi))
                    .ToList();

                log?.Invoke($"[DamperCollection] Found {ductAccessories.Count} Standard/Motorized dampers in document: {document.Title}");

                foreach (var damper in ductAccessories)
                {
                    // ✅ STEP 1: Detect damper type
                    string familyTypeName = damper.Symbol?.Name ?? "";
                    string damperType = _damperTypeDetector.DetectDamperType(familyTypeName);
                    bool isStandard = _damperTypeDetector.IsStandardDamper(damperType);

                    XYZ placementPoint;

                    // ✅ STEP 2: Calculate placement point based on damper type
                    if (isStandard)
                    {
                        // ✅ STANDARD DAMPERS: Use bounding box centroid (DamperPlacementStrategy approach)
                        var bbox = damper.get_BoundingBox(null);
                        if (bbox == null)
                        {
                            log?.Invoke($"[DamperCollection] ⚠️ Damper {damper.Id} has no bounding box - skipping");
                            continue;
                        }

                        // Calculate centroid of bounding box
                        placementPoint = new XYZ(
                            (bbox.Min.X + bbox.Max.X) / 2.0,
                            (bbox.Min.Y + bbox.Max.Y) / 2.0,
                            (bbox.Min.Z + bbox.Max.Z) / 2.0
                        );
                    }
                    else
                    {
                        // ✅ NON-STANDARD DAMPERS (MSFD, MSD, MD, Motorized): Use connector Origin
                        // Use DamperConnectorDetector to get connector, then use connector.Origin
                        string connectorSide = _connectorDetector.DetectConnectorSide(
                            damper, 
                            useWorldCoordinates: true, // Non-standard dampers use world coordinates
                            out Connector connector,
                            wallOrientation: null // No wall orientation available at collection time
                        );

                        if (connector != null && connector.Origin != null)
                        {
                            // Use connector Origin as placement point
                            placementPoint = connector.Origin;
                            log?.Invoke($"[DamperCollection] ✅ Non-standard damper {damper.Id} ({damperType}): Using connector Origin as placement point (Side={connectorSide})");
                        }
                        else
                        {
                            // Fallback: Use bounding box centroid if no connector found
                            var bbox = damper.get_BoundingBox(null);
                            if (bbox == null)
                            {
                                log?.Invoke($"[DamperCollection] ⚠️ Damper {damper.Id} has no bounding box and no connector - skipping");
                                continue;
                            }

                            placementPoint = new XYZ(
                                (bbox.Min.X + bbox.Max.X) / 2.0,
                                (bbox.Min.Y + bbox.Max.Y) / 2.0,
                                (bbox.Min.Z + bbox.Max.Z) / 2.0
                            );
                            log?.Invoke($"[DamperCollection] ⚠️ Non-standard damper {damper.Id} ({damperType}): No connector found, using bounding box centroid as fallback");
                        }
                    }

                    // ✅ STEP 3: Transform placement point if from linked document
                    if (linkTransform != null)
                    {
                        placementPoint = linkTransform.OfPoint(placementPoint);
                    }

                    results.Add((damper, linkTransform, placementPoint));
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[DamperCollection] ❌ Error collecting dampers from {document.Title}: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if a FamilyInstance is a damper (any type: Standard, Motorized, MSFD, MSD, MD).
        /// Uses family/type name to detect damper type.
        /// Excludes VCD and VOLUME dampers (not in walls).
        /// </summary>
        private bool IsStandardOrMotorizedDamper(FamilyInstance fi)
        {
            if (fi?.Symbol == null) return false;

            // Get family and type names
            string familyName = fi.Symbol.Family?.Name ?? "";
            string typeName = fi.Symbol.Name ?? "";
            string combinedName = $"{familyName} {typeName}".Trim().ToUpperInvariant();

            // Check if it contains "Damper" keyword
            if (!combinedName.Contains("DAMPER"))
                return false;

            // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
            if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                return false;

            // ✅ COLLECT ALL DAMPER TYPES: Standard, Motorized, MSFD, MSD, MD
            // The placement point calculation will differ based on type (see CollectDampersFromDocument)
            return true;
        }
    }
}

