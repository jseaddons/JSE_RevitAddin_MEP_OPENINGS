using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Plumbing;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Progressive MEP sleeve placement service with user feedback and crash prevention
    /// </summary>
    public class ProgressiveMepSleeveService
    {
        private readonly Document _doc;
        private readonly Action<string> _log;
        private readonly List<(Element, Transform?)> _structuralElements;
    private readonly bool _processDampers;
        
        public int TotalPlaced { get; private set; }
        public int TotalSkipped { get; private set; }
        public int TotalErrors { get; private set; }

        public ProgressiveMepSleeveService(Document doc, Action<string> log, bool processDampers = true)
        {
            _doc = doc;
            _log = log;
            _processDampers = processDampers;
            
            // Collect structural elements once with section box filtering
            _log("Collecting structural elements (walls, floors, framing) within active section box...");
            _structuralElements = MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, _log);
            _log($"Found {_structuralElements.Count} structural elements in active view bounds");
            
            // Clear any existing geometry cache
            MepIntersectionService.ClearGeometryCache();
        }

        public void ProcessAllMepTypes()
        {
            TotalPlaced = 0;
            TotalSkipped = 0;
            TotalErrors = 0;

            // Process each MEP type progressively with user feedback
            ProcessDucts();
            if (_processDampers)
            {
                _log("=== PROCESSING DAMPERS ===");
                ProcessDampers();
                _log($"DAMPER PROCESSING COMPLETED - Placed: {TotalPlaced}, Skipped: {TotalSkipped}, Errors: {TotalErrors}");
            }
            ProcessPipes(); 
            ProcessCableTrays();
            ProcessConduits();
            
            // Final summary
            _log($"=== FINAL SUMMARY ===");
            _log($"Total placed: {TotalPlaced}, Total skipped: {TotalSkipped}, Total errors: {TotalErrors}");
            
            // Clear cache to free memory
            MepIntersectionService.ClearGeometryCache();
        }

        private void ProcessDampers()
        {
            var damperWallSymbol = GetFamilySymbol("DamperOpeningOnWall");
            if (damperWallSymbol == null)
            {
                _log("SKIPPED: No damper sleeve family found");
                return;
            }

            var damperTuples = Helpers.MepElementCollectorHelper.CollectFamilyInstancesVisibleOnly(_doc, "Damper");
            var wallTuples = Helpers.MepElementCollectorHelper.CollectWallsVisibleOnly(_doc);

            // Filter dampers by active section box to avoid scanning the entire project (expensive)
            try
            {
                if (!_doc.IsLinked)
                {
                    var uiDoc = new UIDocument(_doc);
                    // Convert to element/transform pairs expected by SectionBoxHelper
                    var raw = damperTuples.Select(t => ((Element)t.instance, t.transform)).ToList();
                    var filtered = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.FilterElementsBySectionBox(uiDoc, raw);
                    // Convert back to family-instance tuples
                    damperTuples = filtered.Select(t => ((FamilyInstance)t.element, t.transform)).ToList();
                }
            }
            catch (Exception ex)
            {
                _log($"[ProcessDampers] Section-box filtering failed for damper collection: {ex.Message}");
            }

            _log($"Found {damperTuples.Count} dampers and {wallTuples.Count} walls to consider");

            // Delegate to unified placer service to avoid duplicating complex placement logic
            var damperPlacer = new FireDamperSleevePlacerService(_doc, enableDebugLogging: true);

            // Process in chunks to limit transaction size and give progress updates
            const int chunkSize = 10;
            var chunks = damperTuples.Select((x, i) => new { Index = i, Value = x })
                                   .GroupBy(x => x.Index / chunkSize)
                                   .Select(x => x.Select(v => v.Value).ToList())
                                   .ToList();

            int chunkIndex = 0;
            foreach (var chunk in chunks)
            {
                chunkIndex++;
                _log($"Processing DAMPER chunk {chunkIndex}/{chunks.Count} ({chunk.Count} elements)...");
                try
                {
                    using (var tx = new SubTransaction(_doc))
                    {
                        tx.Start();

                        var subDamperList = chunk;
                        var (placed, skipped, errors) = damperPlacer.ProcessDamperBatch(subDamperList, wallTuples, damperWallSymbol, _log, enableDebugging: true, structuralElements: _structuralElements);
                        TotalPlaced += placed;
                        TotalSkipped += skipped;
                        TotalErrors += errors;

                        tx.Commit();
                        _log($"Completed DAMPER chunk {chunkIndex}/{chunks.Count}");
                    }
                }
                catch (Exception ex)
                {
                    TotalErrors++;
                    _log($"ERROR in DAMPER chunk {chunkIndex}: {ex.Message}");
                }
            }
        }

        // Helper: point-in-bbox inclusive
        private static bool PointInBoundingBox(XYZ pt, BoundingBoxXYZ bbox)
        {
            return pt.X >= bbox.Min.X - 1e-6 && pt.X <= bbox.Max.X + 1e-6 &&
                   pt.Y >= bbox.Min.Y - 1e-6 && pt.Y <= bbox.Max.Y + 1e-6 &&
                   pt.Z >= bbox.Min.Z - 1e-6 && pt.Z <= bbox.Max.Z + 1e-6;
        }

        // Helper: transform bounding box by a transform (same approach as command)
        private static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
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

        // Helper: map side string to world direction (same semantics as command)
        private static XYZ OffsetVector4Way(string side, Transform damperT)
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

        private void ProcessDucts()
        {
            _log("=== PROCESSING DUCTS ===");
            
            // Get duct sleeve symbols
            var ductWallSymbol = GetFamilySymbol("DuctOpeningOnWall");
            var ductSlabSymbol = GetFamilySymbol("DuctOpeningOnSlab");
            
            if (ductWallSymbol == null && ductSlabSymbol == null)
            {
                _log("SKIPPED: No duct sleeve families found (DS#)");
                return;
            }

            // Activate symbols
            using (var tx = new SubTransaction(_doc))
            {
                tx.Start();
                if (ductWallSymbol != null && !ductWallSymbol.IsActive) ductWallSymbol.Activate();
                if (ductSlabSymbol != null && !ductSlabSymbol.IsActive) ductSlabSymbol.Activate();
                tx.Commit();
            }

            // Collect ducts
            var ducts = GetMepElementsInSectionBox(BuiltInCategory.OST_DuctCurves);
            _log($"Found {ducts.Count} ducts to process");

            if (ducts.Count > 0)
            {
                ProcessMepElements(ducts, ductWallSymbol, ductSlabSymbol, "DUCT");
            }
            
            _log($"DUCT PROCESSING COMPLETED - Placed: {TotalPlaced}, Skipped: {TotalSkipped}, Errors: {TotalErrors}");
        }

        private void ProcessPipes()
        {
            _log("=== PROCESSING PIPES ===");
            
            var pipeWallSymbol = GetFamilySymbol("PipeOpeningOnWall");
            var pipeSlabSymbol = GetFamilySymbol("PipeOpeningOnSlab");
            
            if (pipeWallSymbol == null && pipeSlabSymbol == null)
            {
                _log("SKIPPED: No pipe sleeve families found (PS#)");
                return;
            }

            using (var tx = new SubTransaction(_doc))
            {
                tx.Start();
                if (pipeWallSymbol != null && !pipeWallSymbol.IsActive) pipeWallSymbol.Activate();
                if (pipeSlabSymbol != null && !pipeSlabSymbol.IsActive) pipeSlabSymbol.Activate();
                tx.Commit();
            }

            var pipes = GetMepElementsInSectionBox(BuiltInCategory.OST_PipeCurves);
            _log($"Found {pipes.Count} pipes to process");

            if (pipes.Count > 0)
            {
                ProcessMepElements(pipes, pipeWallSymbol, pipeSlabSymbol, "PIPE");
            }
            
            _log($"PIPE PROCESSING COMPLETED - Placed: {TotalPlaced}, Skipped: {TotalSkipped}, Errors: {TotalErrors}");
        }

        private void ProcessCableTrays()
        {
            _log("=== PROCESSING CABLE TRAYS ===");
            
            var cableTrayWallSymbol = GetFamilySymbol("CableTrayOpeningOnWall");
            var cableTraySlabSymbol = GetFamilySymbol("CableTrayOpeningOnSlab");
            
            if (cableTrayWallSymbol == null && cableTraySlabSymbol == null)
            {
                _log("SKIPPED: No cable tray sleeve families found (CTS#)");
                return;
            }

            using (var tx = new SubTransaction(_doc))
            {
                tx.Start();
                if (cableTrayWallSymbol != null && !cableTrayWallSymbol.IsActive) cableTrayWallSymbol.Activate();
                if (cableTraySlabSymbol != null && !cableTraySlabSymbol.IsActive) cableTraySlabSymbol.Activate();
                tx.Commit();
            }

            var cableTrays = GetMepElementsInSectionBox(BuiltInCategory.OST_CableTray);
            _log($"Found {cableTrays.Count} cable trays to process");

            if (cableTrays.Count > 0)
            {
                ProcessMepElements(cableTrays, cableTrayWallSymbol, cableTraySlabSymbol, "CABLE TRAY");
            }
            
            _log($"CABLE TRAY PROCESSING COMPLETED - Placed: {TotalPlaced}, Skipped: {TotalSkipped}, Errors: {TotalErrors}");
        }

        private void ProcessConduits()
        {
            _log("=== PROCESSING CONDUITS ===");
            
            var conduitWallSymbol = GetFamilySymbol("ConduitOpeningOnWall");
            var conduitSlabSymbol = GetFamilySymbol("ConduitOpeningOnSlab");
            
            if (conduitWallSymbol == null && conduitSlabSymbol == null)
            {
                _log("SKIPPED: No conduit sleeve families found");
                return;
            }

            using (var tx = new SubTransaction(_doc))
            {
                tx.Start();
                if (conduitWallSymbol != null && !conduitWallSymbol.IsActive) conduitWallSymbol.Activate();
                if (conduitSlabSymbol != null && !conduitSlabSymbol.IsActive) conduitSlabSymbol.Activate();
                tx.Commit();
            }

            var conduits = GetMepElementsInSectionBox(BuiltInCategory.OST_Conduit);
            _log($"Found {conduits.Count} conduits to process");

            if (conduits.Count > 0)
            {
                ProcessMepElements(conduits, conduitWallSymbol, conduitSlabSymbol, "CONDUIT");
            }
            
            _log($"CONDUIT PROCESSING COMPLETED - Placed: {TotalPlaced}, Skipped: {TotalSkipped}, Errors: {TotalErrors}");
        }

        private void ProcessMepElements(List<(Element, Transform?)> mepElements, FamilySymbol? wallSymbol, FamilySymbol? slabSymbol, string mepType)
        {
            const int chunkSize = 10; // Process in small chunks to prevent crashes
            var chunks = mepElements.Select((x, i) => new { Index = i, Value = x })
                                   .GroupBy(x => x.Index / chunkSize)
                                   .Select(x => x.Select(v => v.Value).ToList())
                                   .ToList();

            int chunkIndex = 0;
            foreach (var chunk in chunks)
            {
                chunkIndex++;
                _log($"Processing {mepType} chunk {chunkIndex}/{chunks.Count} ({chunk.Count} elements)...");
                
                try
                {
                    using (var tx = new SubTransaction(_doc))
                    {
                        tx.Start();
                        
                        foreach (var tuple in chunk)
                        {
                            ProcessSingleMepElement(tuple.Item1, wallSymbol, slabSymbol, mepType);
                        }
                        
                        tx.Commit();
                        _log($"Completed {mepType} chunk {chunkIndex}/{chunks.Count}");
                    }
                }
                catch (Exception ex)
                {
                    TotalErrors++;
                    _log($"ERROR in {mepType} chunk {chunkIndex}: {ex.Message}");
                }
            }
        }

        private void ProcessSingleMepElement(Element mepElement, FamilySymbol? wallSymbol, FamilySymbol? slabSymbol, string mepType)
        {
            try
            {
                var intersections = MepIntersectionService.FindIntersections(mepElement, _structuralElements, _log);
                
                if (intersections == null || intersections.Count == 0)
                {
                    TotalSkipped++;
                    return;
                }

                foreach (var intersection in intersections)
                {
                    var hostElement = intersection.Item1;
                    var intersectionPoint = intersection.Item3;
                    
                    // Determine appropriate placer and symbol
                    if (hostElement is Wall && wallSymbol != null)
                    {
                        var placer = new PipeSleevePlacer(_doc); // Reuse existing placer
                        // Attempt to cast the generic element to a Pipe. If the cast fails, PipeSleevePlacer will
                        // handle null pipe gracefully and return false.
                        var pipeAsPipe = mepElement as Pipe;
                        XYZ pipeDir = XYZ.BasisX;
                        try
                        {
                            var loc = (pipeAsPipe?.Location as LocationCurve);
                            var line = loc?.Curve as Line;
                            if (line != null) pipeDir = line.Direction;
                        }
                        catch { }

                        var success = placer.PlaceSleeve(pipeAsPipe, intersectionPoint, pipeDir, wallSymbol, hostElement);
                        if (success) TotalPlaced++; else TotalSkipped++;
                    }
                    else if ((hostElement.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_Floors) && slabSymbol != null)
                    {
                        var placer = new PipeSleevePlacer(_doc);
                        var pipeAsPipe = mepElement as Pipe;
                        XYZ pipeDir = XYZ.BasisX;
                        try
                        {
                            var loc = (pipeAsPipe?.Location as LocationCurve);
                            var line = loc?.Curve as Line;
                            if (line != null) pipeDir = line.Direction;
                        }
                        catch { }

                        var success = placer.PlaceSleeve(pipeAsPipe, intersectionPoint, pipeDir, slabSymbol, hostElement);
                        if (success) TotalPlaced++; else TotalSkipped++;
                    }
                    else
                    {
                        TotalSkipped++;
                        _log($"SKIPPED: No appropriate symbol for {mepType} intersection with {hostElement.GetType().Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                TotalErrors++;
                _log($"ERROR processing {mepType} element {mepElement.Id}: {ex.Message}");
            }
        }

        private List<(Element, Transform?)> GetMepElementsInSectionBox(BuiltInCategory category)
        {
            var elements = new List<(Element, Transform?)>();
            
            // Get section box for filtering
            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (_doc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBox = view3D.GetSectionBox();
                }
            }
            catch { }

            // Host model
            var hostCollector = new FilteredElementCollector(_doc)
                .OfCategory(category)
                .WhereElementIsNotElementType();
                
            if (sectionBox != null)
            {
                hostCollector = hostCollector.WherePasses(new BoundingBoxIntersectsFilter(new Outline(sectionBox.Min, sectionBox.Max)));
            }
            
            elements.AddRange(hostCollector.ToElements().Select(e => (e, (Transform?)null)));

            // Linked models (visible only)
            var linkInstances = new FilteredElementCollector(_doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>();
                
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null || linkInstance.Category == null ||
                    _doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) ||
                    linkInstance.IsHidden(_doc.ActiveView)) continue;

                var linkTransform = linkInstance.GetTotalTransform();
                var linkedCollector = new FilteredElementCollector(linkDoc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                // Transform section box to link coordinates
                if (sectionBox != null && linkTransform != null)
                {
                    try
                    {
                        var inverseTransform = linkTransform.Inverse;
                        var linkMin = inverseTransform.OfPoint(sectionBox.Min);
                        var linkMax = inverseTransform.OfPoint(sectionBox.Max);
                        var linkBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(Math.Min(linkMin.X, linkMax.X), Math.Min(linkMin.Y, linkMax.Y), Math.Min(linkMin.Z, linkMax.Z)),
                            Max = new XYZ(Math.Max(linkMin.X, linkMax.X), Math.Max(linkMin.Y, linkMax.Y), Math.Max(linkMin.Z, linkMax.Z))
                        };
                        linkedCollector = linkedCollector.WherePasses(new BoundingBoxIntersectsFilter(new Outline(linkBox.Min, linkBox.Max)));
                    }
                    catch { }
                }

                elements.AddRange(linkedCollector.ToElements().Select(e => (e, (Transform?)linkTransform)));
            }

            return elements;
        }

        private FamilySymbol? GetFamilySymbol(string familyName)
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
