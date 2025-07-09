using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing; // For Pipe
using Autodesk.Revit.DB.Mechanical; // For Duct
using Autodesk.Revit.DB.Electrical; // For CableTray
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
[Regeneration(RegenerationOption.Manual)]
public class TestFittingSleevePlacementCommand : IExternalCommand
{
    // Helper to transform a bounding box by a transform (for linked elements)
    private static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
    {
        var corners = new List<XYZ>
        {
            bbox.Min,
            new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
            new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
            new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
            new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
            new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
            new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
            bbox.Max
        };
        var transformed = corners.Select(pt => transform.OfPoint(pt)).ToList();
        var min = new XYZ(transformed.Min(p => p.X), transformed.Min(p => p.Y), transformed.Min(p => p.Z));
        var max = new XYZ(transformed.Max(p => p.X), transformed.Max(p => p.Y), transformed.Max(p => p.Z));
        return new BoundingBoxXYZ { Min = min, Max = max };
    }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            // Use the active 3D view for filtering
            View activeView = doc.ActiveView;
            if (!(activeView is View3D) || activeView.IsTemplate)
            {
                TaskDialog.Show("Error", "Please run this command from a non-template 3D view.");
                return Result.Failed;
            }

            // Initialize log file for this test command
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.InitLogFileOverwrite("TestFittingSleevePlacement_Debug");
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info("Starting TestFittingSleevePlacementCommand execution.");

            // Collect all relevant MEP elements (pipes, ducts, cable trays, dampers, duct fittings, pipe fittings) visible in the active view
            var pipes = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(Pipe)).WhereElementIsNotElementType().ToList();
            var ducts = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(Duct)).WhereElementIsNotElementType().ToList();
            var cableTrays = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(CableTray)).WhereElementIsNotElementType().ToList();
            var dampers = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .WhereElementIsNotElementType().ToList();
            var ductFittings = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .WhereElementIsNotElementType().ToList();
            var pipeFittings = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_PipeFitting)
                .WhereElementIsNotElementType().ToList();

            var allMEPElements = new List<Element>();
            allMEPElements.AddRange(pipes);
            allMEPElements.AddRange(ducts);
            allMEPElements.AddRange(cableTrays);
            allMEPElements.AddRange(dampers);
            allMEPElements.AddRange(ductFittings);
            allMEPElements.AddRange(pipeFittings);

            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"Collected {allMEPElements.Count} MEP elements visible in view '{activeView.Name}'.");


            // Use ViewElementSelectionHelper to collect structural and architectural elements from host and all visible links in the view
            var intersectionElements = new List<(Element element, Transform? linkTransform, string source)>();

            // Add all relevant categories for structure and architecture
            var categories = new List<BuiltInCategory> {
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_Windows,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_ShaftOpening
            };
            var view3D = activeView as View3D;

            // Host elements
            var hostElements = Helpers.ViewElementSelectionHelper.GetElementsInCropOrSectionBox(doc, view3D, categories);
            foreach (var e in hostElements)
            {
                if (categories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                {
                    intersectionElements.Add((e, null, "Host"));
                }
            }

            // Linked elements (with transform)
            var linkedElements = Helpers.ViewElementSelectionHelper.GetElementsFromLinkedModelsWithTransform(doc, view3D, categories);
            foreach (var tuple in linkedElements)
            {
                var e = tuple.element;
                var linkTransform = tuple.transform;
                var linkName = tuple.linkName;
                if (categories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                {
                    intersectionElements.Add((e, linkTransform, $"Link:{linkName}"));
                }
            }

            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"Collected {intersectionElements.Count} structural/architectural elements (host + links) visible in view '{activeView.Name}'.");

            // OOP helpers for sleeve placement, suppression, and logging
            // Correct instantiation for helpers/services
            var pipeSleevePlacer = new Services.PipeSleevePlacer(doc); // Pass doc if required by constructor
            var rectSleevePlacer = new Services.RectangularSleevePlacer();
            // OpeningDuplicationChecker and SleeveLogManager are static classes; use static methods directly

            // --- Collect MEP elements from all linked models (NO view-based filtering) ---
            var linkedMEPElements = new List<(Element element, Document linkDoc, string linkName, Transform linkTransform)>();
            var linkInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null) continue;
                // Collect ALL MEP elements in the link (no view/crop filtering)
                var linkPipes = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Pipe)).WhereElementIsNotElementType().ToList();
                var linkDucts = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Duct)).WhereElementIsNotElementType().ToList();
                var linkCableTrays = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(CableTray)).WhereElementIsNotElementType().ToList();
                var linkDampers = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType().ToList();
                var linkDuctFittings = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .WhereElementIsNotElementType().ToList();
                var linkPipeFittings = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .WhereElementIsNotElementType().ToList();
                var allLinkMEP = new List<Element>();
                allLinkMEP.AddRange(linkPipes);
                allLinkMEP.AddRange(linkDucts);
                allLinkMEP.AddRange(linkCableTrays);
                allLinkMEP.AddRange(linkDampers);
                allLinkMEP.AddRange(linkDuctFittings);
                allLinkMEP.AddRange(linkPipeFittings);
                foreach (var e in allLinkMEP)
                {
                    linkedMEPElements.Add((e, linkDoc, linkInstance.Name, linkInstance.GetTransform()));
                }
            }
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"Collected {linkedMEPElements.Count} MEP elements from all linked models (no view filter).");

            // --- Test intersection detection for linked MEPs vs host/linked structure/archi elements ---
            // Use IntersectionDetectionService to find intersections
            var intersectionService2 = new Services.IntersectionDetectionService();
            var intersectionResults = new List<Services.IntersectionResult>();
            foreach (var (element, linkDoc, linkName, linkTransform) in linkedMEPElements)
            {
                // Build targetTransforms for intersectionElements
                var targetTransforms = intersectionElements.ToDictionary(
                    t => t.element,
                    t => t.linkTransform ?? Transform.Identity);
                var results = intersectionService2.FindIntersections(
                    uiDoc,
                    linkDoc,
                    new List<Element> { element },
                    new Dictionary<Element, Transform> { { element, linkTransform ?? Transform.Identity } },
                    doc,
                    intersectionElements.Select(t => t.element),
                    targetTransforms,
                    view3D);
                intersectionResults.AddRange(results);
            }
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"IntersectionDetectionService found {intersectionResults.Count} intersections between linked MEP elements and host/linked structure/archi elements.");

            // Remove reference to undefined variable structuralElements
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"Collected {intersectionElements.Count} structural elements (host + links) visible in view '{activeView.Name}'.");

            // --- Robust intersection detection for host MEPs vs all structure/archi (host + links) ---
            var intersectionServiceHost = new Services.IntersectionDetectionService();

            // Build targetTransforms for intersectionElements
            var hostTargetTransforms = intersectionElements.ToDictionary(
                t => t.element,
                t => t.linkTransform ?? Transform.Identity);
            var hostIntersectionResults = intersectionServiceHost.FindIntersections(
                uiDoc,
                doc,
                allMEPElements,
                null, // sourceTransforms is null or can be set to Identity for host elements
                doc, // targetDoc is host, but intersectionElements may include linked elements
                intersectionElements.Select(t => t.element),
                hostTargetTransforms,
                view3D);

            int intersectionCount = 0;
            int elementsWithIntersection = 0;
            var processedMEP = new HashSet<Element>();

            foreach (var result in hostIntersectionResults)
            {
                intersectionCount++;
                processedMEP.Add(result.SourceElement);
                var mepElement = result.SourceElement;
                var structuralElement = result.TargetElement;

                // --- Sleeve placement logic using SleeveClearance ---
                JSE_RevitAddin_MEP_OPENINGS.Services.SleeveClearance clearance = null;
                bool isInsulated = false; // TODO: Detect insulation
                bool isDamper = mepElement.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory;
                if (isDamper)
                {
                    // TODO: Detect connector side dynamically (e.g., from connector geometry)
                    string connectorSide = "Right"; // Placeholder for now
                    clearance = JSE_RevitAddin_MEP_OPENINGS.Services.SleeveClearance.Damper(connectorSide);
                }
                else if (isInsulated)
                {
                    clearance = JSE_RevitAddin_MEP_OPENINGS.Services.SleeveClearance.Insulated();
                }
                else
                {
                    clearance = JSE_RevitAddin_MEP_OPENINGS.Services.SleeveClearance.Standard();
                }

                // Use rectangular placer for ducts, trays, dampers
                if (mepElement is Duct || mepElement is CableTray || isDamper)
                {
                    rectSleevePlacer.PlaceSleeve(doc, mepElement, structuralElement, clearance);
                }
                // Use pipe placer for pipes
                else if (mepElement is Pipe)
                {
                    // Example: pipeSleevePlacer.PlaceSleeve(doc, mepElement, structuralElement, ...);
                }
                // Log and suppress duplicates
                // OpeningDuplicationChecker.SuppressDuplicates(...); // Call static method if available
                // SleeveLogManager.LogSleevePlacement(...); // Call static method if available
            }
            elementsWithIntersection = processedMEP.Count;

            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"Processed {allMEPElements.Count} MEP elements. {elementsWithIntersection} had at least one intersection. Total intersections: {intersectionCount}.");
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"SUMMARY: Processed {allMEPElements.Count} MEP elements (pipes, ducts, trays, dampers, duct/pipe fittings) visible in view '{activeView.Name}'. Found {intersectionCount} possible intersections (robust intersection logic). {elementsWithIntersection} elements had at least one intersection.");
            return Result.Succeeded;
        }
    }
}