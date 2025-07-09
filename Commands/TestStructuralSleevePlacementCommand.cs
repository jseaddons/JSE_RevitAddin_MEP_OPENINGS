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
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class TestStructuralSleevePlacementCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            View activeView = doc.ActiveView;
            if (!(activeView is View3D) || activeView.IsTemplate)
            {
                TaskDialog.Show("Error", "Please run this command from a non-template 3D view.");
                return Result.Failed;
            }

            DebugLogger.InitLogFileOverwrite("TestStructuralSleevePlacement_Debug");
            DebugLogger.Info("Starting TestStructuralSleevePlacementCommand execution.");

            // STEP 1: Collect all relevant MEP elements (pipes, ducts, cable trays) visible in the active view
            var pipes = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(Pipe)).WhereElementIsNotElementType().ToList();
            var ducts = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(Duct)).WhereElementIsNotElementType().ToList();
            var cableTrays = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(CableTray)).WhereElementIsNotElementType().ToList();


            // Also collect damper FamilyInstances (category OST_DuctAccessory, family name contains "Damper" but not "DamperOpeningOnWall")
            var dampers = new FilteredElementCollector(doc, activeView.Id)
                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol != null && fi.Symbol.Family != null &&
                    fi.Symbol.Family.Name.Contains("Damper") && !fi.Symbol.Family.Name.Contains("DamperOpeningOnWall"))
                .ToList();

            var allMEPElements = new List<Element>();
            allMEPElements.AddRange(pipes);
            allMEPElements.AddRange(ducts);
            allMEPElements.AddRange(cableTrays);
            allMEPElements.AddRange(dampers);

            DebugLogger.Info($"Collected {allMEPElements.Count} MEP elements visible in view '{activeView.Name}'.");

            // STEP 2: Do NOT collect any structural/arch elements from the active document
            DebugLogger.Info($"Skipping host structural/arch elements. Only using linked models for intersection.");
            var allStructuralElements = new List<Element>();

            // STEP 3: Collect and log elements from visible linked models (structural/architectural)
            var linkInstances = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => !link.IsHidden(activeView))
                .ToList();

            int totalLinkedStruct = 0;
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null) continue;

                // Only structural floors, structural framing, and walls from links
                var linkedFloors = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Floor)).WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                    .Where(e => e.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                    .ToList();
                var linkedFraming = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    .ToList();
                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Wall)).WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls)
                    .ToList();
                int count = linkedFloors.Count + linkedFraming.Count + linkedWalls.Count;
                totalLinkedStruct += count;
                DebugLogger.Info($"Link '{linkInstance.Name}': {linkedFloors.Count} structural floors, {linkedFraming.Count} framing, {linkedWalls.Count} walls (total {count})");
            }
            DebugLogger.Info($"Total structural elements in visible links: {totalLinkedStruct}");

            // --- INTEGRATION WITH IntersectionDetectionService ---
            // Collect all linked structural elements and their transforms
            var structElements = new List<Element>();
            var structTransforms = new Dictionary<Element, Transform>();
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null) continue;
                var linkedFloors = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Floor)).WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                    .Where(e => e.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                    .ToList();
                var linkedFraming = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    .ToList();
                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Wall)).WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls)
                    .ToList();
                foreach (Element s in linkedFloors)
                {
                    structElements.Add(s);
                    structTransforms[s] = linkInstance.GetTransform();
                }
                foreach (Element s in linkedFraming)
                {
                    structElements.Add(s);
                    structTransforms[s] = linkInstance.GetTransform();
                }
                foreach (Element s in linkedWalls)
                {
                    structElements.Add(s);
                    structTransforms[s] = linkInstance.GetTransform();
                }
            }

            // All MEP elements are in the host doc, so identity transform
            var mepTransforms = new Dictionary<Element, Transform>();
            foreach (var mep in allMEPElements)
                mepTransforms[mep] = Transform.Identity;

            // Use IntersectionDetectionService
            var intersectionService = new JSE_RevitAddin_MEP_OPENINGS.Services.IntersectionDetectionService();
            // Log bounding boxes for all MEP elements
            foreach (var mep in allMEPElements)
            {
                var bbox = mep.get_BoundingBox(activeView as View3D);
                if (bbox != null)
                {
                    DebugLogger.Info($"MEP Element {mep.Id} ({mep.GetType().Name}) bbox: Min=({bbox.Min.X:F2},{bbox.Min.Y:F2},{bbox.Min.Z:F2}) Max=({bbox.Max.X:F2},{bbox.Max.Y:F2},{bbox.Max.Z:F2})");
                }
                else
                {
                    DebugLogger.Info($"MEP Element {mep.Id} ({mep.GetType().Name}) has no bounding box in view.");
                }
            }

            // Log bounding boxes for all structural elements
            foreach (var s in structElements)
            {
                var bbox = s.get_BoundingBox(activeView as View3D);
                if (bbox != null)
                {
                    DebugLogger.Info($"Struct Element {s.Id} ({s.GetType().Name}) bbox: Min=({bbox.Min.X:F2},{bbox.Min.Y:F2},{bbox.Min.Z:F2}) Max=({bbox.Max.X:F2},{bbox.Max.Y:F2},{bbox.Max.Z:F2})");
                }
                else
                {
                    DebugLogger.Info($"Struct Element {s.Id} ({s.GetType().Name}) has no bounding box in view.");
                }
            }

            // Log positions (centers) of all MEP elements
            foreach (var mep in allMEPElements)
            {
                var bbox = mep.get_BoundingBox(activeView as View3D);
                if (bbox != null)
                {
                    var center = (bbox.Min + bbox.Max) / 2.0;
                    DebugLogger.Info($"MEP Element {mep.Id} center: ({center.X:F2}, {center.Y:F2}, {center.Z:F2})");
                }
            }

            // Log positions (centers) of all structural elements
            foreach (var s in structElements)
            {
                var bbox = s.get_BoundingBox(activeView as View3D);
                if (bbox != null)
                {
                    var center = (bbox.Min + bbox.Max) / 2.0;
                    DebugLogger.Info($"Struct Element {s.Id} center: ({center.X:F2}, {center.Y:F2}, {center.Z:F2})");
                }
            }

            // --- BEGIN: Detailed pairwise intersection logging ---
            DebugLogger.Info("--- BEGIN: Pairwise MEP-Structural Intersection Checks (BoundingBox & Solid) ---");
            var intersectionServiceForLog = new JSE_RevitAddin_MEP_OPENINGS.Services.IntersectionDetectionService();
            foreach (var mep in allMEPElements)
            {
                var mepBbox = mep.get_BoundingBox(activeView as View3D);
                var mepXf = mepTransforms.ContainsKey(mep) ? mepTransforms[mep] : Autodesk.Revit.DB.Transform.Identity;
                var mepBboxTrans = intersectionServiceForLog.TransformBoundingBox(mepBbox, mepXf);
                var mepSolid = intersectionServiceForLog.GetElementSolid(mep);
                var mepSolidTrans = mepSolid != null ? Autodesk.Revit.DB.SolidUtils.CreateTransformed(mepSolid, mepXf) : null;
                foreach (var s in structElements)
                {
                    var sBbox = s.get_BoundingBox(activeView as View3D);
                    var sXf = structTransforms.ContainsKey(s) ? structTransforms[s] : Autodesk.Revit.DB.Transform.Identity;
                    var sBboxTrans = intersectionServiceForLog.TransformBoundingBox(sBbox, sXf);
                    var sSolid = intersectionServiceForLog.GetElementSolid(s);
                    var sSolidTrans = sSolid != null ? Autodesk.Revit.DB.SolidUtils.CreateTransformed(sSolid, sXf) : null;
                    bool bboxOverlap = false;
                    bool solidOverlap = false;
                    double solidIntersectionVolume = 0.0;
                    if (mepBboxTrans != null && sBboxTrans != null)
                        bboxOverlap = intersectionServiceForLog.BoundingBoxesIntersect(mepBboxTrans, sBboxTrans);
                    if (bboxOverlap && mepSolidTrans != null && sSolidTrans != null)
                    {
                        try
                        {
                            var intersection = Autodesk.Revit.DB.BooleanOperationsUtils.ExecuteBooleanOperation(mepSolidTrans, sSolidTrans, Autodesk.Revit.DB.BooleanOperationsType.Intersect);
                            if (intersection != null && intersection.Volume > 1e-6)
                            {
                                solidOverlap = true;
                                solidIntersectionVolume = intersection.Volume;
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Info($"SOLID INTERSECTION ERROR: MEP {mep.Id} <-> Struct {s.Id}: {ex.Message}");
                        }
                    }
                    DebugLogger.Info($"PAIR: MEP {mep.Id} ({mep.GetType().Name}) <-> Struct {s.Id} ({s.GetType().Name}) | BBoxOverlap: {bboxOverlap} | SolidOverlap: {solidOverlap} | SolidIntersectionVolume: {solidIntersectionVolume:F6}");
                }
            }
            DebugLogger.Info("--- END: Pairwise MEP-Structural Intersection Checks ---");

            // Log distances between every MEP and every struct element center

            // Use robust ReferenceIntersector-based intersection detection (OOP, reliable)

            // Use bounding box only intersection logic to closely mimic FireDamperPlaceCommand
            var bboxIntersections = intersectionService.FindIntersectionsBoundingBoxOnly(
                allMEPElements,
                structElements,
                mepTransforms,
                structTransforms,
                activeView as View3D
            );

            // Log all damper intersections for debugging
            foreach (var result in bboxIntersections)
            {
                string mepType = result.SourceElement.GetType().Name;
                var cat = result.SourceElement.Category;
                if (mepType == "FamilyInstance" && cat != null && cat.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                {
                    var famInst = result.SourceElement as FamilyInstance;
                    if (famInst != null && famInst.Symbol != null && famInst.Symbol.Family != null &&
                        famInst.Symbol.Family.Name.Contains("Damper") && !famInst.Symbol.Family.Name.Contains("DamperOpeningOnWall"))
                    {
                        DebugLogger.Info($"DAMPER INTERSECTION: MEP {result.SourceElement.Id} (Family: {famInst.Symbol.Family.Name}, Cat: {cat.Name}) <-> Struct {result.TargetElement.Id} ({result.TargetElement.GetType().Name})");
                    }
                }
            }

            // For summary, use the bounding box intersection results
            var intersections = bboxIntersections;

            // --- BEGIN: HIT SUMMARY LOGGING ---

            // Enhanced summary: track hits and unique MEPs for each struct type
            int pipeWall = 0, pipeStructFloor = 0, pipeStructFraming = 0;
            int ductWall = 0, ductStructFloor = 0, ductStructFraming = 0;
            int trayWall = 0, trayStructFloor = 0, trayStructFraming = 0;
            int damperWall = 0, damperStructFloor = 0, damperStructFraming = 0;
            int ductFittingWall = 0, ductFittingStructFloor = 0, ductFittingStructFraming = 0;
            int pipeFittingWall = 0, pipeFittingStructFloor = 0, pipeFittingStructFraming = 0;
            var uniquePipeWall = new HashSet<ElementId>();
            var uniquePipeStructFloor = new HashSet<ElementId>();
            var uniquePipeStructFraming = new HashSet<ElementId>();
            var uniqueDuctWall = new HashSet<ElementId>();
            var uniqueDuctStructFloor = new HashSet<ElementId>();
            var uniqueDuctStructFraming = new HashSet<ElementId>();
            var uniqueTrayWall = new HashSet<ElementId>();
            var uniqueTrayStructFloor = new HashSet<ElementId>();
            var uniqueTrayStructFraming = new HashSet<ElementId>();
            var uniqueDamperWall = new HashSet<ElementId>();
            var uniqueDamperStructFloor = new HashSet<ElementId>();
            var uniqueDamperStructFraming = new HashSet<ElementId>();
            var uniqueDuctFittingWall = new HashSet<ElementId>();
            var uniqueDuctFittingStructFloor = new HashSet<ElementId>();
            var uniqueDuctFittingStructFraming = new HashSet<ElementId>();
            var uniquePipeFittingWall = new HashSet<ElementId>();
            var uniquePipeFittingStructFloor = new HashSet<ElementId>();
            var uniquePipeFittingStructFraming = new HashSet<ElementId>();

            foreach (var result in intersections)
            {
                string mepType = result.SourceElement.GetType().Name;
                string structType = result.TargetElement.GetType().Name;
                // Only structural floors (not archi), so treat Floor as struct
                if (mepType == "Pipe")
                {
                    if (structType == "Wall") { pipeWall++; uniquePipeWall.Add(result.SourceElement.Id); }
                    else if (structType == "Floor") { pipeStructFloor++; uniquePipeStructFloor.Add(result.SourceElement.Id); }
                    else if (structType == "FamilyInstance") { pipeStructFraming++; uniquePipeStructFraming.Add(result.SourceElement.Id); }
                }
                else if (mepType == "Duct")
                {
                    if (structType == "Wall") { ductWall++; uniqueDuctWall.Add(result.SourceElement.Id); }
                    else if (structType == "Floor") { ductStructFloor++; uniqueDuctStructFloor.Add(result.SourceElement.Id); }
                    else if (structType == "FamilyInstance") { ductStructFraming++; uniqueDuctStructFraming.Add(result.SourceElement.Id); }
                }
                else if (mepType == "CableTray")
                {
                    if (structType == "Wall") { trayWall++; uniqueTrayWall.Add(result.SourceElement.Id); }
                    else if (structType == "Floor") { trayStructFloor++; uniqueTrayStructFloor.Add(result.SourceElement.Id); }
                    else if (structType == "FamilyInstance") { trayStructFraming++; uniqueTrayStructFraming.Add(result.SourceElement.Id); }
                }
                else if (mepType == "MechanicalFitting") // Duct Fitting
                {
                    if (structType == "Wall") { ductFittingWall++; uniqueDuctFittingWall.Add(result.SourceElement.Id); }
                    else if (structType == "Floor") { ductFittingStructFloor++; uniqueDuctFittingStructFloor.Add(result.SourceElement.Id); }
                    else if (structType == "FamilyInstance") { ductFittingStructFraming++; uniqueDuctFittingStructFraming.Add(result.SourceElement.Id); }
                }
                else if (mepType == "PipeFitting")
                {
                    if (structType == "Wall") { pipeFittingWall++; uniquePipeFittingWall.Add(result.SourceElement.Id); }
                    else if (structType == "Floor") { pipeFittingStructFloor++; uniquePipeFittingStructFloor.Add(result.SourceElement.Id); }
                    else if (structType == "FamilyInstance") { pipeFittingStructFraming++; uniquePipeFittingStructFraming.Add(result.SourceElement.Id); }
                }
                else if (mepType == "FamilyInstance") // Damper (and possibly other MEP family instances)
                {
                    // Try to distinguish damper by category and family name
                    var cat = result.SourceElement.Category;
                    var famInst = result.SourceElement as FamilyInstance;
                    if (cat != null && cat.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory &&
                        famInst != null && famInst.Symbol != null && famInst.Symbol.Family != null &&
                        famInst.Symbol.Family.Name.Contains("Damper") && !famInst.Symbol.Family.Name.Contains("DamperOpeningOnWall"))
                    {
                        if (structType == "Wall") { damperWall++; uniqueDamperWall.Add(result.SourceElement.Id); }
                        else if (structType == "Floor") { damperStructFloor++; uniqueDamperStructFloor.Add(result.SourceElement.Id); }
                        else if (structType == "FamilyInstance") { damperStructFraming++; uniqueDamperStructFraming.Add(result.SourceElement.Id); }
                    }
                }
            }
            DebugLogger.Info($"SUMMARY: Pipe hits - Wall: {pipeWall} (unique: {uniquePipeWall.Count}), StructFloor: {pipeStructFloor} (unique: {uniquePipeStructFloor.Count}), StructFraming: {pipeStructFraming} (unique: {uniquePipeStructFraming.Count})");
            DebugLogger.Info($"SUMMARY: Duct hits - Wall: {ductWall} (unique: {uniqueDuctWall.Count}), StructFloor: {ductStructFloor} (unique: {uniqueDuctStructFloor.Count}), StructFraming: {ductStructFraming} (unique: {uniqueDuctStructFraming.Count})");
            DebugLogger.Info($"SUMMARY: CableTray hits - Wall: {trayWall} (unique: {uniqueTrayWall.Count}), StructFloor: {trayStructFloor} (unique: {uniqueTrayStructFloor.Count}), StructFraming: {trayStructFraming} (unique: {uniqueTrayStructFraming.Count})");
            DebugLogger.Info($"SUMMARY: Damper hits - Wall: {damperWall} (unique: {uniqueDamperWall.Count}), StructFloor: {damperStructFloor} (unique: {uniqueDamperStructFloor.Count}), StructFraming: {damperStructFraming} (unique: {uniqueDamperStructFraming.Count})");
            DebugLogger.Info($"SUMMARY: DuctFitting hits - Wall: {ductFittingWall} (unique: {uniqueDuctFittingWall.Count}), StructFloor: {ductFittingStructFloor} (unique: {uniqueDuctFittingStructFloor.Count}), StructFraming: {ductFittingStructFraming} (unique: {uniqueDuctFittingStructFraming.Count})");
            DebugLogger.Info($"SUMMARY: PipeFitting hits - Wall: {pipeFittingWall} (unique: {uniquePipeFittingWall.Count}), StructFloor: {pipeFittingStructFloor} (unique: {uniquePipeFittingStructFloor.Count}), StructFraming: {pipeFittingStructFraming} (unique: {uniquePipeFittingStructFraming.Count})");
            DebugLogger.Info($"Total MEP-structural intersections (ReferenceIntersector): {intersections.Count}");
            // --- END: HIT SUMMARY LOGGING ---

            return Result.Succeeded;
        }
    }
}
