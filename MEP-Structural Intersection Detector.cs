using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class FindMepStructuralIntersections : IExternalCommand
    {
        private static void Log(string message)
        {
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "JSE_MEP_Openings",
                "intersections.log");

            if (!Directory.Exists(Path.GetDirectoryName(logPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath));
            }

            string logEntry = $"[{DateTime.Now:dd-MM-yyyy HH:mm:ss}] {message}";
            File.AppendAllText(logPath, logEntry + Environment.NewLine);
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Check for 3D view
                var view3D = doc.ActiveView as View3D;
                if (view3D == null)
                {
                    TaskDialog.Show("Error", "Please open a 3D view with a section box and try again.");
                    return Result.Failed;
                }

                Log($"\n=== INTERSECTION DETECTION STARTED ===");
                Log($"Document: {doc.Title}");
                Log($"View: {view3D.Name}");

                // STEP 1: Get section box in model coordinates
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                Transform sectionTransform = sectionBox.Transform;

                List<XYZ> corners = new List<XYZ>
                {
                    sectionTransform.OfPoint(sectionBox.Min),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(sectionBox.Max)
                };

                XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));

                Log($"Section box: Min=({modelMin.X:F1}, {modelMin.Y:F1}, {modelMin.Z:F1}) Max=({modelMax.X:F1}, {modelMax.Y:F1}, {modelMax.Z:F1})");

                // STEP 2: Collect MEP and Structural elements
                var mepElements = new List<Element>();
                var wallElements = new List<Element>();

                CollectElements(doc, modelMin, modelMax, ref mepElements, ref wallElements);

                Log($"Found {mepElements.Count} MEP elements and {wallElements.Count} walls in section box");

                if (mepElements.Count == 0)
                {
                    TaskDialog.Show("No MEP Elements", "No MEP elements found in section box.");
                    return Result.Cancelled;
                }

                if (wallElements.Count == 0)
                {
                    TaskDialog.Show("No Walls", "No walls found in section box.");
                    return Result.Cancelled;
                }

                // STEP 3: Find intersections
            var intersections = FindIntersections(mepElements, wallElements, doc);

                Log($"\n=== INTERSECTION RESULTS ===");
                Log($"Total intersections found: {intersections.Count}");

                // STEP 4: Display results
                ShowResults(intersections);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.Message}");
                Log($"Stack: {ex.StackTrace}");
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
                return Result.Failed;
            }
        }

        private void CollectElements(Document doc, XYZ modelMin, XYZ modelMax, 
                                     ref List<Element> mepElements, ref List<Element> wallElements)
        {
            Outline hostOutline = new Outline(modelMin, modelMax);

            // MEP categories - INCLUDING DAMPERS
            BuiltInCategory[] mepCats = {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctAccessory,  // DAMPERS ARE HERE
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting
            };

            Log($"DEBUG: Section box in active document: Min=({modelMin.X:F2}, {modelMin.Y:F2}, {modelMin.Z:F2}) Max=({modelMax.X:F2}, {modelMax.Y:F2}, {modelMax.Z:F2})");

            // Collect from host
            foreach (var cat in mepCats)
            {
                mepElements.AddRange(
                    new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements()
                );
            }

            wallElements.AddRange(
                new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                    .ToElements()
            );

            // Collect from links
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            Log($"DEBUG: Found {links.Count} linked models");

            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                Transform linkTransform = link.GetTotalTransform();
                Transform invTransform = linkTransform.Inverse;

                XYZ linkMin = invTransform.OfPoint(modelMin);
                XYZ linkMax = invTransform.OfPoint(modelMax);

                XYZ actualMin = new XYZ(
                    Math.Min(linkMin.X, linkMax.X),
                    Math.Min(linkMin.Y, linkMax.Y),
                    Math.Min(linkMin.Z, linkMax.Z)
                );
                XYZ actualMax = new XYZ(
                    Math.Max(linkMin.X, linkMax.X),
                    Math.Max(linkMin.Y, linkMax.Y),
                    Math.Max(linkMin.Z, linkMax.Z)
                );

                Log($"DEBUG: Section box in linked file '{linkDoc.Title}': Min=({actualMin.X:F2}, {actualMin.Y:F2}, {actualMin.Z:F2}) Max=({actualMax.X:F2}, {actualMax.Y:F2}, {actualMax.Z:F2})");

                Outline linkOutline = new Outline(actualMin, actualMax);

                foreach (var cat in mepCats)
                {
                    mepElements.AddRange(
                        new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements()
                    );
                }

                wallElements.AddRange(
                    new FilteredElementCollector(linkDoc)
                        .OfCategory(BuiltInCategory.OST_Walls)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                        .ToElements()
                );
            }
        }

        private List<IntersectionResult> FindIntersections(List<Element> mepElements, List<Element> wallElements, Document doc)
        {
            var intersections = new List<IntersectionResult>();

            Log($"Using MepIntersectionService for clash detection with {mepElements.Count} MEP elements...");

            // Get all linked documents for transform lookup
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var mep in mepElements)
            {
                try
                {
                    bool isDamper = false;
                    string familyName = string.Empty;
                    if (mep is FamilyInstance fi)
                    {
                        familyName = fi.Symbol?.Family?.Name ?? string.Empty;
                        isDamper = familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (isDamper)
                        {
                            Log($"DEBUG: Processing damper element {mep.Id} - Family: {familyName}");
                        }
                    }

                    var mepLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == mep.Document.Title);
                    Transform? mepTransform = mepLink?.GetTotalTransform();
                    if (mepLink != null)
                    {
                        Log($"DEBUG: Found MEP link: {mepLink.Name} -> {mep.Document.Title}");
                        if (mepTransform != null)
                        {
                            Log($"DEBUG: MEP link transform origin: ({mepTransform.Origin.X:F2}, {mepTransform.Origin.Y:F2}, {mepTransform.Origin.Z:F2})");
                        }
                    }

                    var structuralElements = new List<(Element, Transform?)>();
                    foreach (var wall in wallElements)
                    {
                        Log($"DEBUG: Wall {wall.Id} is from document: {wall.Document.Title}");
                        var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                        if (wallLink != null)
                        {
                            var wallTransform = wallLink.GetTotalTransform();
                            Log($"DEBUG: Found wall link: {wallLink.Name} -> {wall.Document.Title}");
                            Log($"DEBUG: Wall {wall.Id} transform: Origin=({wallTransform.Origin.X:F2}, {wallTransform.Origin.Y:F2}, {wallTransform.Origin.Z:F2})");
                            structuralElements.Add((wall, wallTransform));
                        }
                        else
                        {
                            Log($"DEBUG: No link found for wall {wall.Id} from document {wall.Document.Title}");
                            Log($"DEBUG: Available links: {string.Join(", ", links.Select(l => l.GetLinkDocument()?.Title ?? "NULL"))}");
                            structuralElements.Add((wall, null));
                        }
                    }

                    var mepBBox = mep.get_BoundingBox(null);

                    if (isDamper)
                    {
                        var damperHits = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.FindDamperIntersections(
                            mep,
                            structuralElements,
                            mepTransform,
                            Log);

                        Log($"DEBUG: Damper {mep.Id} intersections found: {damperHits.Count}");

                        foreach (var (structuralElement, bbox, center) in damperHits)
                        {
                            intersections.Add(new IntersectionResult
                            {
                                MepElement = mep,
                                StructuralElement = structuralElement,
                                MepCategory = mep.Category?.Name ?? "Unknown",
                                StructuralCategory = structuralElement.Category?.Name ?? "Unknown",
                                MepId = mep.Id,
                                WallId = structuralElement.Id,
                                IntersectionCenter = center
                            });

                            Log($"Intersection found: Damper {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }

                        if (damperHits.Count == 0)
                        {
                            Log($"DEBUG: No intersections found for damper {mep.Id}");
                        }

                        continue;
                    }

                    var locationCurve = mep.Location as LocationCurve;
                    if (locationCurve?.Curve is Line mepLine)
                    {
                        Log($"DEBUG: Processing MEP {mep.Id} ({mep.Category?.Name}) with line from {mepLine.GetEndPoint(0)} to {mepLine.GetEndPoint(1)}");

                        Line mepLineInHostShared;
                        if (mepTransform != null)
                        {
                            mepLineInHostShared = Line.CreateBound(
                                mepTransform.OfPoint(mepLine.GetEndPoint(0)),
                                mepTransform.OfPoint(mepLine.GetEndPoint(1))
                            );
                        }
                        else
                        {
                            mepLineInHostShared = mepLine;
                        }

                        Log($"DEBUG: MEP line in host shared coordinates: {mepLineInHostShared.GetEndPoint(0)} to {mepLineInHostShared.GetEndPoint(1)}");

                        BoundingBoxXYZ? mepBBoxInHostShared = mepBBox;
                        if (mepBBox != null && mepTransform != null)
                        {
                            Log($"DEBUG: MEP bbox before transform: Min=({mepBBox.Min.X:F2}, {mepBBox.Min.Y:F2}, {mepBBox.Min.Z:F2}) Max=({mepBBox.Max.X:F2}, {mepBBox.Max.Y:F2}, {mepBBox.Max.Z:F2})");

                            var mepPts = new[]
                            {
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Min.Z))
                            };

                            var mepNewMin = new XYZ(mepPts.Min(p => p.X), mepPts.Min(p => p.Y), mepPts.Min(p => p.Z));
                            var mepNewMax = new XYZ(mepPts.Max(p => p.X), mepPts.Max(p => p.Y), mepPts.Max(p => p.Z));

                            mepBBoxInHostShared = new BoundingBoxXYZ
                            {
                                Min = mepNewMin,
                                Max = mepNewMax
                            };

                            Log($"DEBUG: MEP bbox after transform: Min=({mepBBoxInHostShared.Min.X:F2}, {mepBBoxInHostShared.Min.Y:F2}, {mepBBoxInHostShared.Min.Z:F2}) Max=({mepBBoxInHostShared.Max.X:F2}, {mepBBoxInHostShared.Max.Y:F2}, {mepBBoxInHostShared.Max.Z:F2})");
                        }

                        var hits = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.FindIntersections(
                            mepLineInHostShared,
                            mepBBoxInHostShared,
                            structuralElements,
                            Log);

                        Log($"DEBUG: MepIntersectionService found {hits.Count} intersections for MEP {mep.Id}");

                        if (hits.Count > 0)
                        {
                            var (structuralElement, bbox, center) = hits[0];
                            var intersection = new IntersectionResult
                            {
                                MepElement = mep,
                                StructuralElement = structuralElement,
                                MepCategory = mep.Category?.Name ?? "Unknown",
                                StructuralCategory = structuralElement.Category?.Name ?? "Unknown",
                                MepId = mep.Id,
                                WallId = structuralElement.Id,
                                IntersectionCenter = center
                            };

                            intersections.Add(intersection);
                            Log($"Intersection found: MEP {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }
                    }
                    else
                    {
                        Log($"DEBUG: MEP element {mep.Id} has no valid location curve or is not a line");
                    }
                }
                catch (Exception ex)
                {
                    Log($"ERROR: Failed to process MEP element {mep.Id}: {ex.Message}");
                }
            }

            return intersections;
        }


        private XYZ GetLineIntersectionPoint(Line line1, Line line2)
        {
            // Simple intersection point calculation
            // For now, return midpoint of line1 as approximation
            return (line1.GetEndPoint(0) + line1.GetEndPoint(1)) / 2;
        }

        private bool BoundingBoxesIntersect(BoundingBoxXYZ box1, BoundingBoxXYZ box2)
        {
            return !(box1.Max.X < box2.Min.X || box1.Min.X > box2.Max.X ||
                     box1.Max.Y < box2.Min.Y || box1.Min.Y > box2.Max.Y ||
                     box1.Max.Z < box2.Min.Z || box1.Min.Z > box2.Max.Z);
        }

        private bool GeometriesIntersect(Element mep, Element wall)
        {
            try
            {
                // Get solid geometry
                Options geomOptions = new Options();
                geomOptions.DetailLevel = ViewDetailLevel.Fine;
                geomOptions.IncludeNonVisibleObjects = false;

                GeometryElement mepGeom = mep.get_Geometry(geomOptions);
                GeometryElement wallGeom = wall.get_Geometry(geomOptions);

                if (mepGeom == null || wallGeom == null)
                    return false;

                // Extract solids
                var mepSolids = ExtractSolids(mepGeom);
                var wallSolids = ExtractSolids(wallGeom);

                if (mepSolids.Count == 0 || wallSolids.Count == 0)
                    return false;

                // Check for intersection
                foreach (var mepSolid in mepSolids)
                {
                    foreach (var wallSolid in wallSolids)
                    {
                        try
                        {
                            Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                                mepSolid, wallSolid, BooleanOperationsType.Intersect);

                            if (intersection != null && intersection.Volume > 1e-6) // Tolerance
                            {
                                return true;
                            }
                        }
                        catch
                        {
                            // Boolean operation failed, skip
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Log($"Geometry check error for MEP {mep.Id} and Wall {wall.Id}: {ex.Message}");
                return false;
            }
        }

        private List<Solid> ExtractSolids(GeometryElement geomElement)
        {
            var solids = new List<Solid>();

            foreach (GeometryObject geomObj in geomElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    solids.Add(solid);
                }
                else if (geomObj is GeometryInstance geomInst)
                {
                    GeometryElement instGeom = geomInst.GetInstanceGeometry();
                    solids.AddRange(ExtractSolids(instGeom));
                }
            }

            return solids;
        }

        private void ShowResults(List<IntersectionResult> intersections)
        {
            if (intersections.Count == 0)
            {
                TaskDialog.Show("No Intersections", "No intersections found between MEP elements and walls.");
                return;
            }

            // Group by MEP category
            var grouped = intersections.GroupBy(i => i.MepCategory);

            string summary = $"🎯 INTERSECTIONS FOUND: {intersections.Count}\n\n";
            
            foreach (var group in grouped.OrderBy(g => g.Key))
            {
                summary += $"📍 {group.Key}: {group.Count()} intersections\n";
            }

            summary += $"\n✅ Check log file for detailed results.";

            // Detailed log
            Log("\n=== DETAILED INTERSECTION LIST ===");
            foreach (var intersection in intersections)
            {
                Log($"MEP: {intersection.MepCategory} (ID: {intersection.MepId}) <-> " +
                    $"Wall (ID: {intersection.WallId})");
            }

            TaskDialog.Show("Intersection Detection Complete", summary);
        }

        private class IntersectionResult
        {
            public Element MepElement { get; set; }
            public Element StructuralElement { get; set; }
            public string MepCategory { get; set; }
            public string StructuralCategory { get; set; }
            public ElementId MepId { get; set; }
            public ElementId WallId { get; set; }
            public XYZ IntersectionCenter { get; set; }
        }
    }
}