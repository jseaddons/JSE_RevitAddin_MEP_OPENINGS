using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ElementCollectorService
    {
        private readonly Action<string> _log;

        public ElementCollectorService(Action<string> log) => _log = log;

        /// <summary>
        /// Returns all MEP + structural elements whose **true** bounding box
        /// intersects the active 3-D view section box.
        /// 1. Host MUST be shared – otherwise user is warned and we exit.
        /// 2. Each link MUST be shared – otherwise user is warned and link is SKIPPED.
        /// 3. No transforms when everyone is shared (fast path).
        /// </summary>
        public void CollectElementsInSectionBox(Document hostDoc, View3D view3D,
                                              out List<Element> mepInSectionBox,
                                              out List<Element> structuralInSectionBox)
        {
            mepInSectionBox = new List<Element>();
            structuralInSectionBox = new List<Element>();

            // 1.  Host MUST be shared
            if (!IsSharedCoordinateSystem(hostDoc.ActiveProjectLocation))
            {
                TaskDialog.Show("Coordinate System Error",
                    "Host model is not shared-coordinate based.\n" +
                    "Use Acquire Coordinates and run again.");
                return;
            }

            // 2.  Get section box and strip view transform → model coordinates
            BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
            sectionBox = RemoveViewTransform(sectionBox, view3D);
            _log($"Section box (model): Min({sectionBox.Min.X:F3}, {sectionBox.Min.Y:F3}, {sectionBox.Min.Z:F3})  " +
                 $"Max({sectionBox.Max.X:F3}, {sectionBox.Max.Y:F3}, {sectionBox.Max.Z:F3})");

            // 3.  Host-shared outline (fast path)
            Transform hostShared = hostDoc.ActiveProjectLocation.GetTotalTransform();
            BoundingBoxXYZ sectionInHostShared = TransformBoundingBox(sectionBox, hostShared);
            Outline hostOutline = new Outline(sectionInHostShared.Min, sectionInHostShared.Max);

            // 4.  Host elements
            BuiltInCategory[] mepCats = {
                BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_DuctAccessory
            };
            foreach (var cat in mepCats)
                mepInSectionBox.AddRange(
                    new FilteredElementCollector(hostDoc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements());

            BuiltInCategory[] structCats = {
                BuiltInCategory.OST_Walls, BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };
            foreach (var cat in structCats)
                structuralInSectionBox.AddRange(
                    new FilteredElementCollector(hostDoc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements());

            // 5.  Linked files – one transform only (link-shared → link-internal)
            var links = new FilteredElementCollector(hostDoc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .ToList();

            foreach (RevitLinkInstance link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                if (!IsLinkSharedCoordinate(link, linkDoc))
                {
                    TaskDialog.Show("Link Coordinate Error",
                        $"Link \"{link.Name}\" is not shared-coordinate based.\n" +
                        "Reload it with Shared positioning and try again.");
                    continue;
                }

                Transform linkSharedToInternal = linkDoc.ActiveProjectLocation
                                                        .GetTotalTransform()
                                                        .Inverse;
                BoundingBoxXYZ sectionInLink = TransformBoundingBox(sectionInHostShared, linkSharedToInternal);
                _log($"Section box in link-internal: Min({sectionInLink.Min.X:F3}, {sectionInLink.Min.Y:F3}, {sectionInLink.Min.Z:F3})  " +
                     $"Max({sectionInLink.Max.X:F3}, {sectionInLink.Max.Y:F3}, {sectionInLink.Max.Z:F3})");

                Outline linkOutline = new Outline(sectionInLink.Min, sectionInLink.Max);

                foreach (var cat in mepCats)
                    mepInSectionBox.AddRange(
                        new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements());

                foreach (var cat in structCats)
                    structuralInSectionBox.AddRange(
                        new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements());
            }

            _log($"TOTAL FOUND: {mepInSectionBox.Count} MEP elements, {structuralInSectionBox.Count} host elements");
        }

        /* -------------------- helpers --------------------------------------- */

        private static bool IsSharedCoordinateSystem(ProjectLocation loc)
            => loc.Name.Equals("Shared", StringComparison.OrdinalIgnoreCase);

        private static bool IsLinkSharedCoordinate(RevitLinkInstance link, Document linkDoc)
        {
            if (linkDoc == null) return false;
            // non-identity shared→internal transform  →  shared file
            Transform t = linkDoc.ActiveProjectLocation.GetTotalTransform().Inverse;
            const double tol = 1e-3;
            return !(Math.Abs(t.Origin.X) < tol && Math.Abs(t.Origin.Y) < tol &&
                     Math.Abs(t.Origin.Z) < tol && Math.Abs(t.BasisX.X - 1) < tol &&
                     Math.Abs(t.BasisY.Y - 1) < tol && Math.Abs(t.BasisZ.Z - 1) < tol);
        }

        private static BoundingBoxXYZ RemoveViewTransform(BoundingBoxXYZ viewBox, View3D view)
        {
            Transform inv = view.CropBox.Transform.Inverse;
            XYZ min = inv.OfPoint(viewBox.Min);
            XYZ max = inv.OfPoint(viewBox.Max);
            return new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
            };
        }

        private static BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bb, Transform t)
        {
            XYZ min = t.OfPoint(bb.Min);
            XYZ max = t.OfPoint(bb.Max);
            return new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
            };
        }
    }
}
