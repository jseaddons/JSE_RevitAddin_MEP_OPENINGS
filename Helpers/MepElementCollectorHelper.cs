using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class MepElementCollectorHelper
    {
        /// <summary>
        /// Collects FamilyInstance elements (host + visible links) of a given family name substring.
        /// </summary>
        public static List<(FamilyInstance instance, Transform? transform)> CollectFamilyInstancesVisibleOnly(Document doc, string familyNameContains)
        {
            var result = new List<(FamilyInstance, Transform?)>();
            // Host model
            var host = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains(familyNameContains))
                .ToList();
            result.AddRange(host.Select(fi => (fi, (Transform?)null)));

            // Visible linked models
            foreach (var link in new FilteredElementCollector(doc)
                         .OfClass(typeof(RevitLinkInstance))
                         .Cast<RevitLinkInstance>())
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc == null ||
                    doc.ActiveView.GetCategoryHidden(link.Category.Id) ||
                    link.IsHidden(doc.ActiveView))
                    continue;

                var tr = link.GetTotalTransform();
                var linked = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains(familyNameContains))
                    .ToList();
                result.AddRange(linked.Select(fi => (fi, (Transform?)tr)));
            }
            return result;
        }

        /// <summary>
        /// Collects Wall elements (host + visible links).
        /// </summary>
        public static List<(Wall wall, Transform? transform)> CollectWallsVisibleOnly(Document doc)
        {
            var result = new List<(Wall, Transform?)>();
            // Host model
            var host = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();
            result.AddRange(host.Select(w => (w, (Transform?)null)));

            // Visible linked models
            foreach (var link in new FilteredElementCollector(doc)
                         .OfClass(typeof(RevitLinkInstance))
                         .Cast<RevitLinkInstance>())
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc == null ||
                    doc.ActiveView.GetCategoryHidden(link.Category.Id) ||
                    link.IsHidden(doc.ActiveView))
                    continue;

                var tr = link.GetTotalTransform();
                var linked = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(Wall))
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
                result.AddRange(linked.Select(w => (w, (Transform?)tr)));
            }
            return result;
        }
        /// <summary>
        /// Collects MEP elements (host + visible links) that are **inside the active 3-D view’s section box**.
        /// If the view has no active section box, behaviour is identical to the legacy routine.
        /// </summary>
        public static List<(Element element, Transform? transform)> CollectElementsVisibleOnly(
            Document doc, IList<BuiltInCategory> categories)
        {
            // 1. Raw list – same as before
            var raw = CollectRawElements(doc, categories);

            // Log raw counts and small samples for diagnostics (non-invasive)
            try
            {
                int hostRawCount = raw.Count(t => t.transform == null);
                int linkedRawCount = raw.Count - hostRawCount;
                DebugLogger.Info($"[SectionBoxDiag] Raw elements - host={hostRawCount}, linked={linkedRawCount}");
                foreach (var sample in raw.Take(3))
                {
                    var el = sample.element;
                    int idVal = el?.Id.IntegerValue ?? -1;
                    string cat = el?.Category != null ? el.Category.Name : "<no-category>";
                    DebugLogger.Info($"[SectionBoxDiag] Raw sample: Id={idVal}, Category={cat}, IsLinked={(sample.transform!=null)}");
                }
            }
            catch { }

            // If the document is a linked doc we cannot construct a UIDocument for filtering
            if (doc.IsLinked) return raw;

            // Use the SectionBoxHelper which performs solid-based filtering and handles
            // linked elements by transforming the section-box into link-local space.
            var uiDoc = new UIDocument(doc);
            var filtered = SectionBoxHelper.FilterElementsBySectionBox(uiDoc, raw.ToList());

            // Log filtered counts and small samples
            try
            {
                int hostFiltered = filtered.Count(t => t.transform == null);
                int linkedFiltered = filtered.Count - hostFiltered;
                DebugLogger.Info($"[SectionBoxDiag] Filtered elements - host={hostFiltered}, linked={linkedFiltered}");
                foreach (var sample in filtered.Take(3))
                {
                    var el = sample.element;
                    int idVal = el?.Id.IntegerValue ?? -1;
                    string cat = el?.Category != null ? el.Category.Name : "<no-category>";
                    DebugLogger.Info($"[SectionBoxDiag] Filtered sample: Id={idVal}, Category={cat}, IsLinked={(sample.transform!=null)}");
                }
            }
            catch { }

            return filtered;
        }

        /// <summary>Legacy façade – pipes, ducts, cable trays, conduits.</summary>
        public static List<(Element element, Transform? transform)> CollectMepElementsVisibleOnly(Document doc)
        {
            return CollectElementsVisibleOnly(doc, new[]
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit
            });
        }

        /* ---------- private helpers ---------- */

        private static List<(Element element, Transform? transform)> CollectRawElements(
            Document doc, IList<BuiltInCategory> categories)
        {
            var result = new List<(Element, Transform?)>();

            // host model
            var host = new FilteredElementCollector(doc)
                .WherePasses(new ElementMulticategoryFilter(categories))
                .WhereElementIsNotElementType()
                .ToElements();
            foreach (var e in host) result.Add((e, null));

            // visible linked models
            foreach (var link in new FilteredElementCollector(doc)
                         .OfClass(typeof(RevitLinkInstance))
                         .Cast<RevitLinkInstance>())
            {
                var linkDoc = link.GetLinkDocument();
                if (linkDoc == null ||
                    doc.ActiveView.GetCategoryHidden(link.Category.Id) ||
                    link.IsHidden(doc.ActiveView))
                    continue;

                var tr = link.GetTotalTransform();
                var linked = new FilteredElementCollector(linkDoc)
                    .WherePasses(new ElementMulticategoryFilter(categories))
                    .WhereElementIsNotElementType()
                    .ToElements();
                foreach (var e in linked) result.Add((e, tr));
            }
            return result;
        }

        // Simple AABB intersection used by section-box filtering
        private static bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            return !(max1.X < min2.X || min1.X > max2.X ||
                     max1.Y < min2.Y || min1.Y > max2.Y ||
                     max1.Z < min2.Z || min1.Z > max2.Z);
        }
    }
}
