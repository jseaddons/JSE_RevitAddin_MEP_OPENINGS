using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class MepElementCollectorHelper
    {
        /// <summary>
        /// Collects MEP elements (host + visible links) that are **inside the active 3-D view’s section box**.
        /// If the view has no active section box, behaviour is identical to the legacy routine.
        /// </summary>
        public static List<(Element element, Transform? transform)> CollectElementsVisibleOnly(
            Document doc, IList<BuiltInCategory> categories)
        {
            // 1. Raw list – same as before
            var raw = CollectRawElements(doc, categories);

            // 2. Optional UIDocument – if we have one, clip by section box
            if (doc.IsLinked)                       // linked-doc case → no UIDoc available
                return raw;

            var uiDoc = new UIDocument(doc);        // host-doc case
            return SectionBoxHelper.FilterElementsBySectionBox(uiDoc, raw);
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
    }
}