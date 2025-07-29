using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class MepElementCollectorHelper
    {
        // Generic: collect any MEP categories from host and visible links
        public static List<(Element, Transform?)> CollectElementsVisibleOnly(Document doc, IList<BuiltInCategory> categories)
        {
            var elements = new List<(Element, Transform?)>();
            // Host model
            var hostElements = new FilteredElementCollector(doc)
                .WherePasses(new ElementMulticategoryFilter(categories))
                .WhereElementIsNotElementType()
                .ToElements();
            foreach (Element e in hostElements) elements.Add((e, null));
            // Linked models (only visible links)
            var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null || doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(doc.ActiveView)) continue;
                var linkTransform = linkInstance.GetTotalTransform();
                var linkedElements = new FilteredElementCollector(linkDoc)
                    .WherePasses(new ElementMulticategoryFilter(categories))
                    .WhereElementIsNotElementType()
                    .ToElements();
                foreach (Element e in linkedElements) elements.Add((e, linkTransform));
            }
            return elements;
        }

        // Backward compatible: collect standard MEP elements (pipes, ducts, cable trays, conduits)
        public static List<(Element, Transform?)> CollectMepElementsVisibleOnly(Document doc)
        {
            return CollectElementsVisibleOnly(doc, new[] {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit
            });
        }
    }
}
