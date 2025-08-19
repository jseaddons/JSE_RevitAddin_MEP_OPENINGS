using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class StructuralElementCollectorHelper
    {
        // Generic: collect any categories from host and visible links
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

        // Backward compatible: collect structural only
        public static List<(Element, Transform?)> CollectStructuralElementsVisibleOnly(Document doc)
        {
            return CollectElementsVisibleOnly(doc, new[] {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            });
        }
    }
} 
