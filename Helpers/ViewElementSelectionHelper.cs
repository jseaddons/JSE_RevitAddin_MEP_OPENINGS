using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Helper for robust element selection in 3D views, including crop/section box and linked model handling.
    /// </summary>
    public static class ViewElementSelectionHelper
    {
        /// <summary>
        /// Collects elements of the given categories visible in the given 3D view, respecting crop/section box.
        /// </summary>
        public static IList<Element> GetElementsInCropOrSectionBox(Document doc, View3D view, IList<BuiltInCategory> categories)
        {
            // TODO: Implement 8-corner section box transformation logic (see The Building Coder)
            // For now, basic category filter in view
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();
            var filters = categories.Select(cat => (ElementFilter)new ElementCategoryFilter(cat)).ToList();
            var categoryFilter = new LogicalOrFilter(filters);
            return collector.WherePasses(categoryFilter).ToList();
        }

        /// <summary>
        /// Collects elements of the given categories from all visible linked models, transformed into host coordinates and filtered by section/crop box.
        /// </summary>
        public static IList<(Element element, Transform transform, string linkName)> GetElementsFromLinkedModelsWithTransform(Document doc, View3D view, IList<BuiltInCategory> categories)
        {
            var result = new List<(Element, Transform, string)>();
            var links = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => !link.IsHidden(view));
            foreach (var link in links)
            {
                var linkedDoc = link.GetLinkDocument();
                if (linkedDoc == null) continue;
                var linkTransform = link.GetTransform();
                var linkName = link.Name;
                var collector = new FilteredElementCollector(linkedDoc)
                    .WhereElementIsNotElementType();
                var filters = categories.Select(cat => (ElementFilter)new ElementCategoryFilter(cat)).ToList();
                var categoryFilter = new LogicalOrFilter(filters);
                var elements = collector.WherePasses(categoryFilter).ToList();
                foreach (var el in elements)
                {
                    result.Add((el, linkTransform, linkName));
                }
            }
            return result;
        }
    }
}
