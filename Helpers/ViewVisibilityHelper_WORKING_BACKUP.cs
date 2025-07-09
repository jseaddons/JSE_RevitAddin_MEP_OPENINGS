using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Helper class to collect elements visible in the current view (respects crop box, view filters, and visibility).
    /// This is reusable for any Revit add-in that needs to process only visible elements.
    /// </summary>
    public static class ViewVisibilityHelper_WORKING_BACKUP
    {
        /// <summary>
        /// Collects all visible MEP elements (Pipe, Duct, CableTray) in the current view, from both the active model and all visible linked MEP models.
        /// </summary>
        /// <param name="doc">The active document.</param>
        /// <param name="view">The current view (crop box, filters, and link visibility are respected).</param>
        /// <returns>List of visible MEP elements from active and visible linked models.</returns>
        public static IList<Element> GetAllVisibleMEPElementsIncludingLinks(Document doc, View view)
        {
            var mepCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray
            };

            // Collect from active document
            var result = new List<Element>(GetVisibleElementsOfCategories(doc, view, mepCategories));

            // Collect from visible linked MEP models (NO crop box filtering, just visible links)
            var linkInstances = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => !link.IsHidden(view))
                .ToList();

            foreach (var linkInstance in linkInstances)
            {
                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) continue;

                var linkedElements = new FilteredElementCollector(linkedDoc)
                    .WhereElementIsNotElementType()
                    .Where(e => mepCategories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                    .ToList();

                result.AddRange(linkedElements);
            }

            return result;
        }

        public static IList<Element> GetVisibleElementsOfClass(Document doc, View view, Type elementType)
        {
            if (doc == null || view == null || elementType == null)
                throw new ArgumentNullException();
            return new FilteredElementCollector(doc, view.Id)
                .OfClass(elementType)
                .WhereElementIsNotElementType()
                .ToList();
        }

        public static IList<Element> GetVisibleElementsOfCategories(Document doc, View view, IList<BuiltInCategory> categories)
        {
            if (doc == null || view == null || categories == null)
                throw new ArgumentNullException();
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();
            IList<ElementFilter> filters = categories.Select(cat => (ElementFilter)new ElementCategoryFilter(cat)).ToList();
            var categoryFilter = new LogicalOrFilter(filters);
            return collector.WherePasses(categoryFilter).ToList();
        }

        public static IList<Element> GetVisibleMEPElements(Document doc, View view)
        {
            var mepCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_CableTray
            };
            return GetVisibleElementsOfCategories(doc, view, mepCategories);
        }

        public static IList<Element> GetVisibleStructuralElements(Document doc, View view)
        {
            var structCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_StructuralFraming
            };
            return GetVisibleElementsOfCategories(doc, view, structCategories);
        }
    }
}
