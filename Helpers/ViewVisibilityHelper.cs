
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For DebugLogger

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Helper class to collect elements visible in the current view (respects crop box, view filters, and visibility).
    /// This is reusable for any Revit add-in that needs to process only visible elements.
    /// </summary>
    public static class ViewVisibilityHelper
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

            // Get the active view from the document if not provided
            View activeView = view ?? doc.ActiveView;
            DebugLogger.Info($"[ViewVisibilityHelper] Using view: {activeView.Name} (Id={activeView.Id.IntegerValue})");

            // Collect only elements visible in the active view (crop box, filters, etc.)
            var visibleInView = new FilteredElementCollector(doc, activeView.Id)
                .WhereElementIsNotElementType()
                .Where(e => mepCategories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                .ToList();
            DebugLogger.Info($"[ViewVisibilityHelper] Active doc: {visibleInView.Count} visible MEP elements in view {activeView.Name} (Id={activeView.Id.IntegerValue})");

            // Collect from visible linked MEP models (NO crop box filtering, just visible links)
            var linkInstances = new FilteredElementCollector(doc, activeView.Id)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => !link.IsHidden(activeView))
                .ToList();

            foreach (var linkInstance in linkInstances)
            {
                var linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) continue;

                var linkedElements = new FilteredElementCollector(linkedDoc)
                    .WhereElementIsNotElementType()
                    .Where(e => mepCategories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                    .ToList();

                DebugLogger.Info($"[ViewVisibilityHelper] Linked doc '{linkedDoc.Title}': {linkedElements.Count} visible MEP elements (link instance: {linkInstance.Name})");
                visibleInView.AddRange(linkedElements);
            }

            DebugLogger.Info($"[ViewVisibilityHelper] Total visible MEP elements (active + links): {visibleInView.Count}");
            return visibleInView;
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
            var elements = collector.WherePasses(categoryFilter).ToList();
            DebugLogger.Info($"[ViewVisibilityHelper] {elements.Count} elements of categories [{string.Join(", ", categories)}] visible in view {view.Name}");
            return elements;
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
