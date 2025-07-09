using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class ReferenceIntersectorHelper
    {
        public static ReferenceIntersector Create(View3D view, ElementFilter filter = null, object findReferenceTarget = null)
        {
#if REVIT2024
            if (filter != null && findReferenceTarget != null)
            {
                // Use 2024+ constructor
                return new ReferenceIntersector(filter, (FindReferenceTarget)findReferenceTarget, view);
            }
            else if (filter != null)
            {
                // Default to Element target
                return new ReferenceIntersector(filter, FindReferenceTarget.Element, view);
            }
            else
            {
                return new ReferenceIntersector(view);
            }
#else
            // For Revit 2020/2023, only the view-based constructor is available
            return new ReferenceIntersector(view);
#endif
        }

        public static IList<ReferenceWithContext> FindFiltered(
            ReferenceIntersector intersector,
            XYZ origin,
            XYZ direction,
            Document doc,
            Type filterType = null)
        {
            var results = intersector.Find(origin, direction);
            if (filterType != null)
            {
                results = results.Where(r =>
                {
                    var element = doc.GetElement(r.GetReference().ElementId);
                    return filterType.IsInstanceOfType(element);
                }).ToList();
            }
            return results;
        }
    }
}
