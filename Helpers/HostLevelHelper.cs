using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class HostLevelHelper
    {
        /// <summary>
        /// Gets the reference Level for a host element (pipe, duct, cable tray, damper).
        /// Returns null if not found.
        /// </summary>
        public static Level GetHostReferenceLevel(Document doc, Element host)
        {
            if (host == null) return null;
            Parameter refLevelParam = host.LookupParameter("Reference Level") ?? host.LookupParameter("Level");
            if (refLevelParam != null && refLevelParam.StorageType == StorageType.ElementId)
            {
                ElementId levelId = refLevelParam.AsElementId();
                if (levelId != ElementId.InvalidElementId)
                    return doc.GetElement(levelId) as Level;
            }
            return null;
        }
    }
}
