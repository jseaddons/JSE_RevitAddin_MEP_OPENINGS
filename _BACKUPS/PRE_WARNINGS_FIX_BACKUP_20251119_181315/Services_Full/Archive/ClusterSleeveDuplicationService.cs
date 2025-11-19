using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for cluster sleeve duplicate suppression: prevents placing a cluster sleeve over another cluster sleeve.
    /// </summary>
    public static class ClusterSleeveDuplicationService
    {
        /// <summary>
        /// Returns true if a cluster sleeve already exists at the given location within the specified tolerance.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <returns>True if a cluster sleeve exists at the location</returns>
        public static bool IsClusterSleeveAtLocation(Document doc, XYZ location, double tolerance)
        {
            // Delegate to OpeningDuplicationChecker which implements section-box aware
            // cluster bounding-box checks and hostType filtering. This avoids scanning all
            // Cluster families in large models.
            try
            {
                if (doc == null)
                {
                    return false;
                }
                // No hostType provided here; the caller should use host-specific checks when possible.
                // Try to respect the active 3D view section box to avoid scanning the entire document
                BoundingBoxXYZ? sectionBox = null;
                try { if (doc?.ActiveView is View3D vb) sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb); } catch { }
                bool inside = OpeningDuplicationChecker.IsLocationWithinClusterBounds(doc!, location, tolerance, null, sectionBox);
                return inside;
            }
            catch
            {
                // Fallback to conservative scan if the checker fails
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase));

                foreach (var fi in collector)
                {
                    var fiLocation = (fi.Location as LocationPoint)?.Point;
                    if (fiLocation == null) continue;
                    if (location.DistanceTo(fiLocation) <= tolerance)
                        return true;
                }
                return false;
            }
        }
    }
}
