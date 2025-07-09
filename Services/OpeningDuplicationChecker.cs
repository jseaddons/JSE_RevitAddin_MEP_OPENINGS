using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Provides methods to check for duplicate or overlapping cluster openings in the model.
    /// </summary>
    public static class OpeningDuplicationChecker
    {
        /// <summary>
        /// Returns all cluster openings of the same type within tolerance of the given location.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="clusterSymbol">The FamilySymbol of the cluster being placed</param>
        /// <returns>List of FamilyInstance clusters within tolerance</returns>
        public static List<FamilyInstance> FindClustersAtLocation(Document doc, XYZ location, double tolerance, FamilySymbol clusterSymbol)
        {
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name == clusterSymbol.Family.Name && fi.Symbol.Name == clusterSymbol.Name);

            var found = new List<FamilyInstance>();
            foreach (var fi in collector)
            {
                var fiLocation = (fi.Location as LocationPoint)?.Point;
                if (fiLocation == null) continue;
                if (location.DistanceTo(fiLocation) <= tolerance)
                    found.Add(fi);
            }
            return found;
        }

        /// <summary>
        /// Returns true if any cluster (of the given type) exists at the location, excluding those in the ignore list.
        /// </summary>
        public static bool IsClusterAtLocation(Document doc, XYZ location, double tolerance, FamilySymbol clusterSymbol, IEnumerable<ElementId> ignoreIds = null)
        {
            var clusters = FindClustersAtLocation(doc, location, tolerance, clusterSymbol);
            if (ignoreIds != null)
                clusters = clusters.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
            return clusters.Any();
        }
    }
}
