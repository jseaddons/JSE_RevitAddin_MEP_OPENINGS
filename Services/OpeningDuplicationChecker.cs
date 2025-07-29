using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                .Where(fi => fi.Symbol.Family.Name == clusterSymbol.Family.Name);

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

        /// <summary>
        /// Returns all individual sleeves (non-cluster) within tolerance of the given location.
        /// Individual sleeves are families ending with "OpeningOnWall" but NOT ending with "Rect".
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="sleeveType">Type of sleeve to check for (e.g., "DS#", "PS#", "CTS#")</param>
        /// <returns>List of FamilyInstance individual sleeves within tolerance</returns>
        public static List<FamilyInstance> FindIndividualSleevesAtLocation(Document doc, XYZ location, double tolerance, string sleeveType = null)
        {

            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => 
                    (fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab"))
                    && !fi.Symbol.Family.Name.ToLower().Contains("cluster")
                ); // Individual sleeves: family name contains OpeningOnWall or OpeningOnSlab, and does NOT contain 'cluster'

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
        /// Returns all cluster sleeves (rectangular/cluster openings) within tolerance of the given location.
        /// Cluster sleeves are families ending with "Rect" (e.g., PipeOpeningOnWallRect, DuctOpeningOnWallRect).
        /// Uses bounding box intersection for accurate detection since cluster sleeves can be large.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <returns>List of FamilyInstance cluster sleeves within tolerance</returns>
        public static List<FamilyInstance> FindAllClusterSleevesAtLocation(Document doc, XYZ location, double tolerance)
        {
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase)); // Cluster sleeves: family name starts with 'Cluster'

            var found = new List<FamilyInstance>();
            foreach (var fi in collector)
            {
                // Try bounding box intersection first (more accurate for large rectangular openings)
                try
                {
                    var boundingBox = fi.get_BoundingBox(null);
                    if (boundingBox != null)
                    {
                        // Expand bounding box by tolerance
                        var expandedMin = boundingBox.Min - new XYZ(tolerance, tolerance, tolerance);
                        var expandedMax = boundingBox.Max + new XYZ(tolerance, tolerance, tolerance);
                        var expandedBounds = new BoundingBoxXYZ { Min = expandedMin, Max = expandedMax };
                        
                        // Check if location is within expanded bounding box
                        if (location.X >= expandedBounds.Min.X && location.X <= expandedBounds.Max.X &&
                            location.Y >= expandedBounds.Min.Y && location.Y <= expandedBounds.Max.Y &&
                            location.Z >= expandedBounds.Min.Z && location.Z <= expandedBounds.Max.Z)
                        {
                            found.Add(fi);
                            DebugLogger.Log($"[OpeningDuplicationChecker] Cluster {fi.Symbol.Family.Name} (ID:{fi.Id.IntegerValue}) detected via bounding box at location {location}");
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Warning: Could not get bounding box for cluster {fi.Id.IntegerValue}: {ex.Message}");
                }
                
                // Fallback to center point distance if bounding box fails
                var fiLocation = (fi.Location as LocationPoint)?.Point;
                if (fiLocation != null && location.DistanceTo(fiLocation) <= tolerance)
                {
                    found.Add(fi);
                    DebugLogger.Log($"[OpeningDuplicationChecker] Cluster {fi.Symbol.Family.Name} (ID:{fi.Id.IntegerValue}) detected via center point distance at location {location}");
                }
            }
            return found;
        }

        /// <summary>
        /// Checks if a point is within the physical bounds of any cluster sleeve (using bounding box intersection).
        /// This is more accurate than center-point distance for large rectangular cluster openings.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="expansionTolerance">Additional tolerance to expand cluster bounding boxes (internal units)</param>
        /// <returns>True if location is within any cluster sleeve's bounding box</returns>
        public static bool IsLocationWithinClusterBounds(Document doc, XYZ location, double expansionTolerance = 0.0)
        {
            var clusterSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.EndsWith("Rect")); // Cluster sleeves end with "Rect"

            foreach (var cluster in clusterSleeves)
            {
                try
                {
                    var boundingBox = cluster.get_BoundingBox(null);
                    if (boundingBox != null)
                    {
                        // Expand bounding box by tolerance if specified
                        var expandedMin = boundingBox.Min - new XYZ(expansionTolerance, expansionTolerance, expansionTolerance);
                        var expandedMax = boundingBox.Max + new XYZ(expansionTolerance, expansionTolerance, expansionTolerance);
                        
                        // Check if location is within expanded bounding box
                        if (location.X >= expandedMin.X && location.X <= expandedMax.X &&
                            location.Y >= expandedMin.Y && location.Y <= expandedMax.Y &&
                            location.Z >= expandedMin.Z && location.Z <= expandedMax.Z)
                        {
                            DebugLogger.Log($"[OpeningDuplicationChecker] Location {location} is within cluster {cluster.Symbol.Family.Name} (ID:{cluster.Id.IntegerValue}) bounding box");
                            return true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[OpeningDuplicationChecker] Warning: Could not check bounding box for cluster {cluster.Id.IntegerValue}: {ex.Message}");
                }
            }
            return false;
        }

        /// <summary>
        /// Comprehensive check: Returns true if ANY sleeve (individual OR cluster) exists at the location.
        /// This is the method that should be used in sleeve placement commands to prevent conflicts.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <param name="ignoreIds">ElementIds to ignore in the check</param>
        /// <returns>True if any sleeve (individual or cluster) exists at the location</returns>
        public static bool IsAnySleeveAtLocation(Document doc, XYZ location, double tolerance, IEnumerable<ElementId> ignoreIds = null)
        {
            // Check for individual sleeves
            var individualSleeves = FindIndividualSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                individualSleeves = individualSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
            
            if (individualSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {individualSleeves.Count} individual sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                return true;
            }

            // Check for cluster sleeves
            var clusterSleeves = FindAllClusterSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                clusterSleeves = clusterSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
            
            if (clusterSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {clusterSleeves.Count} cluster sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                foreach (var cluster in clusterSleeves)
                {
                    DebugLogger.Log($"  - Cluster: {cluster.Symbol.Family.Name} - {cluster.Symbol.Name} (ID:{cluster.Id.IntegerValue})");
                }
                return true;
            }
            
            return false;
        }

        /// <summary>
        /// Enhanced comprehensive check: Returns true if ANY sleeve (individual OR cluster) exists at the location.
        /// Uses bounding box intersection for cluster sleeves and distance checking for individual sleeves.
        /// This is the most accurate method for sleeve placement duplication checking.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance for individual sleeves (internal units)</param>
        /// <param name="clusterExpansion">Additional expansion for cluster bounding boxes (internal units)</param>
        /// <param name="ignoreIds">ElementIds to ignore in the check</param>
        /// <returns>True if any sleeve (individual or cluster) exists at the location</returns>
        public static bool IsAnySleeveAtLocationEnhanced(Document doc, XYZ location, double tolerance, double clusterExpansion = 0.0, IEnumerable<ElementId> ignoreIds = null)
        {
            // Check for individual sleeves using distance
            var individualSleeves = FindIndividualSleevesAtLocation(doc, location, tolerance);
            if (ignoreIds != null)
                individualSleeves = individualSleeves.Where(fi => !ignoreIds.Contains(fi.Id)).ToList();
            
            if (individualSleeves.Any())
            {
                DebugLogger.Log($"[OpeningDuplicationChecker] Found {individualSleeves.Count} individual sleeves within {UnitUtils.ConvertFromInternalUnits(tolerance, UnitTypeId.Millimeters):F0}mm of location {location}");
                return true;
            }

            // Check for cluster sleeves using bounding box intersection (more accurate)
            if (IsLocationWithinClusterBounds(doc, location, clusterExpansion))
            {
                return true;
            }
            
            return false;
        }

        /// <summary>
        /// Returns a summary of what sleeve types exist at the given location for debugging purposes.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="location">Location to check</param>
        /// <param name="tolerance">Distance tolerance (internal units)</param>
        /// <returns>Description of sleeves found at location</returns>
        public static string GetSleevesSummaryAtLocation(Document doc, XYZ location, double tolerance)
        {
            var individual = FindIndividualSleevesAtLocation(doc, location, tolerance);
            var clusters = FindAllClusterSleevesAtLocation(doc, location, tolerance);
            
            var summary = new StringBuilder();
            if (individual.Any())
            {
                summary.AppendLine($"Individual sleeves: {string.Join(", ", individual.Select(s => $"{s.Symbol.Name} (ID:{s.Id.IntegerValue})"))}");
            }
            if (clusters.Any())
            {
                summary.AppendLine($"Cluster sleeves: {string.Join(", ", clusters.Select(s => $"{s.Symbol.Name} (ID:{s.Id.IntegerValue})"))}");
            }
            if (!individual.Any() && !clusters.Any())
            {
                summary.AppendLine("No sleeves found at location");
            }
            return summary.ToString().TrimEnd();
        }
    }
}
