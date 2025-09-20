using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing clash zones and their persistence
    /// </summary>
    public class ClashZoneService
    {
        private readonly ClashZoneStorage _clashZoneStorage;
        private readonly Action<string> _log;
        
        public ClashZoneService(ClashZoneStorage clashZoneStorage, Action<string> log)
        {
            _clashZoneStorage = clashZoneStorage;
            _log = log;
        }
        
        /// <summary>
        /// Detects new clash zones and compares with existing ones
        /// </summary>
        public List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document)
        {
            if (_clashZoneStorage == null)
            {
                _log("ERROR: ClashZoneStorage is null in DetectNewClashZones");
                return new List<ClashZone>();
            }
            
            if (_clashZoneStorage.ClashZones == null)
            {
                _log("ERROR: ClashZones collection is null in DetectNewClashZones");
                return new List<ClashZone>();
            }
            
            var newClashZones = new List<ClashZone>();
            var documentPath = document.PathName;
            var documentHash = CalculateDocumentHash(document);
            
            _log($"Detecting new clash zones for document: {documentPath}");
            _log($"Current intersections count: {currentIntersections.Count}");
            _log($"Existing clash zones count: {_clashZoneStorage.ClashZones.Count}");
            
            foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in currentIntersections)
            {
                // Check if this clash zone already exists
                var existingClashZone = FindExistingClashZone(mepElement.Id, structuralElement.Id, intersectionPoint);

                if (existingClashZone == null)
                {
                    // Create new clash zone
                    var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document);
                    newClashZones.Add(newClashZone);
                    _clashZoneStorage.ClashZones.Add(newClashZone);

                    _log($"New clash zone detected: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                }
                else
                {
                    // Update existing clash zone
                    UpdateExistingClashZone(existingClashZone, mepElement, structuralElement, intersectionPoint, boundingBox, document);
                    _log($"Updated existing clash zone: {existingClashZone.Id}");
                }
            }
            
            // Mark resolved clash zones that are no longer present
            MarkResolvedClashZones(currentIntersections, document);
            
            // Update storage metadata
            _clashZoneStorage.LastUpdated = DateTime.Now;
            _clashZoneStorage.DocumentPath = documentPath;
            _clashZoneStorage.DocumentHash = documentHash;
            
            _log($"Clash zone detection complete. New zones: {newClashZones.Count}");
            return newClashZones;
        }
        
        /// <summary>
        /// Gets clash zones that need to be recalculated due to changes
        /// </summary>
        public List<ClashZone> GetClashZonesNeedingRecalculation(Document document)
        {
            var documentHash = CalculateDocumentHash(document);
            var needsRecalculation = new List<ClashZone>();
            
            // If document hash changed, all clash zones need recalculation
            if (_clashZoneStorage.DocumentHash != documentHash)
            {
                _log($"Document hash changed, all clash zones need recalculation");
                return _clashZoneStorage.ClashZones.ToList();
            }
            
            // Check individual elements for changes
            foreach (var clashZone in _clashZoneStorage.ClashZones)
            {
                if (HasElementChanged(clashZone.MepElementId, clashZone.MepElementGeometryHash, document) ||
                    HasElementChanged(clashZone.StructuralElementId, clashZone.StructuralElementGeometryHash, document))
                {
                    needsRecalculation.Add(clashZone);
                }
            }
            
            _log($"Found {needsRecalculation.Count} clash zones needing recalculation");
            return needsRecalculation;
        }
        
        /// <summary>
        /// Gets all unresolved clash zones
        /// </summary>
        public List<ClashZone> GetUnresolvedClashZones()
        {
            return _clashZoneStorage.ClashZones.Where(cz => !cz.IsResolved).ToList();
        }
        
        /// <summary>
        /// Marks a clash zone as resolved
        /// </summary>
        public void MarkClashZoneResolved(Guid clashZoneId, ElementId sleeveId)
        {
            var clashZone = _clashZoneStorage.ClashZones.FirstOrDefault(cz => cz.Id == clashZoneId);
            if (clashZone != null)
            {
                clashZone.IsResolved = true;
                clashZone.ResolvedSleeveId = sleeveId;
                clashZone.LastUpdated = DateTime.Now;
                _log($"Marked clash zone {clashZoneId} as resolved with sleeve {sleeveId}");
            }
        }
        
        /// <summary>
        /// Clears all clash zones for a fresh start
        /// </summary>
        public void ClearAllClashZones()
        {
            _clashZoneStorage.ClashZones.Clear();
            _clashZoneStorage.LastUpdated = DateTime.Now;
            _log("Cleared all clash zones");
        }
        
        /// <summary>
        /// Gets statistics about clash zones
        /// </summary>
        public (int total, int resolved, int unresolved, int newZones) GetClashZoneStatistics()
        {
            if (_clashZoneStorage == null)
            {
                _log("ERROR: ClashZoneStorage is null in GetClashZoneStatistics");
                return (0, 0, 0, 0);
            }
            
            if (_clashZoneStorage.ClashZones == null)
            {
                _log("ERROR: ClashZones collection is null in GetClashZoneStatistics");
                return (0, 0, 0, 0);
            }
            
            var total = _clashZoneStorage.ClashZones.Count;
            var resolved = _clashZoneStorage.ClashZones.Count(cz => cz.IsResolved);
            var unresolved = total - resolved;
            var newZones = _clashZoneStorage.ClashZones.Count(cz => cz.DetectedAt > DateTime.Now.AddMinutes(-5)); // Last 5 minutes
            
            return (total, resolved, unresolved, newZones);
        }
        
        #region Private Helper Methods
        
        private Element? FindMepElementForIntersection(XYZ intersectionPoint, Document document)
        {
            try
            {
                // Search for MEP elements near the intersection point
                var collector = new FilteredElementCollector(document);
                var mepCategories = new[] { 
                    BuiltInCategory.OST_PipeCurves, 
                    BuiltInCategory.OST_DuctCurves, 
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit
                };
                
                var mepElements = collector
                    .WherePasses(new ElementMulticategoryFilter(mepCategories))
                    .WhereElementIsNotElementType()
                    .ToElements();
                
                // Find the MEP element closest to the intersection point
                Element? closestMepElement = null;
                double closestDistance = double.MaxValue;
                
                foreach (Element mepElement in mepElements)
                {
                    var location = mepElement.Location as LocationCurve;
                    if (location?.Curve is Line line)
                    {
                        // Calculate distance from intersection point to MEP line
                        var distance = line.Distance(intersectionPoint);
                        if (distance < closestDistance && distance < 1.0) // Within 1 foot
                        {
                            closestDistance = distance;
                            closestMepElement = mepElement;
                        }
                    }
                }
                
                _log($"Found MEP element {closestMepElement?.Id} at distance {closestDistance:F2}ft from intersection point");
                return closestMepElement;
            }
            catch (Exception ex)
            {
                _log($"Error finding MEP element for intersection: {ex.Message}");
                return null;
            }
        }
        
        private ClashZone? FindExistingClashZone(ElementId mepElementId, ElementId structuralElementId, XYZ intersectionPoint)
        {
            return _clashZoneStorage.ClashZones.FirstOrDefault(cz => 
                cz.MepElementId == mepElementId && 
                cz.StructuralElementId == structuralElementId &&
                IsPointNear(cz.IntersectionPoint, intersectionPoint, 0.1)); // 0.1 foot tolerance
        }
        
        private ClashZone CreateClashZone(Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document)
        {
            var mepSize = GetMepElementSize(mepElement);
            var requiredClearance = CalculateRequiredClearance(mepSize);
            
            return new ClashZone
            {
                MepElementId = mepElement.Id,
                StructuralElementId = structuralElement.Id,
                IntersectionPoint = intersectionPoint,
                ClashBoundingBox = boundingBox,
                MepElementSize = mepSize,
                RequiredClearance = requiredClearance,
                MepElementGeometryHash = CalculateElementGeometryHash(mepElement),
                StructuralElementGeometryHash = CalculateElementGeometryHash(structuralElement),
                DocumentPath = document.PathName,
                DetectedAt = DateTime.Now,
                LastUpdated = DateTime.Now
            };
        }
        
        private void UpdateExistingClashZone(ClashZone existingZone, Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document)
        {
            existingZone.IntersectionPoint = intersectionPoint;
            existingZone.ClashBoundingBox = boundingBox;
            existingZone.MepElementSize = GetMepElementSize(mepElement);
            existingZone.RequiredClearance = CalculateRequiredClearance(existingZone.MepElementSize);
            existingZone.MepElementGeometryHash = CalculateElementGeometryHash(mepElement);
            existingZone.StructuralElementGeometryHash = CalculateElementGeometryHash(structuralElement);
            existingZone.LastUpdated = DateTime.Now;
        }
        
        private void MarkResolvedClashZones(List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections, Document document)
        {
            var currentStructuralIds = currentIntersections.Select(i => i.Item2.Id).ToHashSet();
            
            foreach (var clashZone in _clashZoneStorage.ClashZones.Where(cz => !cz.IsResolved))
            {
                if (!currentStructuralIds.Contains(clashZone.StructuralElementId))
                {
                    // This structural element is no longer in intersections, mark as resolved
                    clashZone.IsResolved = true;
                    clashZone.LastUpdated = DateTime.Now;
                    _log($"Marked clash zone {clashZone.Id} as resolved (structural element no longer intersecting)");
                }
            }
        }
        
        private bool HasElementChanged(ElementId elementId, string storedHash, Document document)
        {
            try
            {
                var element = document.GetElement(elementId);
                if (element == null) return true; // Element was deleted
                
                var currentHash = CalculateElementGeometryHash(element);
                return currentHash != storedHash;
            }
            catch
            {
                return true; // Assume changed if we can't check
            }
        }
        
        private string CalculateDocumentHash(Document document)
        {
            // Simple hash based on document path and last modified time
            return $"{document.PathName}_{DateTime.Now.Ticks}";
        }
        
        private string CalculateElementGeometryHash(Element element)
        {
            try
            {
                // Simple hash based on element ID and type
                return $"{element.Id}_{element.GetType().Name}_{element.get_BoundingBox(null)?.Min?.ToString() ?? "null"}";
            }
            catch
            {
                return $"error_{element.Id}";
            }
        }
        
        private double GetMepElementSize(Element mepElement)
        {
            try
            {
                var diameterParam = mepElement.LookupParameter("Diameter");
                if (diameterParam != null) return diameterParam.AsDouble();
                
                var sizeParam = mepElement.LookupParameter("Size");
                if (sizeParam != null) return sizeParam.AsDouble();
                
                return 0.5; // Default 6 inches
            }
            catch
            {
                return 0.5;
            }
        }
        
        private double CalculateRequiredClearance(double mepSize)
        {
            // Simple clearance calculation - can be made more sophisticated
            return mepSize + 0.25; // MEP size + 3 inches clearance
        }
        
        private bool IsPointNear(XYZ point1, XYZ point2, double tolerance)
        {
            return point1.DistanceTo(point2) <= tolerance;
        }
        
        #endregion
    }
}
