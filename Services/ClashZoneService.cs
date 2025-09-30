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
        /// CLEANUP: Remove invalid clash zones AND existing duplicates
        /// This ensures only valid, unique clash zones remain
        /// </summary>
        public int CleanupInvalidClashZones(Document document)
        {
            if (_clashZoneStorage?.ClashZones == null)
            {
                _log("No clash zones to clean up");
                return 0;
            }

            var originalCount = _clashZoneStorage.ClashZones.Count;
            _log($"Starting cleanup of {originalCount} clash zones");

            // Step 1: Remove invalid clash zones (null elements)
            var validClashZones = new List<ClashZone>();
            var invalidCount = 0;

            foreach (var clashZone in _clashZoneStorage.ClashZones)
            {
                try
                {
                    var mepElement = document.GetElement(clashZone.MepElementId);
                    var structuralElement = document.GetElement(clashZone.StructuralElementId);
                    
                    if (mepElement != null && structuralElement != null)
                    {
                        validClashZones.Add(clashZone);
                    }
                    else
                    {
                        invalidCount++;
                        _log($"Removing invalid clash zone {clashZone.Id} - MEP: {mepElement != null}, Structural: {structuralElement != null}");
                    }
                }
                catch (Exception ex)
                {
                    invalidCount++;
                    _log($"Removing invalid clash zone {clashZone.Id} due to error: {ex.Message}");
                }
            }

            // Step 2: Remove duplicates (keep only the first occurrence of each MEP+Structural pair)
            var uniqueClashZones = new List<ClashZone>();
            var seenPairs = new HashSet<(ElementId, ElementId)>();
            var duplicateCount = 0;

            foreach (var clashZone in validClashZones)
            {
                var pair = (clashZone.MepElementId, clashZone.StructuralElementId);
                
                if (seenPairs.Add(pair))
                {
                    uniqueClashZones.Add(clashZone);
                }
                else
                {
                    duplicateCount++;
                    _log($"Removing duplicate clash zone {clashZone.Id} - MEP: {clashZone.MepElementId}, Structural: {clashZone.StructuralElementId}");
                }
            }

            // Update storage
            _clashZoneStorage.ClashZones.Clear();
            _clashZoneStorage.ClashZones.AddRange(uniqueClashZones);

            var finalCount = _clashZoneStorage.ClashZones.Count;
            var totalRemoved = originalCount - finalCount;
            
            _log($"Cleanup complete: {originalCount} → {finalCount} clash zones (removed {totalRemoved}: {invalidCount} invalid + {duplicateCount} duplicates)");
            
            return totalRemoved;
        }

        /// <summary>
        /// Filters existing clash zones to only include those matching current selection parameters
        /// </summary>
        public List<ClashZone> FilterClashZonesByCurrentSelection(
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document)
        {
            if (_clashZoneStorage?.ClashZones == null)
            {
                _log("No clash zones to filter");
                return new List<ClashZone>();
            }

            _log($"Filtering {_clashZoneStorage.ClashZones.Count} clash zones by current selection");
            _log($"Selected reference files: {string.Join(", ", selectedReferenceFiles)}");
            _log($"Current clearance settings: {string.Join(", ", currentClearanceSettings.Select(kvp => $"{kvp.Key}={kvp.Value}"))}");
            _log($"Current prefix: {currentPrefix}");

            var filteredZones = new List<ClashZone>();
            var removedCount = 0;

            foreach (var clashZone in _clashZoneStorage.ClashZones.ToList())
            {
                // Check if this clash zone matches current selection criteria
                bool matchesCurrentSelection = DoesClashZoneMatchCurrentSelection(
                    clashZone, selectedReferenceFiles, currentClearanceSettings, currentPrefix, document);

                if (matchesCurrentSelection)
                {
                    filteredZones.Add(clashZone);
                }
                else
                {
                    // DON'T remove clash zone from storage - just skip it for this processing
                    removedCount++;
                    _log($"Skipping clash zone {clashZone.Id} - doesn't match current selection (keeping in storage)");
                }
            }

            _log($"Filtered clash zones: {filteredZones.Count} kept, {removedCount} removed");
            
            // Update storage metadata
            _clashZoneStorage.LastUpdated = DateTime.Now;
            
            return filteredZones;
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
                // CRITICAL FIX: Validate elements before creating clash zones
                if (mepElement == null || structuralElement == null)
                {
                    _log($"SKIP: Invalid elements - MEP={mepElement?.Id}, Structural={structuralElement?.Id}");
                    continue;
                }

                // CRITICAL FIX: Check if elements are still valid in the document
                try
                {
                    var mepElementCheck = document.GetElement(mepElement.Id);
                    var structuralElementCheck = document.GetElement(structuralElement.Id);
                    
                    // For linked elements, GetElement might return null even if element exists
                    // Check if element is from a linked document
                    bool mepIsLinked = mepElementCheck == null && mepElement.Id.IntegerValue > 0;
                    bool structuralIsLinked = structuralElementCheck == null && structuralElement.Id.IntegerValue > 0;
                    
                    if (mepElementCheck == null && !mepIsLinked)
                    {
                        _log($"SKIP: MEP element no longer valid in document - MEP={mepElement.Id}");
                        continue;
                    }
                    
                    if (structuralElementCheck == null && !structuralIsLinked)
                    {
                        _log($"SKIP: Structural element no longer valid in document - Structural={structuralElement.Id}");
                        continue;
                    }
                    
                    if (mepIsLinked || structuralIsLinked)
                    {
                        _log($"INFO: Processing linked elements - MEP={mepElement.Id} (linked={mepIsLinked}), Structural={structuralElement.Id} (linked={structuralIsLinked})");
                    }
                }
                catch (Exception ex)
                {
                    _log($"SKIP: Error validating elements - MEP={mepElement.Id}, Structural={structuralElement.Id}, Error={ex.Message}");
                    continue;
                }

                // DIAGNOSTIC: Log detailed geometry information to determine if penetration is real
                _log($"=== GEOMETRY ANALYSIS ===");
                _log($"MEP Element: {mepElement.Name} (ID: {mepElement.Id})");
                _log($"Structural Element: {structuralElement.Name} (ID: {structuralElement.Id})");
                
                // Get MEP element bounding box
                var mepBBox = mepElement.get_BoundingBox(null);
                if (mepBBox != null)
                {
                    _log($"MEP BBox: Min({mepBBox.Min.X:F3}, {mepBBox.Min.Y:F3}, {mepBBox.Min.Z:F3}) Max({mepBBox.Max.X:F3}, {mepBBox.Max.Y:F3}, {mepBBox.Max.Z:F3})");
                    _log($"MEP Size: X={mepBBox.Max.X - mepBBox.Min.X:F3}, Y={mepBBox.Max.Y - mepBBox.Min.Y:F3}, Z={mepBBox.Max.Z - mepBBox.Min.Z:F3}");
                }
                
                // Get structural element bounding box
                var structBBox = structuralElement.get_BoundingBox(null);
                if (structBBox != null)
                {
                    _log($"Structural BBox: Min({structBBox.Min.X:F3}, {structBBox.Min.Y:F3}, {structBBox.Min.Z:F3}) Max({structBBox.Max.X:F3}, {structBBox.Max.Y:F3}, {structBBox.Max.Z:F3})");
                    _log($"Structural Size: X={structBBox.Max.X - structBBox.Min.X:F3}, Y={structBBox.Max.Y - structBBox.Min.Y:F3}, Z={structBBox.Max.Z - structBBox.Min.Z:F3}");
                }
                
                // Log intersection details
                _log($"Intersection Point: ({intersectionPoint.X:F3}, {intersectionPoint.Y:F3}, {intersectionPoint.Z:F3})");
                if (boundingBox != null)
                {
                    _log($"Intersection BBox: Min({boundingBox.Min.X:F3}, {boundingBox.Min.Y:F3}, {boundingBox.Min.Z:F3}) Max({boundingBox.Max.X:F3}, {boundingBox.Max.Y:F3}, {boundingBox.Max.Z:F3})");
                    _log($"Intersection Size: X={boundingBox.Max.X - boundingBox.Min.X:F3}, Y={boundingBox.Max.Y - boundingBox.Min.Y:F3}, Z={boundingBox.Max.Z - boundingBox.Min.Z:F3}");
                }
                _log($"=== END GEOMETRY ANALYSIS ===");

                // Check if this clash zone already exists
                var existingClashZone = FindExistingClashZone(mepElement.Id, structuralElement.Id, intersectionPoint);

                if (existingClashZone == null)
                {
                    // Check if there's an invalid clash zone for the same geometry (by hash)
                    var invalidClashZone = FindInvalidClashZoneByGeometry(mepElement, structuralElement);
                    
                    if (invalidClashZone != null)
                    {
                        // Replace invalid clash zone with valid one
                        _log($"Replacing invalid clash zone {invalidClashZone.Id} with valid ElementIds");
                        _clashZoneStorage.ClashZones.Remove(invalidClashZone);
                        
                        var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document);
                        newClashZones.Add(newClashZone);
                        _clashZoneStorage.ClashZones.Add(newClashZone);
                        _log($"Replaced invalid clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    }
                    else
                {
                    // Create new clash zone
                    var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document);
                    newClashZones.Add(newClashZone);
                    _clashZoneStorage.ClashZones.Add(newClashZone);
                    _log($"New clash zone detected: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    }
                }
                else if (!existingClashZone.IsResolved)
                {
                    // Only update existing clash zone if it's NOT resolved
                    UpdateExistingClashZone(existingClashZone, mepElement, structuralElement, intersectionPoint, boundingBox, document);
                    _log($"Updated existing unresolved clash zone: {existingClashZone.Id}");
                }
                else
                {
                    // Skip resolved clash zones during refresh - they should not be updated
                    _log($"Skipping resolved clash zone: {existingClashZone.Id} (MEP={mepElement.Id}, Structural={structuralElement.Id})");
                }
            }
            
            // Do NOT change IsResolved during refresh. Resolved state is set only after successful placement
            //MarkResolvedClashZones(currentIntersections, document);
            
            // CRITICAL FIX: Remove duplicate clash zones (same MEP + structural element)
            RemoveDuplicateClashZones();
            
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
        /// Marks a clash zone as resolved (individual sleeve placed)
        /// </summary>
        public void MarkClashZoneResolved(Guid clashZoneId, ElementId sleeveId)
        {
            var clashZone = _clashZoneStorage.ClashZones.FirstOrDefault(cz => cz.Id == clashZoneId);
            if (clashZone != null)
            {
                clashZone.IsResolved = true;
                clashZone.ResolvedSleeveId = sleeveId;
                clashZone.LastUpdated = DateTime.Now;
                _log($"Marked clash zone {clashZoneId} as resolved with individual sleeve {sleeveId}");
            }
        }
        
        /// <summary>
        /// Marks a clash zone as cluster resolved (cluster sleeve placed)
        /// </summary>
        public void MarkClashZoneClusterResolved(Guid clashZoneId, ElementId clusterSleeveId)
        {
            var clashZone = _clashZoneStorage.ClashZones.FirstOrDefault(cz => cz.Id == clashZoneId);
            if (clashZone != null)
            {
                clashZone.IsClusterResolved = true;
                clashZone.ClusterSleeveId = clusterSleeveId;
                clashZone.LastUpdated = DateTime.Now;
                _log($"Marked clash zone {clashZoneId} as cluster resolved with cluster sleeve {clusterSleeveId}");
            }
        }
        
        /// <summary>
        /// Marks multiple clash zones as cluster resolved (batch operation)
        /// </summary>
        public void MarkClashZonesClusterResolved(List<Guid> clashZoneIds, ElementId clusterSleeveId)
        {
            foreach (var clashZoneId in clashZoneIds)
            {
                MarkClashZoneClusterResolved(clashZoneId, clusterSleeveId);
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
        
        /// <summary>
        /// Checks if a clash zone matches the current selection parameters - SIMPLE APPROACH
        /// Keep all clash zones found during refresh
        /// </summary>
        private bool DoesClashZoneMatchCurrentSelection(
            ClashZone clashZone,
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document)
        {
            try
            {
                // STEP 1: Check if MEP element is from selected file
                // First check if ElementId is valid (not null)
                if (clashZone.MepElementId == null || clashZone.MepElementId == ElementId.InvalidElementId)
                {
                    _log($"Clash zone {clashZone.Id} has invalid MEP ElementId - REMOVING invalid clash zone");
                    return false;
                }
                
                Element mepElement = null;
                try
                {
                    mepElement = document.GetElement(clashZone.MepElementId);
                }
                catch (Exception ex)
                {
                    _log($"Error getting MEP element {clashZone.MepElementId?.IntegerValue ?? -1} for clash zone {clashZone.Id}: {ex.Message} - REMOVING invalid clash zone");
                    return false; // Remove invalid clash zone
                }
                
                if (mepElement != null)
                {
                    bool isFromSelectedFile = IsElementFromSelectedFile(mepElement, selectedReferenceFiles);
                    if (!isFromSelectedFile)
                    {
                        _log($"Clash zone {clashZone.Id} MEP element not from selected files - skipping");
                        return false;
                    }
                }
                else
                {
                    // MEP element temporarily unavailable - remove invalid clash zone
                    _log($"Clash zone {clashZone.Id} MEP element temporarily unavailable - REMOVING invalid clash zone");
                    return false;
                }

                // STEP 2: Check if clash zone is already resolved (sleeve placed)
                if (clashZone.IsResolved)
                {
                    _log($"Clash zone {clashZone.Id} already resolved (sleeve placed) - skipping");
                    return false;
                }

                // STEP 3: Check if clash zone is visible in current 3D section box
                bool isVisibleInSectionBox = IsClashZoneVisibleInCurrentSectionBox(clashZone, document);
                if (!isVisibleInSectionBox)
                {
                    _log($"Clash zone {clashZone.Id} not visible in current 3D section box - skipping");
                    return false;
                }

                _log($"Clash zone {clashZone.Id} matches current selection (from selected file + not resolved + visible in section box)");
                return true;
            }
            catch (Exception ex)
            {
                _log($"Error checking clash zone {clashZone.Id} against current selection: {ex.Message} - REMOVING invalid clash zone");
                return false; // Remove clash zone on error - it's invalid
            }
        }

        /// <summary>
        /// Checks if a clash zone is visible in the current 3D section box
        /// </summary>
        private bool IsClashZoneVisibleInCurrentSectionBox(ClashZone clashZone, Document document)
        {
            try
            {
                // Check if we have an active 3D view with section box
                if (!(document.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
                {
                    _log($"No active 3D section box - clash zone {clashZone.Id} considered visible");
                    return true; // No section box = all visible
                }

                // Get section box bounds
                var sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                if (sectionBox == null)
                {
                    _log($"Could not get section box bounds - clash zone {clashZone.Id} considered visible");
                    return true; // Can't get bounds = all visible
                }

                // Check if clash zone intersection point is within section box
                var intersectionPoint = clashZone.IntersectionPoint;
                
                bool isVisible = intersectionPoint.X >= sectionBox.Min.X && intersectionPoint.X <= sectionBox.Max.X &&
                               intersectionPoint.Y >= sectionBox.Min.Y && intersectionPoint.Y <= sectionBox.Max.Y &&
                               intersectionPoint.Z >= sectionBox.Min.Z && intersectionPoint.Z <= sectionBox.Max.Z;
                
                if (isVisible)
                {
                    _log($"Clash zone {clashZone.Id} is visible in section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})");
                }
                else
                {
                    _log($"Clash zone {clashZone.Id} is outside section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})");
                }
                
                return isVisible;
            }
            catch (Exception ex)
            {
                _log($"Error checking clash zone {clashZone.Id} section box visibility: {ex.Message} - considering visible");
                return true; // Consider visible on error
            }
        }

        /// <summary>
        /// Checks if an element is from a selected reference file
        /// </summary>
        private bool IsElementFromSelectedFile(Element element, List<string> selectedReferenceFiles)
        {
            try
            {
                // If no reference files selected, assume all elements are valid
                if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                    return true;

                // For linked elements, check if the link is in selected files
                var linkInstances = new FilteredElementCollector(element.Document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null && element.Document == linkDoc)
                    {
                        var linkName = linkInstance.Name;
                        bool isFromSelectedFile = selectedReferenceFiles.Any(file => file.Contains(linkName) || linkName.Contains(file));
                        if (isFromSelectedFile)
                        {
                            _log($"Element {element.Id} is from selected linked file: {linkName}");
                            return true;
                        }
                    }
                }

                // Check if element is from host document by comparing document titles
                // For host elements, we'll check if the document title matches any selected file
                var elementDocTitle = element.Document.Title;
                
                // Check if this document title is in the selected reference files
                bool isFromSelectedHostFile = selectedReferenceFiles.Any(file => 
                    file.Contains(elementDocTitle) || elementDocTitle.Contains(file));
                
                if (isFromSelectedHostFile)
                {
                    _log($"Element {element.Id} is from selected host file: {elementDocTitle}");
                    return true;
                }

                _log($"Element {element.Id} is NOT from any selected file. Document: {elementDocTitle}");
                return false;
            }
            catch (Exception ex)
            {
                _log($"Error checking if element {element.Id} is from selected file: {ex.Message}");
                return true; // Default to true if we can't determine
            }
        }

        /// <summary>
        /// Checks if clash zone clearance matches current clearance settings
        /// </summary>
        private bool DoesClearanceMatch(ClashZone clashZone, Dictionary<string, double> currentClearanceSettings)
        {
            try
            {
                // If no clearance settings, assume match
                if (currentClearanceSettings == null || currentClearanceSettings.Count == 0)
                    return true;

                // Check if the clash zone's required clearance matches any current clearance setting
                foreach (var clearanceSetting in currentClearanceSettings)
                {
                    // Compare clearance in mm (both RequiredClearance and clearanceSetting.Value are in mm)
                    // Allow some tolerance for clearance comparison
                    if (Math.Abs(clashZone.RequiredClearance - clearanceSetting.Value) < 1.0) // 1mm tolerance
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return true; // Default to true if we can't determine
            }
        }

        /// <summary>
        /// Checks if clash zone prefix matches current prefix
        /// </summary>
        private bool DoesPrefixMatch(ClashZone clashZone, string currentPrefix)
        {
            try
            {
                // If no prefix specified, assume match
                if (string.IsNullOrEmpty(currentPrefix))
                    return true;

                // Check if clash zone metadata contains the current prefix
                if (clashZone.Metadata != null && clashZone.Metadata.ContainsKey("Prefix"))
                {
                    return clashZone.Metadata["Prefix"] == currentPrefix;
                }

                // If no prefix metadata, assume match for now
                return true;
            }
            catch
            {
                return true; // Default to true if we can't determine
            }
        }
        
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
            // FIXED: Only check MEP and structural element IDs, not intersection point
            // This prevents creating multiple clash zones for the same physical clash
            return _clashZoneStorage.ClashZones.FirstOrDefault(cz => 
                cz.MepElementId == mepElementId && 
                cz.StructuralElementId == structuralElementId);
        }
        
        /// <summary>
        /// Finds invalid clash zones (with ElementId = -1) that match the current geometry
        /// This method ignores IsResolved status - invalid clash zones should always be replaced
        /// </summary>
        private ClashZone? FindInvalidClashZoneByGeometry(Element mepElement, Element structuralElement)
        {
            if (_clashZoneStorage?.ClashZones == null) return null;
            
            var mepGeometryHash = CalculateElementGeometryHash(mepElement);
            var structuralGeometryHash = CalculateElementGeometryHash(structuralElement);
            
            // Also check for old hash formats that included intersection coordinates
            var mepElementId = mepElement.Id.ToString();
            var structuralElementId = structuralElement.Id.ToString();
            var mepCategory = mepElement.Category?.Name ?? "Unknown";
            var structuralCategory = structuralElement.Category?.Name ?? "Unknown";
            
            return _clashZoneStorage.ClashZones.FirstOrDefault(cz => 
                (cz.MepElementId == null || cz.MepElementId == ElementId.InvalidElementId || cz.MepElementIdValue == -1) &&
                (cz.StructuralElementId == null || cz.StructuralElementId == ElementId.InvalidElementId || cz.StructuralElementIdValue == -1) &&
                (
                    // Check new hash format (elementId_category)
                    (cz.MepElementGeometryHash == mepGeometryHash && cz.StructuralElementGeometryHash == structuralGeometryHash) ||
                    // Check old hash format (elementId_category_(x,y,z)) - for backward compatibility
                    (cz.MepElementGeometryHash.StartsWith($"{mepElementId}_{mepCategory}_") && cz.StructuralElementGeometryHash.StartsWith($"{structuralElementId}_{structuralCategory}_"))
                ));
        }
        
        /// <summary>
        /// Remove duplicate clash zones (same MEP + structural element) keeping the most recent one
        /// </summary>
        private void RemoveDuplicateClashZones()
        {
            var originalCount = _clashZoneStorage.ClashZones.Count;
            
            // Group by MEP + Structural element IDs and keep only the most recent in each group
            var uniqueClashZones = _clashZoneStorage.ClashZones
                .GroupBy(cz => new { cz.MepElementId, cz.StructuralElementId })
                .Select(group => group.OrderByDescending(cz => cz.LastUpdated).First())
                .ToList();
            
            _clashZoneStorage.ClashZones.Clear();
            _clashZoneStorage.ClashZones.AddRange(uniqueClashZones);
            
            var removedCount = originalCount - _clashZoneStorage.ClashZones.Count;
            if (removedCount > 0)
            {
                _log($"DEDUPLICATION: Removed {removedCount} duplicate clash zones. Kept {_clashZoneStorage.ClashZones.Count} unique clash zones.");
            }
        }
        
        private ClashZone CreateClashZone(Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document)
        {
            var mepSize = GetMepElementSize(mepElement);
            var requiredClearance = CalculateRequiredClearance(mepSize);
            
            var clashZone = new ClashZone
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
            
            // DEBUG: Log the intersection point being set
            _log($"[DEBUG] Created ClashZone {clashZone.Id}:");
            _log($"[DEBUG]   IntersectionPoint: {intersectionPoint}");
            _log($"[DEBUG]   IntersectionPointX: {clashZone.IntersectionPointX}");
            _log($"[DEBUG]   IntersectionPointY: {clashZone.IntersectionPointY}");
            _log($"[DEBUG]   IntersectionPointZ: {clashZone.IntersectionPointZ}");
            
            return clashZone;
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
            // Intentionally left as no-op for IsResolved. We only mark IsResolved=true after sleeve placement.
            // Optionally, we could track a transient flag like IsCurrentlyClashing here without touching IsResolved.
            var currentStructuralIds = currentIntersections.Select(i => i.Item2.Id).ToHashSet();
            int noLongerIntersecting = _clashZoneStorage.ClashZones
                .Count(cz => !cz.IsResolved && !currentStructuralIds.Contains(cz.StructuralElementId));
            if (noLongerIntersecting > 0)
            {
                _log($"{noLongerIntersecting} clash zones are not present in this refresh; preserving IsResolved state (no auto-resolve).");
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
        
        /// <summary>
        /// Checks if a bounding box represents a valid intersection (not a tangent contact)
        /// </summary>
        private bool IsInvalidBoundingBox(BoundingBoxXYZ boundingBox)
        {
            if (boundingBox == null || boundingBox.Min == null || boundingBox.Max == null)
                return true;

            // Check if any dimension has zero or negative width (tangent contact)
            var widthX = Math.Abs(boundingBox.Max.X - boundingBox.Min.X);
            var widthY = Math.Abs(boundingBox.Max.Y - boundingBox.Min.Y);
            var widthZ = Math.Abs(boundingBox.Max.Z - boundingBox.Min.Z);

            // If all dimensions are effectively zero, it's a tangent contact
            const double tolerance = 1e-6; // 1 micron tolerance
            return widthX < tolerance && widthY < tolerance && widthZ < tolerance;
        }
        
        private string CalculateElementGeometryHash(Element element)
        {
            try
            {
                // Stable hash based on element ID and category - this should remain constant for the same element
                // regardless of intersection point changes
                return $"{element.Id}_{element.Category?.Name ?? "Unknown"}";
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
            // Return clearance in mm for consistency with UI
            // TODO: Use actual clearance settings from UI instead of hardcoded value
            return 50.0; // 50mm clearance (stored in mm, not feet)
        }
        
        private double CalculateRequiredClearance(double mepSize, Dictionary<string, double> clearanceSettings, Element mepElement)
        {
            try
            {
                // Use actual clearance settings from UI
                double clearanceInMm = 50.0; // Default fallback
                
                if (clearanceSettings != null && clearanceSettings.Count > 0)
                {
                    // Determine if duct is insulated (simplified logic for now)
                    bool isInsulated = false;
                    
                    // Try to get clearance based on element type
                    if (mepElement.Category?.Name == "Ducts")
                    {
                        string clearanceKey = isInsulated ? "ducts_insulated_clearance" : "ducts_normal_clearance";
                        if (clearanceSettings.ContainsKey(clearanceKey))
                        {
                            clearanceInMm = clearanceSettings[clearanceKey];
                        }
                    }
                    else if (mepElement.Category?.Name == "Cable Tray")
                    {
                        string clearanceKey = isInsulated ? "cabletray_other_insulated" : "cabletray_other_normal";
                        if (clearanceSettings.ContainsKey(clearanceKey))
                        {
                            clearanceInMm = clearanceSettings[clearanceKey];
                        }
                    }
                }
                
                // Return clearance directly in mm (UI stores values in mm)
                // Note: mepSize is in Revit internal units (feet), but for now return in mm for consistency
                return clearanceInMm;
            }
            catch (Exception ex)
            {
                _log($"Error calculating clearance: {ex.Message}, using default");
                return mepSize + (50.0 / 304.8); // Fallback to 50mm
            }
        }
        
        private bool IsPointNear(XYZ point1, XYZ point2, double tolerance)
        {
            return point1.DistanceTo(point2) <= tolerance;
        }
        
        #endregion
    }
}
        {
            if (_clashZoneStorage?.ClashZones == null) return null;
            
            var mepGeometryHash = CalculateElementGeometryHash(mepElement);
            var structuralGeometryHash = CalculateElementGeometryHash(structuralElement);
            
            // Also check for old hash formats that included intersection coordinates
            var mepElementId = mepElement.Id.ToString();
            var structuralElementId = structuralElement.Id.ToString();
            var mepCategory = mepElement.Category?.Name ?? "Unknown";
            var structuralCategory = structuralElement.Category?.Name ?? "Unknown";
            
            return _clashZoneStorage.ClashZones.FirstOrDefault(cz => 
                (cz.MepElementId == null || cz.MepElementId == ElementId.InvalidElementId || cz.MepElementIdValue == -1) &&
                (cz.StructuralElementId == null || cz.StructuralElementId == ElementId.InvalidElementId || cz.StructuralElementIdValue == -1) &&
                (
                    // Check new hash format (elementId_category)
                    (cz.MepElementGeometryHash == mepGeometryHash && cz.StructuralElementGeometryHash == structuralGeometryHash) ||
                    // Check old hash format (elementId_category_(x,y,z)) - for backward compatibility
                    (cz.MepElementGeometryHash.StartsWith($"{mepElementId}_{mepCategory}_") && cz.StructuralElementGeometryHash.StartsWith($"{structuralElementId}_{structuralCategory}_"))
                ));
        }
        
        /// <summary>
        /// Remove duplicate clash zones (same MEP + structural element) keeping the most recent one
        /// </summary>
        private void RemoveDuplicateClashZones()
        {
            var originalCount = _clashZoneStorage.ClashZones.Count;
            
            // Group by MEP + Structural element IDs and keep only the most recent in each group
            var uniqueClashZones = _clashZoneStorage.ClashZones
                .GroupBy(cz => new { cz.MepElementId, cz.StructuralElementId })
                .Select(group => group.OrderByDescending(cz => cz.LastUpdated).First())
                .ToList();
            
            _clashZoneStorage.ClashZones.Clear();
            _clashZoneStorage.ClashZones.AddRange(uniqueClashZones);
            
            var removedCount = originalCount - _clashZoneStorage.ClashZones.Count;
            if (removedCount > 0)
            {
                _log($"DEDUPLICATION: Removed {removedCount} duplicate clash zones. Kept {_clashZoneStorage.ClashZones.Count} unique clash zones.");
            }
        }
        
        private ClashZone CreateClashZone(Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document)
        {
            var mepSize = GetMepElementSize(mepElement);
            var requiredClearance = CalculateRequiredClearance(mepSize);
            
            var clashZone = new ClashZone
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
            
            // DEBUG: Log the intersection point being set
            _log($"[DEBUG] Created ClashZone {clashZone.Id}:");
            _log($"[DEBUG]   IntersectionPoint: {intersectionPoint}");
            _log($"[DEBUG]   IntersectionPointX: {clashZone.IntersectionPointX}");
            _log($"[DEBUG]   IntersectionPointY: {clashZone.IntersectionPointY}");
            _log($"[DEBUG]   IntersectionPointZ: {clashZone.IntersectionPointZ}");
            
            return clashZone;
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
            // Intentionally left as no-op for IsResolved. We only mark IsResolved=true after sleeve placement.
            // Optionally, we could track a transient flag like IsCurrentlyClashing here without touching IsResolved.
            var currentStructuralIds = currentIntersections.Select(i => i.Item2.Id).ToHashSet();
            int noLongerIntersecting = _clashZoneStorage.ClashZones
                .Count(cz => !cz.IsResolved && !currentStructuralIds.Contains(cz.StructuralElementId));
            if (noLongerIntersecting > 0)
            {
                _log($"{noLongerIntersecting} clash zones are not present in this refresh; preserving IsResolved state (no auto-resolve).");
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
        
        /// <summary>
        /// Checks if a bounding box represents a valid intersection (not a tangent contact)
        /// </summary>
        private bool IsInvalidBoundingBox(BoundingBoxXYZ boundingBox)
        {
            if (boundingBox == null || boundingBox.Min == null || boundingBox.Max == null)
                return true;

            // Check if any dimension has zero or negative width (tangent contact)
            var widthX = Math.Abs(boundingBox.Max.X - boundingBox.Min.X);
            var widthY = Math.Abs(boundingBox.Max.Y - boundingBox.Min.Y);
            var widthZ = Math.Abs(boundingBox.Max.Z - boundingBox.Min.Z);

            // If all dimensions are effectively zero, it's a tangent contact
            const double tolerance = 1e-6; // 1 micron tolerance
            return widthX < tolerance && widthY < tolerance && widthZ < tolerance;
        }
        
        private string CalculateElementGeometryHash(Element element)
        {
            try
            {
                // Stable hash based on element ID and category - this should remain constant for the same element
                // regardless of intersection point changes
                return $"{element.Id}_{element.Category?.Name ?? "Unknown"}";
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
            // Return clearance in mm for consistency with UI
            // TODO: Use actual clearance settings from UI instead of hardcoded value
            return 50.0; // 50mm clearance (stored in mm, not feet)
        }
        
        private double CalculateRequiredClearance(double mepSize, Dictionary<string, double> clearanceSettings, Element mepElement)
        {
            try
            {
                // Use actual clearance settings from UI
                double clearanceInMm = 50.0; // Default fallback
                
                if (clearanceSettings != null && clearanceSettings.Count > 0)
                {
                    // Determine if duct is insulated (simplified logic for now)
                    bool isInsulated = false;
                    
                    // Try to get clearance based on element type
                    if (mepElement.Category?.Name == "Ducts")
                    {
                        string clearanceKey = isInsulated ? "ducts_insulated_clearance" : "ducts_normal_clearance";
                        if (clearanceSettings.ContainsKey(clearanceKey))
                        {
                            clearanceInMm = clearanceSettings[clearanceKey];
                        }
                    }
                    else if (mepElement.Category?.Name == "Cable Tray")
                    {
                        string clearanceKey = isInsulated ? "cabletray_other_insulated" : "cabletray_other_normal";
                        if (clearanceSettings.ContainsKey(clearanceKey))
                        {
                            clearanceInMm = clearanceSettings[clearanceKey];
                        }
                    }
                }
                
                // Return clearance directly in mm (UI stores values in mm)
                // Note: mepSize is in Revit internal units (feet), but for now return in mm for consistency
                return clearanceInMm;
            }
            catch (Exception ex)
            {
                _log($"Error calculating clearance: {ex.Message}, using default");
                return mepSize + (50.0 / 304.8); // Fallback to 50mm
            }
        }
        
        private bool IsPointNear(XYZ point1, XYZ point2, double tolerance)
        {
            return point1.DistanceTo(point2) <= tolerance;
        }
        
        #endregion
    }
}
