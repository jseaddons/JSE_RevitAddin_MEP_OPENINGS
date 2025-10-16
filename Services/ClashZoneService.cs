using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using static JSE_RevitAddin_MEP_OPENINGS.Models.MepCategoryConstants;

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
            Document document,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null)
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
            
            // FOOLPROOF: Auto-detect dampers even if user forgot to select "Duct Accessories" category
            List<(Element, Element, BoundingBoxXYZ, XYZ)> enhancedIntersections;
            List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations;
            
            try
            {
                _log($"[DEBUG] Starting duct-damper optimization with {currentIntersections.Count} intersections");
                enhancedIntersections = AutoDetectMissingDampers(document, currentIntersections);
                _log($"[FOOLPROOF] Enhanced intersections: {enhancedIntersections.Count} (Original: {currentIntersections.Count})");
                
                // OPTIMIZATION: Pre-calculate damper locations from XML + enhanced intersections (Calculate Once, Use Many Times)
                damperLocations = PreCalculateDamperLocationsFromXmlAndCurrent(document, enhancedIntersections);
                _log($"[OPTIMIZATION] Pre-calculated {damperLocations.Count} damper locations from XML + enhanced intersections");
            }
            catch (Exception ex)
            {
                _log($"[ERROR] Failed in foolproof/optimization methods: {ex.Message}");
                // Fallback to original intersections
                enhancedIntersections = currentIntersections;
                damperLocations = new List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)>();
            }
            
            _log($"Detecting new clash zones for document: {documentPath}");
            _log($"Current intersections count: {currentIntersections.Count}");
            
            // FOOLPROOF METHOD: Always process dampers first, then ducts (category-based priority)
            List<(Element, Element, BoundingBoxXYZ, XYZ)> prioritizedIntersections;
            try
            {
                prioritizedIntersections = PrioritizeIntersectionsByCategory(enhancedIntersections);
                _log($"[PRIORITY] Processed {prioritizedIntersections.Count} intersections in priority order");
            }
            catch (Exception ex)
            {
                _log($"[ERROR] Failed in priority method: {ex.Message}");
                prioritizedIntersections = enhancedIntersections;
            }
            
            // DEBUG: Log each intersection being processed in priority order
            for (int i = 0; i < prioritizedIntersections.Count; i++)
            {
                var (mepElement, structuralElement, boundingBox, intersectionPoint) = prioritizedIntersections[i];
                var mepCategory = GetElementCategoryName(mepElement);
                _log($"[DEBUG] Processing intersection {i + 1}/{prioritizedIntersections.Count}: MEP={mepElement.Id} ({mepCategory}) <-> Structural={structuralElement.Id}");
            }
            _log($"Existing clash zones count: {_clashZoneStorage.ClashZones.Count}");
            
            foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in prioritizedIntersections)
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

                // Duct-Damper priority filter: skip ducts when damper is present (cost-effective)
                try
                {
                    var mepCat = GetElementCategoryName(mepElement);
                    _log($"[DUCT-DAMPER] Checking element {mepElement.Id} with category '{mepCat}' against {damperLocations.Count} damper locations");
                    if (string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase))
                    {
                        if (IsDuctNearDamper(mepElement, damperLocations))
                        {
                            _log($"SKIP: Duct {mepElement.Id} - damper present in same intersection, prioritizing damper sleeve");
                            continue;
                        }
                        else
                        {
                            _log($"[DUCT-DAMPER] Duct {mepElement.Id} - no damper nearby, proceeding with sleeve placement");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log($"Error in duct-damper priority filter: {ex.Message}");
                }
                // Penetration adequacy filter: skip shallow/grazing intersections
                try
                {
                    // Apply penetration adequacy only to Walls and Floors; skip for Structural Framing
                    var hostTypeName = GetStructuralElementType(structuralElement);
                    if (hostTypeName == "Wall" || hostTypeName == "Floor")
                    {
                        var hostThickness = GetElementThickness(structuralElement);
                        var mepDir = GetMepElementOrientation(mepElement);
                        var hostNormal = GetStructuralElementNormal(structuralElement);
                        var (mepW, mepH) = GetMepElementDimensions(mepElement);
                        var mepCat = GetElementCategoryName(mepElement);

                        _log($"[DEBUG] Penetration calculation for {mepCat}: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                        _log($"[DEBUG]   MEP Direction: ({mepDir.X:F3}, {mepDir.Y:F3}, {mepDir.Z:F3})");
                        _log($"[DEBUG]   Host Normal: ({hostNormal.X:F3}, {hostNormal.Y:F3}, {hostNormal.Z:F3})");
                        _log($"[DEBUG]   Host Thickness: {hostThickness:F3}");
                        _log($"[DEBUG]   MEP Dimensions: W={mepW:F3}, H={mepH:F3}");

                        double crossSize = 0.0;
                        if (string.Equals(mepCat, "Pipes", StringComparison.OrdinalIgnoreCase))
                        {
                            crossSize = Math.Max(mepW, mepH); // diameter in feet
                        }
                        else
                        {
                            // ducts, trays, accessories – use smaller side of the rectangle
                            crossSize = Math.Min(mepW, mepH);
                        }

                        // Normalize vectors (defensive)
                        var dLen = Math.Sqrt(mepDir.X * mepDir.X + mepDir.Y * mepDir.Y + mepDir.Z * mepDir.Z);
                        var nLen = Math.Sqrt(hostNormal.X * hostNormal.X + hostNormal.Y * hostNormal.Y + hostNormal.Z * hostNormal.Z);

                        double penetrationRatio = 1.0; // default pass-through if not computable
                        if (hostThickness > 1e-6 && crossSize > 1e-6 && dLen > 1e-6 && nLen > 1e-6)
                        {
                            var dNorm = new XYZ(mepDir.X / dLen, mepDir.Y / dLen, mepDir.Z / dLen);
                            var nNorm = new XYZ(hostNormal.X / nLen, hostNormal.Y / nLen, hostNormal.Z / nLen);
                            var dot = Math.Abs(dNorm.X * nNorm.X + dNorm.Y * nNorm.Y + dNorm.Z * nNorm.Z);
                            penetrationRatio = (hostThickness * dot) / crossSize;
                            
                            _log($"[DEBUG]   Normalized MEP Dir: ({dNorm.X:F3}, {dNorm.Y:F3}, {dNorm.Z:F3})");
                            _log($"[DEBUG]   Normalized Host Normal: ({nNorm.X:F3}, {nNorm.Y:F3}, {nNorm.Z:F3})");
                            _log($"[DEBUG]   Dot Product: {dot:F3}");
                            _log($"[DEBUG]   Cross Size: {crossSize:F3}");
                        }

                        // Threshold: require at least 5% penetration of cross-section (reduced for cable trays)
                        const double MinPenetrationRatio = 0.05;
                        _log($"[DEBUG] Penetration check: ratio={penetrationRatio:F3}, threshold={MinPenetrationRatio:F2}, hostThickness={hostThickness:F3}, crossSize={crossSize:F3}");
                        if (penetrationRatio < MinPenetrationRatio)
                        {
                            _log($"SKIP: Insufficient penetration (ratio={penetrationRatio:F3} < {MinPenetrationRatio:F2}) for {hostTypeName}. MEP={mepElement.Id}, Structural={structuralElement.Id}");
                            continue;
                        }
                        _log($"[DEBUG] Penetration check PASSED: ratio={penetrationRatio:F3} >= {MinPenetrationRatio:F2}");
                    }
                    else if (hostTypeName == "Structural Framing")
                    {
                        _log("[PENETRATION] Bypassing penetration filter for Structural Framing (using legacy behavior)");
                    }
                }
                catch (Exception ex)
                {
                    _log($"WARN: Penetration filter failed ({ex.Message}) – proceeding without filter.");
                }

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
                        
                        var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document, clearanceSettings);
                        newClashZones.Add(newClashZone);
                        _clashZoneStorage.ClashZones.Add(newClashZone);
                        _log($"Replaced invalid clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    }
                    else
                {
                    // Create new clash zone
                    _log($"[DEBUG] About to create clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    try
                    {
                        var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document, clearanceSettings);
                        newClashZones.Add(newClashZone);
                        _clashZoneStorage.ClashZones.Add(newClashZone);
                        _log($"New clash zone detected: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    }
                    catch (Exception ex)
                    {
                        _log($"[ERROR] Failed to create clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}, Error={ex.Message}");
                    }
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
                    // Preserve resolved clash zones during refresh - keep them in the list
                    _log($"Preserved resolved clash zone: {existingClashZone.Id} (IsResolved={existingClashZone.IsResolved})");
                }
            }
            
            // CRITICAL: Check for existing sleeves and reset IsResolved flag if sleeves were deleted
            // This allows re-placement of sleeves after manual deletion
            // ⚠️ CRITICAL FIX: Only reset flags for selected categories to avoid affecting other filters/categories
            ResetResolvedFlagForDeletedSleeves(document, selectedCategories);
            
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
        /// Resets all resolved flags to allow re-placement of sleeves
        /// This is useful when user wants to place sleeves again after refresh
        /// </summary>
        public void ResetAllResolvedFlags()
        {
            int resetCount = 0;
            foreach (var clashZone in _clashZoneStorage.ClashZones)
            {
                if (clashZone.IsResolved || clashZone.IsClusterResolved)
                {
                    clashZone.IsResolved = false;
                    clashZone.IsClusterResolved = false;
                    clashZone.ResolvedSleeveId = null;
                    clashZone.ClusterSleeveId = null;
                    clashZone.SleeveInstanceId = -1;
                    clashZone.SleeveFamilyName = string.Empty;
                    clashZone.LastUpdated = DateTime.Now;
                    resetCount++;
                }
            }
            
            if (resetCount > 0)
            {
                _clashZoneStorage.LastUpdated = DateTime.Now;
                _log($"Reset resolved flags for {resetCount} clash zones to allow re-placement");
            }
            else
            {
                _log("No resolved clash zones found to reset");
            }
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
                    // Try to get element from host document first
                    mepElement = document.GetElement(clashZone.MepElementId);
                    
                    // If element is null, try to find it in linked documents
                    if (mepElement == null)
                    {
                        var linkInstances = new FilteredElementCollector(document)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>();

                        foreach (var linkInstance in linkInstances)
                        {
                            var linkDoc = linkInstance.GetLinkDocument();
                            if (linkDoc != null)
                            {
                                try
                                {
                                    mepElement = linkDoc.GetElement(clashZone.MepElementId);
                                    if (mepElement != null)
                                    {
                                        _log($"Found MEP element {clashZone.MepElementId?.IntegerValue ?? -1} in linked document {linkDoc.Title}");
                                        break;
                                    }
                                }
                                catch (Exception linkEx)
                                {
                                    // Continue searching in other linked documents
                                    continue;
                                }
                            }
                        }
                    }
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

                // STEP 4: Check if host type matches current UI selection
                bool hostTypeMatches = DoesHostTypeMatchCurrentSelection(clashZone);
                if (!hostTypeMatches)
                {
                    _log($"Clash zone {clashZone.Id} host type '{clashZone.StructuralElementType}' doesn't match current UI selection - skipping");
                    return false;
                }

                _log($"Clash zone {clashZone.Id} matches current selection (from selected file + not resolved + visible in section box + host type matches)");
                return true;
            }
            catch (Exception ex)
            {
                _log($"Error checking clash zone {clashZone.Id} against current selection: {ex.Message} - REMOVING invalid clash zone");
                return false; // Remove clash zone on error - it's invalid
            }
        }

        /// <summary>
        /// Check if host type matches current UI selection
        /// </summary>
        private bool DoesHostTypeMatchCurrentSelection(ClashZone clashZone)
        {
            try
            {
                // Get currently selected host types from UI
                var selectedHostTypes = FilterUiStateProvider.GetSelectedHostElementTypes?.Invoke() ?? new List<string>();
                
                if (selectedHostTypes.Count == 0)
                {
                    // If no host types selected, allow all (backward compatibility)
                    return true;
                }
                
                // Handle plural/singular mismatch: "Walls" (UI) vs "Wall" (Revit)
                bool hostTypeMatch = selectedHostTypes.Contains(clashZone.StructuralElementType) ||
                                   selectedHostTypes.Contains(clashZone.StructuralElementType + "s") ||
                                   selectedHostTypes.Any(t => t.TrimEnd('s').Equals(clashZone.StructuralElementType, StringComparison.OrdinalIgnoreCase));
                
                _log($"[HOST_TYPE_FILTER] ClashZone {clashZone.Id}: StructuralElementType='{clashZone.StructuralElementType}', SelectedHostTypes=[{string.Join(", ", selectedHostTypes)}], Match={hostTypeMatch}");
                
                return hostTypeMatch;
            }
            catch (Exception ex)
            {
                _log($"[HOST_TYPE_FILTER] Error checking host type for clash zone {clashZone.Id}: {ex.Message}");
                return true; // Default to allowing on error (backward compatibility)
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
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Reset IsResolved flag for clash zones where sleeves no longer exist
        /// This allows re-placement of sleeves after manual deletion
        /// Called during refresh to detect deleted sleeves
        /// ⚠️ CRITICAL FIX: Only reset flags for clash zones in current UI context (selected filters/categories)
        /// </summary>
        private void ResetResolvedFlagForDeletedSleeves(Document document, List<string> selectedCategories = null)
        {
            try
            {
                int resetCount = 0;
                
                // ⚠️ CRITICAL FIX: Only process clash zones that match current UI context
                var clashZonesToCheck = _clashZoneStorage.ClashZones;
                
                if (selectedCategories != null && selectedCategories.Count > 0)
                {
                    // Filter to only clash zones from selected categories
                    clashZonesToCheck = _clashZoneStorage.ClashZones
                        .Where(cz => selectedCategories.Contains(cz.MepElementCategory, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                    
                    _log($"[ResetResolvedFlag] Filtering to {clashZonesToCheck.Count} clash zones from selected categories: {string.Join(", ", selectedCategories)}");
                }
                else
                {
                    // ⚠️ CRITICAL FIX: If no categories specified, DO NOT process any clash zones
                    // This prevents accidental reset of IsResolved flags during dialog initialization
                    _log($"[ResetResolvedFlag] No category filter specified - SKIPPING reset to prevent accidental sleeve deletion");
                    return;
                }
                
                foreach (var clashZone in clashZonesToCheck)
                {
                    if (clashZone.IsResolved)
                    {
                        // Validate placement point exists
                        if (clashZone.SleevePlacementPoint == null)
                        {
                            _log($"[ResetResolvedFlag] WARNING: ClashZone {clashZone.Id} has null SleevePlacementPoint - skipping");
                            continue;
                        }
                        
                        // Check if sleeve actually exists at the placement point
                        bool sleeveExists = CheckForExistingSleeve(clashZone.SleevePlacementPoint, document);
                        
                        _log($"[ResetResolvedFlag] Checking clash zone {clashZone.Id} ({clashZone.MepElementCategory}): sleeveExists={sleeveExists}, IsResolved={clashZone.IsResolved}");
                        
                        if (!sleeveExists)
                        {
                            clashZone.IsResolved = false;
                            clashZone.LastUpdated = DateTime.Now;
                            resetCount++;
                            _log($"[ResetResolvedFlag] ✓ Reset IsResolved to FALSE for clash zone {clashZone.Id} ({clashZone.MepElementCategory}) - sleeve no longer exists");
                        }
                        else
                        {
                            _log($"[ResetResolvedFlag] Sleeve still exists at {clashZone.SleevePlacementPoint} - keeping IsResolved=true");
                        }
                    }
                }
                
                if (resetCount > 0)
                {
                    _log($"[ResetResolvedFlag] Reset IsResolved flag for {resetCount} clash zones where sleeves were deleted (from selected categories only)");
                }
            }
            catch (Exception ex)
            {
                _log($"[ResetResolvedFlag] Error resetting IsResolved flags: {ex.Message}");
            }
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
        
        private ClashZone CreateClashZone(Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document, Dictionary<string, double> clearanceSettings = null)
        {
            // IMPORTANT: The intersection point is already at the wall center (mid-plane)
            // The MepIntersectionService finds intersections with wall faces and CreateBoundingBox()
            // averages entry/exit points, giving us the wall center automatically.
            // No additional offset needed since family insertion point is at middle.
            
            // DEBUG: Log placement point calculation
            _log($"[DEBUG] Placement Point Calculation for {structuralElement.Id}:");
            _log($"[DEBUG]   Intersection Point: {intersectionPoint} (already at wall center - using as placement point)");
            
            // OPTIMIZATION: Store structural element type and thickness for depth calculation
            var structuralElementType = GetStructuralElementType(structuralElement);
            _log($"[DEBUG] StructuralElementType for {structuralElement.Id}: '{structuralElementType}' (Element: {structuralElement.GetType().Name})");
            
            // OPTIMIZATION: Calculate MEP element dimensions and orientation during refresh
            var (mepWidth, mepHeight) = GetMepElementDimensions(mepElement);
            var mepOrientation = GetMepElementOrientation(mepElement);
            
            // 🛡️ ARCHITECTURE FIX: Store RAW dimensions only (no pre-calculated clearance)
            // All clearance (simple and complex) will be handled by CONDITIONS service during placement
            // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
            var mepCategoryForClearance = GetElementCategoryName(mepElement);
            
            // Store raw dimensions for ALL categories - clearance handled by CONDITIONS service
            double finalWidth = mepWidth;
            double finalHeight = mepHeight;
            
            DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id}: Raw dimensions {mepWidth:F3}x{mepHeight:F3} (clearance will be handled by CONDITIONS service during placement)");
            
            // Get pipe opening type if applicable
            var pipeOpeningType = GetPipeOpeningType(mepElement);
            
            // OPTIMIZATION: Get MEP element level information during refresh (no linked file access needed during placement)
            var (levelName, levelElevation) = GetMepElementLevelInfo(mepElement);
            
            // Check for existing sleeve at intersection point (which is the placement point)
            var hasExistingSleeve = CheckForExistingSleeve(intersectionPoint, document);
            
            // ⚠️ CRITICAL: Get MEP element category for category-specific processing ⚠️
            // DO NOT REMOVE: This is essential for each placement service to validate its category
            var mepCategory = GetElementCategoryName(mepElement);
            DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id} ({mepElement.GetType().Name}): Category='{mepCategory}', Element.Category.Name='{mepElement.Category?.Name}'");
            
            // ⚠️ CRITICAL: Get duct shape from family name (Round or Rectangular) ⚠️
            // DO NOT REMOVE: This determines correct sleeve family selection for round vs rectangular ducts
            var ductShape = GetDuctShape(mepElement);
            
            // ⚠️ CRITICAL: Detect insulation type (Normal or Insulated) ⚠️
            // DO NOT REMOVE: This determines which clearance value to use (normal vs insulated)
            var insulationType = GetInsulationType(mepElement);
            
            // Pre-calculate formatted size and system abbreviation to eliminate linked file access during placement
            var formattedSize = FormatMepElementSize(mepElement, mepWidth, mepHeight, ductShape);
            var systemAbbreviation = GetMepSystemAbbreviation(mepElement);
            
            // For fire dampers: detect MSFD type and connector side
            var (isMSFD, connectorSide) = GetDamperConnectorInfo(mepElement, mepCategory);
            
            // Placement point: intersection point is already at host center for FULL penetrations
            // MepIntersectionService.GetBoundingBoxCenter() averages entry/exit points to get mid-depth
            // NOTE: For partial penetrations that pass the 20% filter, intersection point is used as-is
            // (the working code doesn't have special handling for partial penetrations)
            XYZ placementPoint = intersectionPoint;

            var clashZone = new ClashZone
            {
                MepElementId = mepElement.Id,
                StructuralElementId = structuralElement.Id,
                IntersectionPoint = intersectionPoint,
                ClashBoundingBox = boundingBox,
                MepElementSize = 0.0, // Legacy field, not used
                RequiredClearance = 0.0, // Clearance will be calculated during placement
                MepElementGeometryHash = CalculateElementGeometryHash(mepElement),
                StructuralElementGeometryHash = CalculateElementGeometryHash(structuralElement),
                MepElementCategory = MepCategoryConstants.Normalize(mepCategory), // Store STANDARDIZED category name
                DuctShape = ductShape, // Store duct shape (Round/Rectangular) from family name
                InsulationType = insulationType, // Store insulation type (Normal/Insulated) for clearance selection
                DocumentPath = document.PathName,
                StructuralElementDocumentTitle = structuralElement.Document.Title,
                StructuralElementType = structuralElementType,
                HostOrientation = GetHostOrientation(structuralElement), // Pre-calculate orientation (X/Y for walls/framing)
                StructuralElementThickness = GetElementThickness(structuralElement),
                
                StructuralElementNormal = GetStructuralElementNormal(structuralElement), // Pre-calculate normal/direction for orientation
                
                // NEW: Pre-calculated placement data (calculated during refresh, used during placement)
                SleevePlacementPoint = placementPoint,
                MepElementWidth = finalWidth,
                MepElementHeight = finalHeight,
                MepElementOrientationDirection = GetMepElementOrientationFromBbox(mepElement),
                PipeOpeningType = pipeOpeningType,
                MepElementLevelName = levelName,
                MepElementLevelElevation = levelElevation,
                MepElementUniqueId = mepElement?.UniqueId ?? string.Empty, // Pre-calculated unique ID for robust tracking
                MepElementFormattedSize = formattedSize, // Pre-calculated formatted size (e.g., "600x300", "Ø200")
                MepElementSystemAbbreviation = systemAbbreviation, // Pre-calculated system abbreviation (e.g., "SA", "RA")
                DamperConnectorSide = connectorSide, // Pre-calculated connector side for MSFD dampers ("Left", "Right", "Top", "Bottom")
                IsMSFDDamper = isMSFD, // Pre-calculated MSFD flag for offset calculation during placement
                IsResolved = hasExistingSleeve,
                
                DetectedAt = DateTime.Now,
                LastUpdated = DateTime.Now
            };
            try
            {
                var th = clashZone.StructuralElementThickness;
                var thMm = UnitUtils.ConvertFromInternalUnits(th, UnitTypeId.Millimeters);
                Services.DebugLogger.Info($"[CLASH-THICKNESS-ASSIGN] structuralId={structuralElement.Id.IntegerValue} thickness={th:F6}ft ({thMm:F1}mm)");
            }
            catch { }
            
            // DEBUG: Log the pre-calculated data
            _log($"[DEBUG] Created ClashZone {clashZone.Id}:");
            _log($"[DEBUG]   IntersectionPoint: {intersectionPoint} (used as placement point)");
            _log($"[DEBUG]   SleevePlacementPoint: {intersectionPoint} (same as intersection point)");
            _log($"[DEBUG]   MepElementWidth: {finalWidth}, MepElementHeight: {finalHeight}");
            _log($"[DEBUG]   MepElementOrientation: {mepOrientation}");
            _log($"[DEBUG]   PipeOpeningType: {pipeOpeningType}");
            _log($"[DEBUG]   IsResolved: {hasExistingSleeve}");
            
            return clashZone;
        }
        
        /// <summary>
        /// Get structural element type name for depth calculation
        /// </summary>
        /// <summary>
        /// Get host orientation (X or Y) for walls and structural framing based on their direction
        /// Pre-calculated during refresh for efficient clustering
        /// </summary>
        private string GetHostOrientation(Element structuralElement)
        {
            try
            {
                if (structuralElement is Wall wall)
                {
                    if (wall.Location is LocationCurve locationCurve)
                    {
                        var curve = locationCurve.Curve;
                        if (curve is Line line)
                        {
                            var direction = line.Direction.Normalize();
                            
                            // Check if wall is more aligned with X or Y axis
                            double absX = Math.Abs(direction.X);
                            double absY = Math.Abs(direction.Y);
                            
                            // If wall runs along X axis (direction is primarily in X), orientation is X
                            // If wall runs along Y axis (direction is primarily in Y), orientation is Y
                            if (absX > absY)
                            {
                                _log($"[HOST-ORIENT] Wall {wall.Id}: Direction=({direction.X:F3},{direction.Y:F3}), absX={absX:F3} > absY={absY:F3} → Orientation=X");
                                return "X"; // Wall runs along X axis
                            }
                            else
                            {
                                _log($"[HOST-ORIENT] Wall {wall.Id}: Direction=({direction.X:F3},{direction.Y:F3}), absY={absY:F3} > absX={absX:F3} → Orientation=Y");
                                return "Y"; // Wall runs along Y axis
                            }
                        }
                    }
                }
                else if (structuralElement is FamilyInstance famInst && 
                         famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    if (famInst.Location is LocationCurve locationCurve)
                    {
                        var curve = locationCurve.Curve;
                        if (curve is Line line)
                        {
                            var direction = line.Direction.Normalize();
                            
                            // Same logic as walls
                            double absX = Math.Abs(direction.X);
                            double absY = Math.Abs(direction.Y);
                            
                            if (absX > absY)
                            {
                                _log($"[HOST-ORIENT] Framing {famInst.Id}: Direction=({direction.X:F3},{direction.Y:F3}), absX={absX:F3} > absY={absY:F3} → Orientation=X");
                                return "X";
                            }
                            else
                            {
                                _log($"[HOST-ORIENT] Framing {famInst.Id}: Direction=({direction.X:F3},{direction.Y:F3}), absY={absY:F3} > absX={absX:F3} → Orientation=Y");
                                return "Y";
                            }
                        }
                    }
                }
                
                // Floors don't need orientation
                return "";
            }
            catch (Exception ex)
            {
                _log($"[HOST-ORIENT] Error determining orientation: {ex.Message}");
                return "";
            }
        }

        private string GetStructuralElementType(Element element)
        {
            if (element is Wall)
                return "Wall";
            else if (element is Floor)
                return "Floor";
            else if (element is FamilyInstance famInst && 
                     famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                return "Structural Framing";
            else
                return "Unknown";
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Get MEP element category name for category-specific processing
        /// This is essential for validating that each placement service only processes its own category
        /// Prevents cross-category contamination (e.g., pipes in duct service)
        /// </summary>
        private string GetElementCategoryName(Element element)
        {
            try
            {
                // ⚠️ CRITICAL FIX: Check element type FIRST for Duct elements
                // This prevents Duct elements from being misclassified as "Duct Accessories"
                if (element is Autodesk.Revit.DB.Mechanical.Duct) return "Ducts";
                if (element is Autodesk.Revit.DB.Plumbing.Pipe) return "Pipes";
                if (element is Autodesk.Revit.DB.Electrical.CableTray) return "Cable Trays";
                
                // Get category name from element (fallback method)
                var categoryName = element?.Category?.Name;
                if (!string.IsNullOrEmpty(categoryName))
                {
                    return categoryName;
                }
                
                // Check by category ID for Duct Accessories
                if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                    return "Duct Accessories";
                
                return "Unknown";
            }
            catch (Exception ex)
            {
                _log($"Error getting element category name: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Detect if MEP element is insulated (checks InsulationThickness parameter)
        /// Returns "Insulated" or "Normal" - used to select appropriate clearance
        /// </summary>
        private string GetInsulationType(Element element)
        {
            try
            {
                // Check for InsulationThickness parameter (common for ducts and pipes)
                var insulationParam = element.LookupParameter("InsulationThickness") ?? 
                                     element.LookupParameter("Insulation Thickness") ??
                                     element.LookupParameter("Insulation");
                
                if (insulationParam != null && insulationParam.AsDouble() > 0.0)
                {
                    double insulationMm = UnitUtils.ConvertFromInternalUnits(insulationParam.AsDouble(), UnitTypeId.Millimeters);
                    _log($"[DEBUG] Element {element.Id} has insulation: {insulationMm:F1}mm");
                    return "Insulated";
                }
                
                return "Normal";
            }
            catch (Exception ex)
            {
                _log($"Error detecting insulation type: {ex.Message}");
                return "Normal"; // Safe default
            }
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Get duct shape from family name (Round or Rectangular)
        /// This is essential for determining correct sleeve family (DuctOpeningOnWall vs DuctOpeningOnWallround)
        /// Round ducts use UI selection, Rectangular ducts always use rectangular sleeves
        /// </summary>
        private string GetDuctShape(Element element)
        {
            try
            {
                if (!(element is Autodesk.Revit.DB.Mechanical.Duct))
                    return string.Empty; // Not a duct
                
                // Get duct type family name
                var duct = element as Autodesk.Revit.DB.Mechanical.Duct;
                var ductType = duct?.DuctType;
                var familyName = ductType?.FamilyName ?? string.Empty;
                
                _log($"[DEBUG] Duct {element.Id} family name: {familyName}");
                
                // Check family name for shape indicators
                if (familyName.Contains("Round", StringComparison.OrdinalIgnoreCase))
                {
                    return "Round";
                }
                else if (familyName.Contains("Rectangular", StringComparison.OrdinalIgnoreCase))
                {
                    return "Rectangular";
                }
                
                // Fallback: Check by dimensions (if width ≈ height, likely round)
                var (width, height) = GetMepElementDimensions(element);
                double widthMm = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                double heightMm = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                bool isRoundByDimensions = Math.Abs(widthMm - heightMm) < 10.0;
                
                return isRoundByDimensions ? "Round" : "Rectangular";
            }
            catch (Exception ex)
            {
                _log($"Error getting duct shape: {ex.Message}");
                return "Rectangular"; // Safe default
            }
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Check if a sleeve exists at the EXACT placement point stored in XML
        /// This is essential for refresh to detect deleted sleeves and reset IsResolved flag
        /// Without this, deleted sleeves cannot be re-placed (duplication suppressor prevents it)
        /// </summary>
        private bool CheckForExistingSleeve(XYZ placementPoint, Document document)
        {
            try
            {
                var openingFamilies = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .ToList();

                _log($"[CheckExistingSleeve] Checking for sleeves at EXACT placement point {placementPoint}");

                foreach (var opening in openingFamilies)
                {
                    XYZ openingLocation = null;
                    
                    // Handle both LocationPoint and LocationCurve
                    if (opening.Location is LocationPoint locationPoint)
                    {
                        openingLocation = locationPoint.Point;
                    }
                    else if (opening.Location is LocationCurve locationCurve)
                    {
                        // For LocationCurve, use the midpoint
                        openingLocation = locationCurve.Curve.Evaluate(0.5, true);
                    }
                    
                    if (openingLocation != null)
                    {
                        // Check for EXACT match (within Revit's precision tolerance ~1mm)
                        double distance = openingLocation.DistanceTo(placementPoint);
                        if (distance < 0.003) // ~1mm tolerance for Revit precision
                        {
                            _log($"[CheckExistingSleeve] ✓ Found existing sleeve {opening.Id} at EXACT placement point");
                            return true;
                        }
                    }
                }
                
                _log($"[CheckExistingSleeve] No existing sleeve found at EXACT placement point {placementPoint}");
                return false;
            }
            catch (Exception ex)
            {
                _log($"[CheckExistingSleeve] Error: {ex.Message}");
                return false; // Assume no sleeve if error
            }
        }
        
        /// <summary>
        /// Get element thickness for walls, floors, and structural framing
        /// </summary>
        private double GetElementThickness(Element element)
        {
            try
            {
                if (element is Wall wall)
                {
                    return wall.Width;
                }
                else if (element is Floor floor)
                {
                    return floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.1;
                }
                else if ((element?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming))
                {
                    // Structural framing: read TYPE parameter 'b' (case-insensitive) regardless of instance/type wrapper
                    try
                    {
                        // Resolve the type element from any element (FamilyInstance or not)
                        ElementId typeId = ElementId.InvalidElementId;
                        try { typeId = (element as FamilyInstance)?.GetTypeId() ?? element.GetTypeId(); } catch { }
                        var typeElem = element.Document?.GetElement(typeId);

                        if (typeElem == null)
                        {
                            DebugLogger.Warning($"[FRAMING-THICKNESS] Could not get type element for framing {element.Id.IntegerValue}");
                            return 0.1;
                        }

                        Parameter p = null;
                        double bVal = 0.0;

                        // Try multiple parameter names for breadth/depth
                        string[] possibleNames = { "b", "B", "Breadth", "Depth", "Width", "Height", "d", "D" };

                        foreach (var paramName in possibleNames)
                        {
                            try
                            {
                                p = typeElem.LookupParameter(paramName);
                                if (p != null && !p.IsReadOnly)
                                {
                                    bVal = p.AsDouble();
                                    if (bVal > 0)
                                    {
                                        DebugLogger.Info($"[FRAMING-THICKNESS] Found parameter '{paramName}' = {bVal:F6}ft on framing {element.Id.IntegerValue}");
                                        break;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Warning($"[FRAMING-THICKNESS] Error reading parameter '{paramName}': {ex.Message}");
                            }
                        }

                        // Log all available parameters for debugging
                        if (bVal <= 0.0)
                        {
                            try
                            {
                                var parts = new System.Collections.Generic.List<string>();
                                foreach (Parameter tp in typeElem.Parameters)
                                {
                                    var name = tp?.Definition?.Name ?? "<null>";
                                    string val = string.Empty;
                                    if (tp.StorageType == StorageType.Double)
                                    {
                                        double d = tp.AsDouble();
                                        double mm = UnitUtils.ConvertFromInternalUnits(d, UnitTypeId.Millimeters);
                                        val = mm.ToString("F1") + "mm";
                                    }
                                    else
                                    {
                                        val = tp.AsString() ?? tp.AsValueString() ?? string.Empty;
                                    }
                                    parts.Add(name + ":" + val);
                                }
                                DebugLogger.Info($"[FRAMING-THICKNESS-PARAMS] typeId={typeId.IntegerValue}: {string.Join(", ", parts)}");
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Warning($"[FRAMING-THICKNESS] Error logging parameters: {ex.Message}");
                            }
                        }

                        // Convert to mm for logging
                        try
                        {
                            var valMm = UnitUtils.ConvertFromInternalUnits(bVal, UnitTypeId.Millimeters);
                            DebugLogger.Info($"[FRAMING-THICKNESS] id={element.Id.IntegerValue}: key={(p?.Definition?.Name ?? "<null>")} value={valMm:F1}mm");
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[FRAMING-THICKNESS] Error converting units: {ex.Message}");
                        }

                        return bVal > 0.0 ? bVal : 0.1;
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[FRAMING-THICKNESS] Error getting framing thickness: {ex.Message}");
                        return 0.1;
                    }
                }
                return 0.1; // Default fallback
            }
            catch
            {
                return 0.1; // Default fallback
            }
        }
        
        /// <summary>
        /// Format MEP element size as string (e.g., "600x300", "Ø200")
        /// </summary>
        private string FormatMepElementSize(Element mepElement, double width, double height, string shape)
        {
            try
            {
                if (shape == "Round" || shape == "Circular")
                {
                    // Convert from feet to mm and format as "Ø200"
                    double diameterMm = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                    return $"Ø{Math.Round(diameterMm, 0)}";
                }
                else
                {
                    // Convert from feet to mm and format as "600x300"
                    double widthMm = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                    double heightMm = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                    return $"{Math.Round(widthMm, 0)}x{Math.Round(heightMm, 0)}";
                }
            }
            catch
            {
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get MEP system abbreviation (e.g., "SA", "RA", "EX")
        /// </summary>
        private string GetMepSystemAbbreviation(Element mepElement)
        {
            try
            {
                // Try to get system abbreviation parameter
                var abbrevParam = mepElement.LookupParameter("System Abbreviation");
                if (abbrevParam != null && abbrevParam.StorageType == StorageType.String)
                {
                    return abbrevParam.AsString() ?? string.Empty;
                }
                
                // Fallback: try to get system name and abbreviate it
                var systemNameParam = mepElement.LookupParameter("System Name");
                if (systemNameParam != null && systemNameParam.StorageType == StorageType.String)
                {
                    var systemName = systemNameParam.AsString();
                    if (!string.IsNullOrEmpty(systemName))
                    {
                        // Take first 2-3 characters as abbreviation
                        return systemName.Length > 3 ? systemName.Substring(0, 3).ToUpper() : systemName.ToUpper();
                    }
                }
                
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Get damper connector info (MSFD type and connector side)
        /// Returns (isMSFD, connectorSide) where connectorSide is "Left", "Right", "Top", "Bottom", or empty string
        /// </summary>
        private (bool isMSFD, string connectorSide) GetDamperConnectorInfo(Element mepElement, string mepCategory)
        {
            try
            {
                // Only process duct accessories (dampers)
                if (mepCategory != "Duct Accessories")
                {
                    return (false, string.Empty);
                }
                
                var damper = mepElement as FamilyInstance;
                if (damper == null)
                {
                    return (false, string.Empty);
                }
                
                // Check if MSFD type by checking family type name
                string familyTypeName = damper.Symbol?.Name ?? "";
                bool isMSFD = familyTypeName.Trim().ToUpperInvariant().Contains("MSFD");
                
                // Get connector side using FireDamperSleevePlacerService logic
                string connectorSide = string.Empty;
                if (isMSFD)
                {
                    connectorSide = FireDamperSleevePlacerService.GetConnectorSideWorld(damper, out _);
                    _log($"[DEBUG] MSFD Damper {damper.Id}: Connector side = {connectorSide}");
                }
                
                return (isMSFD, connectorSide);
            }
            catch (Exception ex)
            {
                _log($"[DEBUG] Error getting damper connector info: {ex.Message}");
                return (false, string.Empty);
            }
        }
        
        /// <summary>
        /// Get structural element normal/direction vector for orientation calculation
        /// For walls: returns wall normal (perpendicular to wall direction)
        /// For floors: returns null (use MEP orientation instead)
        /// For framing: returns framing direction
        /// </summary>
        private XYZ GetStructuralElementNormal(Element element)
        {
            try
            {
                if (element is Wall wall)
                {
                    // Get wall normal (perpendicular to wall direction)
                    var locationCurve = wall.Location as LocationCurve;
                    if (locationCurve != null)
                    {
                        var curve = locationCurve.Curve as Line;
                        if (curve != null)
                        {
                            var wallDirection = curve.Direction;
                            // FIX: For X-wall (Direction=(1,0,0)), normal should be (0,1,0) or (0,-1,0)
                            // For Y-wall (Direction=(0,1,0)), normal should be (1,0,0) or (-1,0,0)
                            // The original formula was correct: normal = (-Y, X, 0)
                            var wallNormal = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                            
                            // DEBUG: Log wall direction and normal
                            DebugLogger.Info($"[WALL-DIR] Wall {wall.Id.IntegerValue}: Direction=({wallDirection.X:F3},{wallDirection.Y:F3},{wallDirection.Z:F3}), Normal=({wallNormal.X:F3},{wallNormal.Y:F3},{wallNormal.Z:F3})");
                            
                            // ALSO log to placement_debug.log for immediate visibility
                            try
                            {
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                    $"[WALL-DIR] Wall {wall.Id.IntegerValue}: Direction=({wallDirection.X:F3},{wallDirection.Y:F3},{wallDirection.Z:F3}), Normal=({wallNormal.X:F3},{wallNormal.Y:F3},{wallNormal.Z:F3})\n");
                            }
                            catch { }
                            
                            return wallNormal;
                        }
                    }
                }
                else if (element is Floor)
                {
                    // For floors, we don't need the normal - we use MEP orientation instead
                    return null;
                }
                else if (element is FamilyInstance famInst && 
                         famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For framing, get direction vector
                    var locationCurve = famInst.Location as LocationCurve;
                    if (locationCurve != null)
                    {
                        var curve = locationCurve.Curve as Line;
                        if (curve != null)
                        {
                            return curve.Direction;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Get element normal vector for walls, floors, and structural framing
        /// </summary>
        private XYZ GetElementNormal(Element element)
        {
            try
            {
                if (element is Wall wall)
                {
                    if (wall.Location is LocationCurve curve)
                    {
                        var line = curve.Curve as Line;
                        if (line != null)
                        {
                            var wallDirection = line.Direction;
                            // Wall normal is perpendicular to wall curve direction in horizontal plane
                            return new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                        }
                    }
                    return XYZ.BasisX; // Default fallback
                }
                else if (element is Floor)
                {
                    // For floors, normal is typically upward (Z direction)
                    return XYZ.BasisZ;
                }
                else if (element is FamilyInstance famInst && 
                         famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For structural framing, use default upward direction
                    return XYZ.BasisZ;
                }
                return XYZ.BasisX; // Default fallback
            }
            catch
            {
                return XYZ.BasisX; // Default fallback
            }
        }
        
        /// <summary>
        /// Get MEP element dimensions (width and height)
        /// </summary>
        private (double width, double height) GetMepElementDimensions(Element mepElement)
        {
            try
            {
                if (mepElement is Duct duct)
                {
                    var width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                    var height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                    
                    // Handle round ducts
                    var diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                    if ((width <= 0.0 || height <= 0.0) && diameter > 0.0)
                    {
                        width = diameter;
                        height = diameter;
                    }
                    
                    return (width, height);
                }
                else if (mepElement is Pipe pipe)
                {
                    var diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0;
                    return (diameter, diameter); // Pipes are round
                }
                else if (mepElement is Autodesk.Revit.DB.Electrical.CableTray cableTray)
                {
                    var width = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)?.AsDouble() ?? 0;
                    var height = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? 0;
                    return (width, height);
                }
                else if (mepElement is Conduit conduit)
                {
                    var width = conduit.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)?.AsDouble() ?? 0;
                    var height = conduit.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? 0;
                    return (width, height);
                }
                else if (mepElement is FamilyInstance famInst)
                {
                    // Handle duct accessories (dampers), cable trays, etc.
                    // Try damper-specific parameters first
                    var widthParam = famInst.LookupParameter("Damper Width") ?? 
                                    famInst.LookupParameter("Width") ?? 
                                    famInst.LookupParameter("width");
                    var heightParam = famInst.LookupParameter("Damper Height") ?? 
                                     famInst.LookupParameter("Height") ?? 
                                     famInst.LookupParameter("height");
                    
                    double width = widthParam?.AsDouble() ?? 0.1;
                    double height = heightParam?.AsDouble() ?? 0.1;
                    
                    return (width, height);
                }
                
                return (0.1, 0.1); // Default fallback
            }
            catch
            {
                return (0.1, 0.1); // Default fallback
            }
        }
        
        /// <summary>
        /// FOOLPROOF: Auto-detect dampers even if user forgot to select "Duct Accessories" category
        /// This prevents the scenario where user only selects "Ducts" and misses dampers
        /// </summary>
        private List<(Element, Element, BoundingBoxXYZ, XYZ)> AutoDetectMissingDampers(
            Document document, List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections)
        {
            var enhancedIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>(currentIntersections);
            
            try
            {
                // Check if user selected ducts but forgot duct accessories
                var hasDucts = currentIntersections.Any(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase);
                });
                
                var hasDuctAccessories = currentIntersections.Any(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return string.Equals(mepCat, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                });
                
                if (hasDucts && !hasDuctAccessories)
                {
                    _log($"[FOOLPROOF] User selected ducts but forgot duct accessories - auto-detecting dampers");
                    
                    // Get all structural elements that ducts intersect with
                    var ductWallIds = currentIntersections
                        .Where(i => string.Equals(GetElementCategoryName(i.Item1), "Ducts", StringComparison.OrdinalIgnoreCase))
                        .Select(i => i.Item2.Id)
                        .Distinct()
                        .ToList();
                    
                    // Find dampers that intersect with the same walls
                    var autoDetectedDampers = FindDampersIntersectingSameWalls(document, ductWallIds);
                    
                    foreach (var (damperElement, wallElement, boundingBox, intersectionPoint) in autoDetectedDampers)
                    {
                        // Avoid duplicates
                        if (!enhancedIntersections.Any(i => i.Item1.Id == damperElement.Id))
                        {
                            enhancedIntersections.Add((damperElement, wallElement, boundingBox, intersectionPoint));
                            _log($"[FOOLPROOF] Auto-detected damper {damperElement.Id} intersecting wall {wallElement.Id}");
                        }
                    }
                    
                    _log($"[FOOLPROOF] Auto-detected {autoDetectedDampers.Count} dampers that user missed");
                }
                else if (hasDuctAccessories)
                {
                    _log($"[FOOLPROOF] User correctly selected duct accessories - no auto-detection needed");
                }
                else
                {
                    _log($"[FOOLPROOF] No ducts selected - no auto-detection needed");
                }
            }
            catch (Exception ex)
            {
                _log($"Error in auto-detection: {ex.Message}");
            }
            
            return enhancedIntersections;
        }
        
        /// <summary>
        /// Find dampers that intersect with the same walls as ducts
        /// </summary>
        private List<(Element, Element, BoundingBoxXYZ, XYZ)> FindDampersIntersectingSameWalls(
            Document document, List<ElementId> wallIds)
        {
            var damperIntersections = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            try
            {
                // Get all duct accessories (dampers) from the document
                var ductAccessories = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .Cast<Element>()
                    .ToList();
                
                // Also check family instances with "damper" in the name
                var damperFamilyInstances = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.ToLowerInvariant().Contains("damper") == true)
                    .Cast<Element>()
                    .ToList();
                
                var allDampers = ductAccessories.Concat(damperFamilyInstances).ToList();
                
                foreach (var damper in allDampers)
                {
                    var damperBbox = damper.get_BoundingBox(null);
                    if (damperBbox == null) continue;
                    
                    // Check if damper intersects with any of the walls that ducts intersect
                    foreach (var wallId in wallIds)
                    {
                        var wallElement = document.GetElement(wallId);
                        if (wallElement == null) continue;
                        
                        var wallBbox = wallElement.get_BoundingBox(null);
                        if (wallBbox == null) continue;
                        
                        // Check if bounding boxes intersect
                        if (BoundingBoxesIntersect(damperBbox.Min, damperBbox.Max, wallBbox.Min, wallBbox.Max))
                        {
                            // Calculate intersection point (center of intersection)
                            var intersectionMin = new XYZ(
                                Math.Max(damperBbox.Min.X, wallBbox.Min.X),
                                Math.Max(damperBbox.Min.Y, wallBbox.Min.Y),
                                Math.Max(damperBbox.Min.Z, wallBbox.Min.Z)
                            );
                            
                            var intersectionMax = new XYZ(
                                Math.Min(damperBbox.Max.X, wallBbox.Max.X),
                                Math.Min(damperBbox.Max.Y, wallBbox.Max.Y),
                                Math.Min(damperBbox.Max.Z, wallBbox.Max.Z)
                            );
                            
                            var intersectionPoint = new XYZ(
                                (intersectionMin.X + intersectionMax.X) / 2,
                                (intersectionMin.Y + intersectionMax.Y) / 2,
                                (intersectionMin.Z + intersectionMax.Z) / 2
                            );
                            
                            var intersectionBbox = new BoundingBoxXYZ
                            {
                                Min = intersectionMin,
                                Max = intersectionMax
                            };
                            
                            damperIntersections.Add((damper, wallElement, intersectionBbox, intersectionPoint));
                            break; // Found intersection with this wall, move to next damper
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log($"Error finding dampers intersecting same walls: {ex.Message}");
            }
            
            return damperIntersections;
        }
        
        /// <summary>
        /// Check if two bounding boxes intersect
        /// </summary>
        private bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            return min1.X <= max2.X && max1.X >= min2.X &&
                   min1.Y <= max2.Y && max1.Y >= min2.Y &&
                   min1.Z <= max2.Z && max1.Z >= min2.Z;
        }

        /// <summary>
        /// FOOLPROOF METHOD: Prioritize intersections by category to ensure dampers are always processed first
        /// Priority Order: 1) Dampers, 2) Other Duct Accessories, 3) Ducts, 4) Everything else
        /// </summary>
        private List<(Element, Element, BoundingBoxXYZ, XYZ)> PrioritizeIntersectionsByCategory(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> intersections)
        {
            var prioritized = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            try
            {
                // Priority 1: Dampers (highest priority)
                var dampers = intersections.Where(i => IsDamperElement(i.Item1)).ToList();
                prioritized.AddRange(dampers);
                _log($"[PRIORITY] Added {dampers.Count} dampers (Priority 1)");
                
                // Priority 2: Other Duct Accessories (medium priority)
                var otherDuctAccessories = intersections.Where(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return string.Equals(mepCat, "Duct Accessories", StringComparison.OrdinalIgnoreCase) && 
                           !IsDamperElement(i.Item1);
                }).ToList();
                prioritized.AddRange(otherDuctAccessories);
                _log($"[PRIORITY] Added {otherDuctAccessories.Count} other duct accessories (Priority 2)");
                
                // Priority 3: Ducts (lower priority - will be skipped if damper nearby)
                var ducts = intersections.Where(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase);
                }).ToList();
                prioritized.AddRange(ducts);
                _log($"[PRIORITY] Added {ducts.Count} ducts (Priority 3)");
                
                // Priority 4: Everything else (lowest priority)
                var others = intersections.Where(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    return !string.Equals(mepCat, "Duct Accessories", StringComparison.OrdinalIgnoreCase) &&
                           !string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase);
                }).ToList();
                prioritized.AddRange(others);
                _log($"[PRIORITY] Added {others.Count} other MEP elements (Priority 4)");
                
                _log($"[PRIORITY] Total prioritized intersections: {prioritized.Count} (Original: {intersections.Count})");
            }
            catch (Exception ex)
            {
                _log($"Error prioritizing intersections: {ex.Message}");
                return intersections; // Fallback to original order
            }
            
            return prioritized;
        }

        /// <summary>
        /// Pre-calculate damper locations from XML + current intersections (Calculate Once, Use Many Times)
        /// This solves the cross-XML cycle problem where dampers and ducts are processed separately
        /// </summary>
        private List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> PreCalculateDamperLocationsFromXmlAndCurrent(
            Document document, List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections)
        {
            var damperLocations = new List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)>();
            
            try
            {
                // STEP 1: Get dampers from saved XML data (from previous refresh cycles)
                if (_clashZoneStorage?.ClashZones != null)
                {
                    foreach (var clashZone in _clashZoneStorage.ClashZones)
                    {
                        if (IsDamperClashZone(clashZone))
                        {
                            // Try to get the damper element from document
                            var damperElement = document.GetElement(clashZone.MepElementId);
                            if (damperElement != null)
                            {
                                var damperBbox = damperElement.get_BoundingBox(null);
                                if (damperBbox != null)
                                {
                                    damperLocations.Add((damperId: clashZone.MepElementId, bbox: damperBbox, wallId: clashZone.StructuralElementId));
                                    _log($"[PRE-CALC-XML] Damper {clashZone.MepElementId} location cached from XML");
                                }
                            }
                        }
                    }
                }
                
                // STEP 2: Add dampers from current intersections (from current refresh cycle)
                foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in currentIntersections)
                {
                    var mepCat = GetElementCategoryName(mepElement);
                    
                    // Check if it's a damper
                    if (string.Equals(mepCat, "Duct Accessories", StringComparison.OrdinalIgnoreCase) ||
                        IsDamperElement(mepElement))
                    {
                        var damperBbox = mepElement.get_BoundingBox(null);
                        if (damperBbox != null)
                        {
                            // Avoid duplicates (check if already added from XML)
                            if (!damperLocations.Any(d => d.damperId == mepElement.Id))
                            {
                                damperLocations.Add((damperId: mepElement.Id, bbox: damperBbox, wallId: structuralElement.Id));
                                _log($"[PRE-CALC-CURRENT] Damper {mepElement.Id} location cached from current intersections");
                            }
                        }
                    }
                }
                
                _log($"[PRE-CALC] Total damper locations: {damperLocations.Count} (XML: {_clashZoneStorage?.ClashZones?.Count(c => IsDamperClashZone(c)) ?? 0}, Current: {currentIntersections.Count(i => IsDamperElement(i.Item1))})");
            }
            catch (Exception ex)
            {
                _log($"Error pre-calculating damper locations: {ex.Message}");
            }
            
            return damperLocations;
        }
        
        /// <summary>
        /// Check if a clash zone represents a damper
        /// </summary>
        private bool IsDamperClashZone(ClashZone clashZone)
        {
            try
            {
                // First check if it's a duct accessory
                if (!string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                    return false;
                
                // Then check if it's specifically a damper by looking at the system abbreviation or other indicators
                // Dampers typically have system abbreviations like "FD", "MSFD", "MSD", "MD", etc.
                var systemAbbr = clashZone.MepElementSystemAbbreviation?.ToUpperInvariant() ?? "";
                
                // Check for damper-related system abbreviations
                if (systemAbbr.Contains("FD") || systemAbbr.Contains("MSFD") || systemAbbr.Contains("MSD") || 
                    systemAbbr.Contains("MD") || systemAbbr.Contains("DAMPER"))
                {
                    return true;
                }
                
                // Also check if it's marked as MSFD damper
                if (clashZone.IsMSFDDamper)
                {
                    return true;
                }
                
                // Fallback: check if the formatted size suggests it's a damper
                // Dampers typically have rectangular sizes (e.g., "400x200") rather than round sizes
                var formattedSize = clashZone.MepElementFormattedSize?.ToUpperInvariant() ?? "";
                if (formattedSize.Contains("X") && !formattedSize.Contains("Ø"))
                {
                    // It's rectangular, likely a damper
                    return true;
                }
                
                // If we can't determine, be conservative and exclude it
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Check if a duct is near a damper using pre-calculated locations (Efficient O(n) lookup)
        /// </summary>
        private bool IsDuctNearDamper(Element ductElement, List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations)
        {
            try
            {
                const double proximityTolerance = 0.02; // 1/4 inch tolerance (as requested)
                
                // Get duct bounding box
                var ductBbox = ductElement.get_BoundingBox(null);
                if (ductBbox == null) return false;
                
                // Check pre-calculated damper locations (Efficient O(n) lookup)
                foreach (var (damperId, damperBbox, wallId) in damperLocations)
                {
                    // Skip if it's the same element
                    if (damperId == ductElement.Id) continue;
                    
                    // Check if duct and damper are within proximity tolerance
                    var distance = GetMinimumDistanceBetweenBoundingBoxes(ductBbox, damperBbox);
                    
                    if (distance <= proximityTolerance)
                    {
                        _log($"[DUCT-DAMPER] Duct {ductElement.Id} is {distance:F4}ft from Damper {damperId} (tolerance: {proximityTolerance}ft = 1/4\")");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _log($"Error checking duct-damper proximity: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Check if an element is a damper based on family name or category
        /// </summary>
        private bool IsDamperElement(Element element)
        {
            try
            {
                // First check if it's a duct accessory
                if (element.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                {
                    // Then check if it's specifically a damper family
                    if (element is FamilyInstance famInst)
                    {
                        var familyName = famInst.Symbol?.Family?.Name?.ToLowerInvariant();
                        if (familyName != null && familyName.Contains("damper"))
                            return true;
                    }
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Calculate minimum distance between two bounding boxes
        /// </summary>
        private double GetMinimumDistanceBetweenBoundingBoxes(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            try
            {
                // Calculate distance between closest corners
                var min1 = bbox1.Min;
                var max1 = bbox1.Max;
                var min2 = bbox2.Min;
                var max2 = bbox2.Max;
                
                // Find closest points
                var closest1 = new XYZ(
                    Math.Max(min1.X, Math.Min(max1.X, (min2.X + max2.X) / 2)),
                    Math.Max(min1.Y, Math.Min(max1.Y, (min2.Y + max2.Y) / 2)),
                    Math.Max(min1.Z, Math.Min(max1.Z, (min2.Z + max2.Z) / 2))
                );
                
                var closest2 = new XYZ(
                    Math.Max(min2.X, Math.Min(max2.X, (min1.X + max1.X) / 2)),
                    Math.Max(min2.Y, Math.Min(max2.Y, (min1.Y + max1.Y) / 2)),
                    Math.Max(min2.Z, Math.Min(max2.Z, (min1.Z + max1.Z) / 2))
                );
                
                return closest1.DistanceTo(closest2);
            }
            catch
            {
                return double.MaxValue; // Return large distance if calculation fails
            }
        }

        /// <summary>
        /// Get MEP element orientation from bounding box (X or Y)
        /// </summary>
        private string GetMepElementOrientationFromBbox(Element mepElement)
        {
            try
            {
                // Get MEP element's bounding box
                BoundingBoxXYZ mepBbox = mepElement.get_BoundingBox(null);
                
                if (mepBbox == null)
                {
                    DebugLogger.Warning($"[GetMepElementOrientationFromBbox] No bounding box for element {mepElement.Id}");
                    return "X"; // Default to X orientation
                }
                
                // Calculate dimensions
                double bboxWidth = mepBbox.Max.X - mepBbox.Min.X;  // X-axis dimension
                double bboxHeight = mepBbox.Max.Y - mepBbox.Min.Y; // Y-axis dimension
                
                // Determine orientation
                string mepOrientation;
                if (bboxWidth > bboxHeight)
                {
                    mepOrientation = "X"; // Width runs in X-axis
                }
                else
                {
                    mepOrientation = "Y"; // Width runs in Y-axis
                }
                
                DebugLogger.Info($"[GetMepElementOrientationFromBbox] Element {mepElement.Id}: BboxWidth={bboxWidth:F3}, BboxHeight={bboxHeight:F3}, Orientation={mepOrientation}");
                
                return mepOrientation;
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[GetMepElementOrientationFromBbox] Error getting orientation for element {mepElement?.Id}: {ex.Message}");
                return "X"; // Default to X orientation
            }
        }

        /// <summary>
        /// Get MEP element orientation vector
        /// </summary>
        private XYZ GetMepElementOrientation(Element mepElement)
        {
            try
            {
                DebugLogger.Info($"[GetMepElementOrientation] Element {mepElement.Id}: Type={mepElement.GetType().Name}, Category={mepElement.Category?.Name}, Location={mepElement.Location?.GetType().Name}");
                
                if (mepElement is Duct duct && duct.Location is LocationCurve curve)
                {
                    var line = curve.Curve as Line;
                    if (line != null)
                    {
                        var direction = line.Direction;
                        DebugLogger.Info($"[GetMepElementOrientation] Duct {mepElement.Id}: Direction=({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3})");
                        return direction;
                    }
                }
                else if (mepElement is Duct verticalDuct && verticalDuct.Location is LocationPoint point)
                {
                    // Vertical ducts might have LocationPoint instead of LocationCurve
                    DebugLogger.Info($"[GetMepElementOrientation] Duct {mepElement.Id}: Has LocationPoint, checking for vertical orientation");
                    // For vertical ducts, we might need to check other properties
                    return XYZ.BasisZ; // Default to vertical for now
                }
                else if (mepElement is FamilyInstance ductAccessory && 
                         ductAccessory.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory &&
                         ductAccessory.Location is LocationCurve ductAccessoryCurve)
                {
                    var line = ductAccessoryCurve.Curve as Line;
                    if (line != null)
                    {
                        var direction = line.Direction;
                        DebugLogger.Info($"[GetMepElementOrientation] DuctAccessory {mepElement.Id}: Direction=({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3})");
                        return direction;
                    }
                }
                else if (mepElement is Pipe pipe && pipe.Location is LocationCurve pipeCurve)
                {
                    var line = pipeCurve.Curve as Line;
                    if (line != null)
                    {
                        return line.Direction;
                    }
                }
                else if (mepElement is Conduit conduit && conduit.Location is LocationCurve conduitCurve)
                {
                    var line = conduitCurve.Curve as Line;
                    if (line != null)
                    {
                        return line.Direction;
                    }
                }
                else if (mepElement is Autodesk.Revit.DB.Electrical.CableTray cableTray && cableTray.Location is LocationCurve cableTrayCurve)
                {
                    var line = cableTrayCurve.Curve as Line;
                    if (line != null)
                    {
                        return line.Direction;
                    }
                }
                else if (mepElement.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
                {
                    // ✅ FIX: Handle dampers (FamilyInstance) - get orientation from transform
                    var damper = mepElement as FamilyInstance;
                    if (damper != null)
                    {
                        var transform = damper.GetTotalTransform();
                        // Use BasisX as the "flow direction" (similar to ducts/pipes)
                        var damperDirection = transform.BasisX;
                        DebugLogger.Info($"[GetMepElementOrientation] Damper {mepElement.Id}: Direction=({damperDirection.X:F3}, {damperDirection.Y:F3}, {damperDirection.Z:F3})");
                        return damperDirection;
                    }
                }

                return XYZ.BasisX; // Default fallback
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[GetMepElementOrientation] Error getting orientation for element {mepElement?.Id}: {ex.Message}");
                return XYZ.BasisX; // Default fallback
            }
        }
        
        /// <summary>
        /// Get pipe opening type (Circular or Rectangular)
        /// </summary>
        private string GetPipeOpeningType(Element mepElement)
        {
            try
            {
                if (mepElement is Pipe)
                {
                    // TODO: Get from UI settings or element properties
                    // For now, default to Circular
                    return "Circular";
                }
                return string.Empty; // Not a pipe
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Get MEP element level information for sleeve placement
        /// </summary>
        private (string levelName, double levelElevation) GetMepElementLevelInfo(Element mepElement)
        {
            try
            {
                // Try to get level from MEP element's LevelId
                if (mepElement.LevelId != ElementId.InvalidElementId)
                {
                    var level = mepElement.Document.GetElement(mepElement.LevelId) as Level;
                    if (level != null)
                    {
                        return (level.Name, level.Elevation);
                    }
                }
                
                // Fallback: try to get level from location
                if (mepElement.Location is LocationPoint locationPoint)
                {
                    var elevation = locationPoint.Point.Z;
                    return ($"Auto-Level-{elevation:F2}", elevation);
                }
                else if (mepElement.Location is LocationCurve locationCurve)
                {
                    var startPoint = locationCurve.Curve.GetEndPoint(0);
                    var elevation = startPoint.Z;
                    return ($"Auto-Level-{elevation:F2}", elevation);
                }
                
                // Final fallback
                return ("Level 1", 0.0);
            }
            catch (Exception ex)
            {
                _log($"[ClashZoneService] Error getting MEP element level info: {ex.Message}");
                return ("Level 1", 0.0);
            }
        }
        
        /// <summary>
        /// Check for existing sleeve at placement point
        /// </summary>
        private bool CheckForExistingSleeveAtPlacementPoint(XYZ placementPoint, Document document, double tolerance = 0.05) // 50mm tolerance
        {
            try
            {
                // ACTUAL CHECK: Look for existing sleeves near the placement point
                // This prevents the IsResolved flag from staying true when sleeves are deleted
                
                // Get all opening families in the document
                var openingFamilies = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .ToList();
                
                // Check if any opening exists within tolerance of the placement point
                foreach (var opening in openingFamilies)
                {
                    if (opening.Location is LocationPoint locationPoint)
                    {
                        double distance = locationPoint.Point.DistanceTo(placementPoint);
                        if (distance <= tolerance)
                        {
                            _log($"[CheckForExistingSleeve] Found existing opening at distance {distance:F3} from placement point");
                            return true;
                        }
                    }
                }
                
                _log($"[CheckForExistingSleeve] No existing sleeve found at placement point {placementPoint}");
                return false;
            }
            catch (Exception ex)
            {
                _log($"[CheckForExistingSleeve] Error checking for existing sleeve: {ex.Message}");
                return false;
            }
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
            // Return clearance in feet (Revit internal units) for consistency
            // Convert 50mm to feet
            return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
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
                
                // Convert clearance from mm to feet (Revit internal units)
                // UI stores values in mm, but Revit uses feet internally
                return UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
            }
            catch (Exception ex)
            {
                _log($"Error calculating clearance: {ex.Message}, using default");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // Fallback to 50mm converted to feet
            }
        }
        
        private bool IsPointNear(XYZ point1, XYZ point2, double tolerance)
        {
            return point1.DistanceTo(point2) <= tolerance;
        }
        
        #endregion
    }
}
