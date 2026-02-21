using System;
using System.Globalization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using static JSE_RevitAddin_MEP_OPENINGS.Models.MepCategoryConstants;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.DamperDetection;
using JSE_RevitAddin_MEP_OPENINGS.Services.InsulationDetection;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing clash zones and their persistence
    /// </summary>
    public class ClashZoneService
    {
        private readonly ClashZoneStorage? _clashZoneStorage;
        private readonly Action<string>? _log;
        
        // ⚠️ CRITICAL: Memory manager for timeout/memory limit protection (can be null if not provided)
        private MemoryManager? _memoryManager;
        
        // ✅ MEMORY PROFILING: Profiler for tracking actual memory usage per clash zone
        private MemoryProfiler? _memoryProfiler;
        
        // ✅ PERFORMANCE MONITOR: Track sub-operations during clash zone detection
        private PerformanceMonitor? _performanceMonitor;
        
        // ✅ OOP REFACTORING: Optional FlagManager for centralized flag operations
        private readonly Services.Interfaces.Refactor.IFlagManager? _flagManager;
        
        // ✅ OOP REFACTORING: Optional GuidManager for centralized GUID operations
        private readonly GuidManager? _guidManager;

        // PERF OPT: Single shared InsulationDetector — stateless, safe to reuse across all CreateClashZone calls
        private readonly InsulationDetector _insulationDetector = new InsulationDetector();

        // PERF OPT: Shared strategy instances — all are stateless (no mutable fields), safe to reuse
        private readonly ISleevePlacementStrategy _ductStrategy = new Strategies.DuctPlacementStrategy();
        private readonly ISleevePlacementStrategy _pipeStrategy = new Strategies.PipePlacementStrategy();
        private readonly ISleevePlacementStrategy _cableTrayStrategy = new Strategies.CableTrayPlacementStrategy();
        // DamperPlacementStrategy holds a Document reference — cache by document to avoid re-allocation
        private Strategies.DamperPlacementStrategy _cachedDamperStrategy;
        private Document _cachedDamperStrategyDoc;

        // ✅ PARALLEL PROCESSING: Threshold for enabling parallel zone creation
        // Below this threshold: single-threaded (lower overhead)
        // Above this threshold: parallel processing (better utilization)
        private const int PARALLEL_THRESHOLD = 100;

        public ClashZoneService(ClashZoneStorage? clashZoneStorage, Action<string>? log, Services.Interfaces.Refactor.IFlagManager? flagManager = null, GuidManager? guidManager = null)
        {
            _clashZoneStorage = clashZoneStorage;
            _log = log;
            _flagManager = flagManager; // Optional dependency for backward compatibility
            _guidManager = guidManager; // Optional dependency for backward compatibility
        }

        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZATION: Wrap slow disk logging in a diagnostic guard.
        /// This method is called in tight loops and avoids hundreds of slow SafeFileLogger calls.
        /// </summary>
        private void SafeLog(string filename, string message)
        {
            if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText(filename, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
            }
        }
        
        /// <summary>
        /// Set memory manager for timeout/memory protection during clash zone detection
        /// </summary>
        public void SetMemoryManager(MemoryManager memoryManager)
        {
            _memoryManager = memoryManager;
        }
        
        /// <summary>
        /// Set memory profiler for tracking actual vs theoretical memory usage
        /// </summary>
        public void SetMemoryProfiler(MemoryProfiler memoryProfiler)
        {
            _memoryProfiler = memoryProfiler;
        }

        /// <summary>
        /// Set performance monitor for detailed sub-timing of clash zone detection
        /// </summary>
        public void SetPerformanceMonitor(PerformanceMonitor performanceMonitor)
        {
            _performanceMonitor = performanceMonitor;
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
                    var mepElement = (clashZone.MepElementId != null && clashZone.MepElementId.GetIntegerValue() != -1) ? document.GetElement(clashZone.MepElementId) : null;
                    var structuralElement = (clashZone.StructuralElementId != null && clashZone.StructuralElementId.GetIntegerValue() != -1) ? document.GetElement(clashZone.StructuralElementId) : null;
                    
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

            _log($"Returning all {_clashZoneStorage.ClashZones.Count} clash zones (filtering already done in IntersectionDetectionService)");
            
            // ✅ SIMPLIFIED: Return all clash zones since filtering is already done in IntersectionDetectionService
            // The 5-step filtering (Section Box, Reference File, MEP Categories, Host File, Host Categories)
            // is already applied during element collection in IntersectionDetectionService
            
            // Update storage metadata
            _clashZoneStorage.LastUpdated = DateTime.Now;
            
            return _clashZoneStorage.ClashZones.ToList();
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
            // ✅ CRITICAL FIX: Update Section Box in DB BEFORE detection
            // This ensures logic relying on "IsPointInBoundingBox" (SetReadyForPlacement) uses the FRESH section box
            try
            {
                if (document.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(document))
                    {
                        var sectionBoxService = new SectionBoxService();
                        sectionBoxService.CaptureAndStore(view3D, dbContext.Connection);
                        _log?.Invoke($"[SECTION-BOX] Updated DB Section Box from View: {view3D.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                _log?.Invoke($"[SECTION-BOX] Failed to update section box: {ex.Message}");
            }

            _log($"[METHOD3] DEBUG: DetectNewClashZones called with {currentIntersections?.Count ?? 0} intersections");
            
            // ✅ PERFORMANCE OPTIMIZATION: Streamlined fast-path for validated intersections
            // Force check flag and log execution path
            bool useStreamlined = OptimizationFlags.UseStreamlinedClashZoneCreation;
            _log($"[PERFORMANCE] Clash Zone Creation Strategy: UseStreamlinedClashZoneCreation = {useStreamlined}");
            
            if (useStreamlined)
            {
                _log($"[PERFORMANCE] Executing Streamlined Clash Zone Creation (Fast Path)...");
                return DetectNewClashZonesStreamlined(currentIntersections, document, clearanceSettings, selectedCategories);
            }
            else
            {
                _log($"[PERFORMANCE] Executing LEGACY Clash Zone Creation (Slow Path)...");
            }
            
            // ⚠️ CRITICAL: Check memory/timeout before starting heavy processing
            if (_memoryManager != null)
            {
                try
                {
                    _memoryManager.CheckLimits();
                    _log($"[MEMORY-MGR] Initial check passed: {_memoryManager.GetStatus()}");
                }
                catch (TimeoutException ex)
                {
                    _log($"[MEMORY-MGR] ⏱ TIMEOUT before processing: {ex.Message}");
                    SafeFileLogger.SafeAppendText("clash_zone_timeouts.log", $"Timeout before DetectNewClashZones processing: {ex.Message}");
                    throw; // Re-throw to let caller handle gracefully
                }
                catch (OutOfMemoryException ex)
                {
                    _log($"[MEMORY-MGR] 💾 MEMORY LIMIT before processing: {ex.Message}");
                    SafeFileLogger.SafeAppendText("clash_zone_memory.log", $"Memory limit before DetectNewClashZones processing: {ex.Message}");
                    throw; // Re-throw to let caller handle gracefully
                }
            }
            
            // Add build timestamp to refresh_debug.log
            try
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                var ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location)?.FileVersion ?? "?";
                var ts = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLASH_DEBUG] [BUILD] {ts} Assembly={System.IO.Path.GetFileName(asm.Location)} Version={ver} Path={asm.Location}\n");
                }
            }
            catch { }
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
            
            // ✅ DEBUG: Initialize counters BEFORE priority sorting (must be accessible for summary logging)
            int ductWallBeforePriority = 0;
            int ductWallAfterPriority = 0;
            
            // ✅ DEBUG: Log Duct-Wall count BEFORE priority sorting
            ductWallBeforePriority = enhancedIntersections.Where(i => 
            {
                var mepCat = GetElementCategoryName(i.Item1);
                var structType = i.Item2.Category?.Name ?? "";
                return (mepCat == "Ducts" || mepCat == "Duct Curves" || mepCat == "Duct Accessories") 
                    && (structType == "Walls" || structType == "Wall");
            }).Count();
            _log($"[DEBUG-COUNT] BEFORE PRIORITY SORT: {ductWallBeforePriority} Duct-Wall intersections");
            
            // FOOLPROOF METHOD: Always process dampers first, then ducts (category-based priority)
            List<(Element, Element, BoundingBoxXYZ, XYZ)> prioritizedIntersections;
            try
            {
                prioritizedIntersections = PrioritizeIntersectionsByCategory(enhancedIntersections);
                _log($"[PRIORITY] Processed {prioritizedIntersections.Count} intersections in priority order");
                
                // ✅ DEBUG: Log Duct-Wall count AFTER priority sorting
                ductWallAfterPriority = prioritizedIntersections.Where(i => 
                {
                    var mepCat = GetElementCategoryName(i.Item1);
                    var structType = i.Item2.Category?.Name ?? "";
                    return (mepCat == "Ducts" || mepCat == "Duct Curves" || mepCat == "Duct Accessories") 
                        && (structType == "Walls" || structType == "Wall");
                }).Count();
                _log($"[DEBUG-COUNT] AFTER PRIORITY SORT: {ductWallAfterPriority} Duct-Wall intersections");
            }
            catch (Exception ex)
            {
                _log($"[ERROR] Failed in priority method: {ex.Message}");
                prioritizedIntersections = enhancedIntersections;
                // If priority sort failed, after count equals before count
                ductWallAfterPriority = ductWallBeforePriority;
            }
            
            // ✅ OPTIMIZATION: Pre-collect sleeves by category AFTER priority sort (calculate once, use many times)
            // OOP Pattern: Two paths - optimized if sleeves exist, fallback if no sleeves
            var sleeveSpatialIndexesByCategory = new Dictionary<string, (Dictionary<(double X, double Y, double Z), int> SpatialIndex, HashSet<int> SleeveIds, List<FamilyInstance> CategorySleeves)>();
            
            if (selectedCategories != null && selectedCategories.Count > 0)
            {
                foreach (var category in selectedCategories)
                {
                    if (string.IsNullOrWhiteSpace(category))
                        continue;
                    
                    var (spatialIndex, sleeveIds, categorySleeves) = PreCollectSleevesByCategory(document, category);
                    
                    // Only store if sleeves exist (optimized path)
                    if (categorySleeves.Count > 0)
                    {
                        sleeveSpatialIndexesByCategory[category] = (spatialIndex, sleeveIds, categorySleeves);
                        if (!DeploymentConfiguration.DeploymentMode)
                            _log($"[OPTIMIZED-SLEEVE-LOOKUP] Category '{category}': {categorySleeves.Count} sleeves found → Using OPTIMIZED path");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            _log($"[OPTIMIZED-SLEEVE-LOOKUP] Category '{category}': No sleeves found → Using FALLBACK path");
                    }
                }
            }
            else
            {
                // If no selected categories, extract unique categories from intersections
                var uniqueCategories = prioritizedIntersections
                    .Select(i => GetElementCategoryName(i.Item1))
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                
                foreach (var category in uniqueCategories)
                {
                    var (spatialIndex, sleeveIds, categorySleeves) = PreCollectSleevesByCategory(document, category);
                    
                    if (categorySleeves.Count > 0)
                    {
                        sleeveSpatialIndexesByCategory[category] = (spatialIndex, sleeveIds, categorySleeves);
                        if (!DeploymentConfiguration.DeploymentMode)
                            _log($"[OPTIMIZED-SLEEVE-LOOKUP] Category '{category}': {categorySleeves.Count} sleeves found → Using OPTIMIZED path");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            _log($"[OPTIMIZED-SLEEVE-LOOKUP] Category '{category}': No sleeves found → Using FALLBACK path");
                    }
                }
            }
            
            // DEBUG: Log each intersection being processed in priority order
            for (int i = 0; i < prioritizedIntersections.Count; i++)
            {
                var (mepElement, structuralElement, boundingBox, intersectionPoint) = prioritizedIntersections[i];
                var mepCategory = GetElementCategoryName(mepElement);
                _log($"[DEBUG] Processing intersection {i + 1}/{prioritizedIntersections.Count}: MEP={mepElement.Id} ({mepCategory}) <-> Structural={structuralElement.Id}");
            }
            
            // ✅ OPTIMIZATION: Build O(1) Lookup Cache for existing clash zones
            // This prevents O(N) allocation + search for every single intersection
            BuildClashZoneLookup();
            
            _log($"Existing clash zones count: {_clashZoneStorage.ClashZones.Count}");
            
            // ✅ DEBUG: Initialize counters for tracking Duct-Wall clash zones through filters
            int ductWallAfterValidation = 0;
            int ductWallAfterDamperCheck = 0;
            int ductWallAfterPenetration = 0;
            int ductWallAfterExistingCheck = 0;
            int ductWallClashZonesCreated = 0;
            int ductWallSkippedInvalid = 0;
            int ductWallSkippedDamper = 0;
            int ductWallSkippedPenetration = 0;
            int ductWallSkippedExisting = 0;
            
            int processedCount = 0;
            const int MEMORY_CHECK_INTERVAL = 100; // Check memory every 100 intersections
            
            foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in prioritizedIntersections)
            {
                // ✅ WALL CENTERLINE POINT: Calculate final placement point using bbox method (same as dampers)
                // This is passed to CreateClashZone to save directly to SleevePlacementPoint - avoids adjustment during placement
                // For dampers, this is calculated in DamperProcessingService using bbox method
                // ✅ PERFORMANCE: Bbox method is much cheaper than ray-trace (no ReferenceIntersector, no 3D view lookup)
                XYZ? calculatedWallCenterline = null;
                string mepCategoryName = GetElementCategoryName(mepElement);
                if (structuralElement is Wall wall)
                {
                    // ✅ BBOX METHOD: For ducts, pipes, and cable trays, use bbox method (same as dampers)
                    // This calculates the final placement point at wall centerline, saving directly to SleevePlacementPoint
                    // Eliminates need for WallCenterlinePoint columns and adjustment logic during placement
                    calculatedWallCenterline = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetWallCenterlinePointFromBbox(
                        wall, intersectionPoint, document);
                    
                    // ✅ DIAGNOSTIC: Log wall centerline calculation for non-damper categories
                    if (!DeploymentConfiguration.DeploymentMode && !string.Equals(mepCategoryName, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                    {
                        SafeFileLogger.SafeAppendText("wall_centerline_calc.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CALC] Zone (MEP={mepElement.Id}, Host={wall.Id}, Category={mepCategoryName}): " +
                            $"Calculated PlacementPoint=({calculatedWallCenterline?.X:F6}ft, {calculatedWallCenterline?.Y:F6}ft, {calculatedWallCenterline?.Z:F6}ft), " +
                            $"Intersection=({intersectionPoint.X:F6}ft, {intersectionPoint.Y:F6}ft, {intersectionPoint.Z:F6}ft)\n");
                    }
                }
                else if (structuralElement is FamilyInstance framingInstance && 
                         framingInstance.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For framing, use the element centerline method
                    calculatedWallCenterline = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                        structuralElement, intersectionPoint, document);
                    
                    // ✅ DIAGNOSTIC: Log framing centerline calculation
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("wall_centerline_calc.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CALC] Zone (MEP={mepElement.Id}, Host={framingInstance.Id}, Category={mepCategoryName}, Type=Framing): " +
                            $"Calculated Centerline=({calculatedWallCenterline?.X:F6}ft, {calculatedWallCenterline?.Y:F6}ft, {calculatedWallCenterline?.Z:F6}ft), " +
                            $"Intersection=({intersectionPoint.X:F6}ft, {intersectionPoint.Y:F6}ft, {intersectionPoint.Z:F6}ft)\n");
                    }
                }
                // For floors or other structural elements, calculatedWallCenterline remains null (will fallback to intersectionPoint)
                
                processedCount++;
                
                // ⚠️ CRITICAL: Check memory/timeout every N intersections to prevent crashes on large files
                if (_memoryManager != null && processedCount % MEMORY_CHECK_INTERVAL == 0)
                {
                    try
                    {
                        _memoryManager.CheckLimits();
                        _log($"[MEMORY-MGR] Check passed at intersection {processedCount}/{prioritizedIntersections.Count} - {_memoryManager.GetStatus()}");
                    }
                    catch (TimeoutException ex)
                    {
                        _log($"[MEMORY-MGR] ⏱ TIMEOUT at intersection {processedCount}/{prioritizedIntersections.Count}: {ex.Message}");
                        SafeFileLogger.SafeAppendText("clash_zone_timeouts.log", 
                            $"Timeout during DetectNewClashZones: Processed {processedCount}/{prioritizedIntersections.Count} intersections. {ex.Message}");
                        
                        // Return partial results instead of crashing
                        _log($"[MEMORY-MGR] Returning {newClashZones.Count} clash zones processed before timeout");
                        return newClashZones;
                    }
                    catch (OutOfMemoryException ex)
                    {
                        _log($"[MEMORY-MGR] 💾 MEMORY LIMIT at intersection {processedCount}/{prioritizedIntersections.Count}: {ex.Message}");
                        SafeFileLogger.SafeAppendText("clash_zone_memory.log", 
                            $"Memory limit during DetectNewClashZones: Processed {processedCount}/{prioritizedIntersections.Count} intersections. {ex.Message}");
                        
                        // Try to free memory once
                        try
                        {
                            _memoryManager.ForceCleanup();
                            
                            // Check again after cleanup
                            _memoryManager.CheckLimits();
                            _log($"[MEMORY-MGR] Memory cleanup succeeded, continuing...");
                        }
                        catch
                        {
                            // If cleanup didn't help, return partial results
                            _log($"[MEMORY-MGR] Memory cleanup failed, returning {newClashZones.Count} clash zones processed");
                            return newClashZones;
                        }
                    }
                }
                
                // ✅ DEBUG: Check if this is Duct-Wall for tracking
                bool isDuctWall = false;
                try
                {
                    var mepCat = GetElementCategoryName(mepElement);
                    var structType = structuralElement?.Category?.Name ?? "";
                    isDuctWall = (mepCat == "Ducts" || mepCat == "Duct Curves" || mepCat == "Duct Accessories") 
                        && (structType == "Walls" || structType == "Wall");
                }
                catch { }
                
                // CRITICAL FIX: Validate elements before creating clash zones
                if (mepElement == null || structuralElement == null)
                {
                    _log($"[OPTIMIZATION] ❌ SKIP VALIDATION: Invalid elements - MEP={mepElement?.Id}, Structural={structuralElement?.Id}");
                    if (isDuctWall) ductWallSkippedInvalid++;
                    continue;
                }
                
                _log($"[OPTIMIZATION] ✅ PASS VALIDATION: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                if (isDuctWall) ductWallAfterValidation++;

                // Method 3 is now implemented in IntersectionDetectionService.cs
                // Ducts with dampers at their ends are filtered out at intersection level

                // ✅ CRITICAL FIX: Check if elements are still valid in the document
                // ⚠️ OPTIMIZATION SAFETY: Since intersections come from IntersectionDetectionService which already validated elements,
                // we should trust them and NOT skip unless we're absolutely certain the element is deleted
                try
                {
                    var mepElementCheck = document.GetElement(mepElement.Id);
                    var structuralElementCheck = document.GetElement(structuralElement.Id);
                    
                    // ✅ OPTIMIZATION SAFETY: Be very conservative - only skip if we're CERTAIN element was deleted
                    // Linked elements return null from GetElement but are still valid
                    // Since intersections were already validated, we trust the element exists unless proven otherwise
                    
                    // Only skip if we can PROVE the element was deleted (check for negative/invalid IDs)
                    if (mepElement.Id.GetIntegerValue() <= 0)
                    {
                        _log($"SKIP: Invalid MEP element ID - MEP={mepElement.Id}");
                        if (isDuctWall) ductWallSkippedInvalid++;
                        continue;
                    }
                    
                    if (structuralElement.Id.GetIntegerValue() <= 0)
                    {
                        _log($"SKIP: Invalid structural element ID - Structural={structuralElement.Id}");
                        if (isDuctWall) ductWallSkippedInvalid++;
                        continue;
                    }
                    
                    // ✅ OPTIMIZATION SAFETY: Log linked elements but DON'T skip them
                    // GetElement returns null for linked elements, but they're still valid for clash zone creation
                    if (mepElementCheck == null || structuralElementCheck == null)
                    {
                        _log($"INFO: Processing linked element(s) - MEP={mepElement.Id} (found={mepElementCheck != null}), Structural={structuralElement.Id} (found={structuralElementCheck != null})");
                        // Continue processing - linked elements are valid
                    }
                }
                catch (Exception ex)
                {
                    // ✅ OPTIMIZATION SAFETY: Log error but DON'T skip unless we're certain
                    // Errors in validation shouldn't prevent clash zone creation
                    _log($"WARN: Error validating elements - MEP={mepElement.Id}, Structural={structuralElement.Id}, Error={ex.Message} - Continuing anyway");
                    // Continue processing - assume elements are valid if validation fails
                }

                // DIAGNOSTIC: Log detailed geometry information only when enabled
                if (OptimizationFlags.UseDiagnosticMode)
                {
                    _log($"=== GEOMETRY ANALYSIS ===");
                    _log($"MEP Element: {mepElement.Name} (ID: {mepElement.Id})");
                    _log($"Structural Element: {structuralElement.Name} (ID: {structuralElement.Id})");
                }
                
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

                // ✅ DUCT-DAMPER FLAG TRACKING: Track if we detected damper for this duct (to set flag on new clash zone)
                // Declared BEFORE duct-damper check so it's accessible throughout the method
                bool ductHasDamperNearby = false;
                
                // ✅ SOLID-COMPLIANT DAMPER FILTER: Gated by flag for safe rollout
                // Uses flag-based caching (HasDamperNearby) + proximity check fallback
                // Implements Open/Closed Principle - extensible through ZoneFilterService without modifying this logic
                if (OptimizationFlags.UseSOLIDCompliantDamperFilter)
                {
                    try
                    {
                        var mepCat = GetElementCategoryName(mepElement);
                        _log($"[DUCT-DAMPER] Checking element {mepElement.Id} with category '{mepCat}' against {damperLocations.Count} damper locations");
                        
                        if (string.Equals(mepCat, "Ducts", StringComparison.OrdinalIgnoreCase))
                        {
                            // ✅ STEP 1: Check existing clash zone for saved flag (fastest check - no proximity calculation needed)
                            var existingDuctClashZone = FindExistingClashZone(mepElement.Id, structuralElement.Id, intersectionPoint);
                            if (existingDuctClashZone != null && existingDuctClashZone.HasDamperNearby)
                            {
                                _log($"[OPTIMIZATION] ❌ SKIP DAMPER CHECK (FLAG): Duct {mepElement.Id} - HasDamperNearby flag is true from previous run, skipping duct");
                                if (isDuctWall) ductWallSkippedDamper++;
                                continue;
                            }
                            
                            // ✅ STEP 2: Run proximity check only if flag is not set (first run or flag reset)
                            bool isNearDamper = IsDuctNearDamperOnSameWall(mepElement, structuralElement.Id, intersectionPoint, damperLocations);
                            
                            if (isNearDamper)
                            {
                                _log($"[OPTIMIZATION] ❌ SKIP DAMPER CHECK (DETECTED): Duct {mepElement.Id} - damper present on same wall ({structuralElement.Id}) at intersection point, prioritizing damper sleeve");
                                
                                // ✅ CRITICAL: Set flag on existing clash zone if found
                                if (existingDuctClashZone != null)
                                {
                                    existingDuctClashZone.HasDamperNearby = true;
                                    _log($"[DUCT-DAMPER] ✓ Set HasDamperNearby=true on existing clash zone {existingDuctClashZone.Id}");
                                }
                                else
                                {
                                    // ✅ Track that this duct has damper nearby - will set flag if new clash zone is created
                                    ductHasDamperNearby = true;
                                    _log($"[DUCT-DAMPER] ✓ Duct {mepElement.Id} has damper nearby - will set flag on new clash zone if created");
                                }
                                
                                if (isDuctWall) ductWallSkippedDamper++;
                                continue;
                            }
                            else
                            {
                                // ✅ Clear flag if damper no longer nearby (damper might have been deleted)
                                if (existingDuctClashZone != null && existingDuctClashZone.HasDamperNearby)
                                {
                                    existingDuctClashZone.HasDamperNearby = false;
                                    _log($"[DUCT-DAMPER] ✓ Cleared HasDamperNearby flag - damper no longer nearby for clash zone {existingDuctClashZone.Id}");
                                }
                                
                                _log($"[OPTIMIZATION] ✅ PASS DAMPER CHECK: Duct {mepElement.Id} - no damper nearby on wall {structuralElement.Id}, proceeding with sleeve placement");
                                if (isDuctWall) ductWallAfterDamperCheck++;
                            }
                        }
                        else
                        {
                            // Not a duct, so damper check doesn't apply - count it as passed
                            if (isDuctWall) ductWallAfterDamperCheck++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _log($"Error in duct-damper priority filter: {ex.Message}");
                    }
                }
                else
                {
                    // ✅ FALLBACK: When flag is disabled, skip damper check and treat all elements as passed
                    _log($"[INFO] Damper filter disabled (UseSOLIDCompliantDamperFilter=false) - skipping damper check for element {mepElement.Id}");
                    if (isDuctWall) ductWallAfterDamperCheck++;
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
                        var hostNormal = WallDirectionService.GetStructuralElementNormal(structuralElement); // ✅ OOP: Use centralized service
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
                        bool isPerpendicularPenetration = false;
                        double dot = 0.0; // ✅ FIX: Declare dot outside if block for logging
                        
                        if (hostThickness > 1e-6 && crossSize > 1e-6 && dLen > 1e-6 && nLen > 1e-6)
                        {
                            var dNorm = new XYZ(mepDir.X / dLen, mepDir.Y / dLen, mepDir.Z / dLen);
                            var nNorm = new XYZ(hostNormal.X / nLen, hostNormal.Y / nLen, hostNormal.Z / nLen);
                            dot = Math.Abs(dNorm.X * nNorm.X + dNorm.Y * nNorm.Y + dNorm.Z * nNorm.Z);
                            
                            _log($"[DEBUG]   Normalized MEP Dir: ({dNorm.X:F3}, {dNorm.Y:F3}, {dNorm.Z:F3})");
                            _log($"[DEBUG]   Normalized Host Normal: ({nNorm.X:F3}, {nNorm.Y:F3}, {nNorm.Z:F3})");
                            _log($"[DEBUG]   Dot Product: {dot:F3}");
                            _log($"[DEBUG]   Cross Size: {crossSize:F3}");
                            
                            // ✅ CRITICAL FIX: Detect perpendicular penetrations (duct running parallel to wall face)
                            // When dot < 0.1, the MEP element is nearly perpendicular to wall normal (parallel to wall face)
                            // This is a VALID penetration - the element goes through the wall at 90 degrees
                            const double PerpendicularThreshold = 0.1;
                            if (dot < PerpendicularThreshold)
                            {
                                isPerpendicularPenetration = true;
                                // For perpendicular penetrations, check if cross-section is significant relative to wall thickness
                                // If crossSize >= 10% of wall thickness, it's a valid penetration
                                penetrationRatio = crossSize / hostThickness;
                                _log($"[DEBUG]   PERPENDICULAR PENETRATION detected (dot={dot:F3} < {PerpendicularThreshold})");
                                _log($"[DEBUG]   Using cross-section/thickness ratio: {penetrationRatio:F3}");
                            }
                            else
                            {
                                // Normal angled penetration - use standard formula
                                penetrationRatio = (hostThickness * dot) / crossSize;
                            }
                        }

                        // Threshold: require at least 1% penetration of cross-section OR 10% cross-section/thickness for perpendicular
                        const double MinPenetrationRatio = 0.01;
                        const double MinPerpendicularRatio = 0.10; // For perpendicular penetrations
                        
                        _log($"[DEBUG] Penetration check: ratio={penetrationRatio:F3}, threshold={MinPenetrationRatio:F2}, hostThickness={hostThickness:F3}, crossSize={crossSize:F3}, perpendicular={isPerpendicularPenetration}");
                        
                        bool shouldSkip = false;
                        if (isPerpendicularPenetration)
                        {
                            // For perpendicular penetrations, require cross-section to be at least 10% of wall thickness
                            shouldSkip = penetrationRatio < MinPerpendicularRatio;
                            if (shouldSkip)
                            {
                                _log($"[OPTIMIZATION] ❌ SKIP PENETRATION: Insufficient perpendicular penetration (crossSize/thickness={penetrationRatio:F3} < {MinPerpendicularRatio:F2}) for {hostTypeName}. MEP={mepElement.Id}, Structural={structuralElement.Id}");
                                _log($"[OPTIMIZATION]   Details: crossSize={crossSize:F3}ft, hostThickness={hostThickness:F3}ft, ratio={penetrationRatio:F3}");
                            }
                        }
                        else
                        {
                            // For angled penetrations, use original threshold
                            shouldSkip = penetrationRatio < MinPenetrationRatio;
                            if (shouldSkip)
                            {
                                _log($"[OPTIMIZATION] ❌ SKIP PENETRATION: Insufficient penetration (ratio={penetrationRatio:F3} < {MinPenetrationRatio:F2}) for {hostTypeName}. MEP={mepElement.Id}, Structural={structuralElement.Id}");
                                _log($"[OPTIMIZATION]   Details: hostThickness={hostThickness:F3}ft, crossSize={crossSize:F3}ft, dot={dot:F3}, ratio={penetrationRatio:F3}");
                            }
                        }
                        
                        if (shouldSkip)
                        {
                            if (isDuctWall) ductWallSkippedPenetration++;
                            continue;
                        }
                        
                        // Log pass message with correct threshold
                        if (isPerpendicularPenetration)
                        {
                            _log($"[OPTIMIZATION] ✅ PASS PENETRATION (perpendicular): ratio={penetrationRatio:F3} >= {MinPerpendicularRatio:F2}");
                        }
                        else
                        {
                            _log($"[OPTIMIZATION] ✅ PASS PENETRATION: ratio={penetrationRatio:F3} >= {MinPenetrationRatio:F2}");
                        }
                        if (isDuctWall) ductWallAfterPenetration++;
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
                    _log($"[OPTIMIZATION] ✅ PASS EXISTING CHECK: No existing clash zone found for MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    if (isDuctWall) ductWallAfterExistingCheck++;
                }
                else
                {
                    _log($"[OPTIMIZATION] ❌ SKIP EXISTING CHECK: Clash zone already exists for MEP={mepElement.Id}, Structural={structuralElement.Id} (Existing ID: {existingClashZone.Id})");
                    if (isDuctWall) ductWallSkippedExisting++;
                }

                if (existingClashZone == null)
                {
                    // Check if there's an invalid clash zone for the same geometry (by hash)
                    var invalidClashZone = FindInvalidClashZoneByGeometry(mepElement, structuralElement);
                    
                    if (invalidClashZone != null)
                    {
                        // Replace invalid clash zone with valid one
                        _log($"Replacing invalid clash zone {invalidClashZone.Id} with valid ElementIds");
                        _clashZoneStorage.ClashZones.Remove(invalidClashZone);
                        
                        // ✅ OOP PATTERN: Pass spatial index if available for optimized path
                        var mepCategoryForLookupInvalid = GetElementCategoryName(mepElement);
                        Dictionary<(double X, double Y, double Z), int> spatialIndexForCategoryInvalid = null;
                        HashSet<int> sleeveIdsForCategoryInvalid = null;
                        
                        if (sleeveSpatialIndexesByCategory.TryGetValue(mepCategoryForLookupInvalid, out var sleeveDataInvalid))
                        {
                            spatialIndexForCategoryInvalid = sleeveDataInvalid.SpatialIndex;
                            sleeveIdsForCategoryInvalid = sleeveDataInvalid.SleeveIds;
                        }
                        
                        // ✅ WALL CENTERLINE POINT: Calculate once using WallCenterlineHelper (for ducts, pipes, cable trays)
                        // This is passed to CreateClashZone to save in DB - avoids duplication
                        // For dampers, this is calculated in DamperProcessingService using bbox method
                        XYZ? calculatedWallCenterlineInvalid = null;
                        if (structuralElement is Wall wallInvalid)
                        {
                            // ✅ RAY-TRACE METHOD: For ducts, pipes, and cable trays, use ray-trace to find 2 wall faces and calculate midpoint
                            calculatedWallCenterlineInvalid = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetWallCenterlinePointFromBbox(
                                wallInvalid, intersectionPoint, document);
                            
                            // ✅ DIAGNOSTIC: Log wall centerline calculation for invalid category path
                            if (!DeploymentConfiguration.DeploymentMode && !string.Equals(mepCategoryForLookupInvalid, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                            {
                                SafeFileLogger.SafeAppendText("wall_centerline_calc.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CALC-INVALID] Zone (MEP={mepElement.Id}, Host={wallInvalid.Id}, Category={mepCategoryForLookupInvalid}): " +
                                    $"Calculated WallCenterline=({calculatedWallCenterlineInvalid?.X:F6}ft, {calculatedWallCenterlineInvalid?.Y:F6}ft, {calculatedWallCenterlineInvalid?.Z:F6}ft), " +
                                    $"Intersection=({intersectionPoint.X:F6}ft, {intersectionPoint.Y:F6}ft, {intersectionPoint.Z:F6}ft)\n");
                            }
                        }
                        else if (structuralElement is FamilyInstance framingInstanceInvalid && 
                                 framingInstanceInvalid.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
                        {
                            // For framing, use the element centerline method
                            calculatedWallCenterlineInvalid = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                                structuralElement, intersectionPoint, document);
                            
                            // ✅ DIAGNOSTIC: Log framing centerline calculation for invalid category path
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("wall_centerline_calc.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [CALC-INVALID] Zone (MEP={mepElement.Id}, Host={framingInstanceInvalid.Id}, Category={mepCategoryForLookupInvalid}, Type=Framing): " +
                                    $"Calculated Centerline=({calculatedWallCenterlineInvalid?.X:F6}ft, {calculatedWallCenterlineInvalid?.Y:F6}ft, {calculatedWallCenterlineInvalid?.Z:F6}ft), " +
                                    $"Intersection=({intersectionPoint.X:F6}ft, {intersectionPoint.Y:F6}ft, {intersectionPoint.Z:F6}ft)\n");
                            }
                        }
                        // For floors or other structural elements, calculatedWallCenterlineInvalid remains null (will fallback to intersectionPoint)
                        
                        var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document, clearanceSettings, spatialIndexForCategoryInvalid, sleeveIdsForCategoryInvalid, calculatedWallCenterlineInvalid);
                        
                        // ✅ CRITICAL: Set HasDamperNearby flag if duct-damper combo was detected
                        if (ductHasDamperNearby)
                        {
                            newClashZone.HasDamperNearby = true;
                            _log($"[DUCT-DAMPER] ✓ Set HasDamperNearby=true on REPLACED clash zone {newClashZone.Id} for duct {mepElement.Id}");
                        }
                        
                        newClashZone.IsCurrentClash = true; // ✅ DEBUG: Mark as current refresh clash
                        // ✅ MEMORY: Drop heavy API objects immediately after populating numeric fields
                        newClashZone.ClearRevitApiObjects();
                        newClashZones.Add(newClashZone);
                        _clashZoneStorage.ClashZones.Add(newClashZone);
                        _log($"Replaced invalid clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                        if (isDuctWall) ductWallClashZonesCreated++;
                        
                        // ✅ MEMORY PROFILING: Track memory AFTER clash zone is added (correct timing)
                        if (_memoryProfiler != null)
                        {
                            _memoryProfiler.RecordClashZoneProcessing(newClashZones.Count, prioritizedIntersections.Count);
                        }
                    }
                    else
                {
                    // Create new clash zone
                    _log($"[DEBUG] About to create clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                    try
                    {
                        // ✅ CRITICAL DEBUG: Log intersection point BEFORE creating ClashZone
                        bool isZeroPoint = Math.Abs(intersectionPoint.X) < 1e-9 && Math.Abs(intersectionPoint.Y) < 1e-9 && Math.Abs(intersectionPoint.Z) < 1e-9;
                        if (isZeroPoint)
                        {
                            _log($"[⚠️ ZERO POINT WARNING] Creating ClashZone with ZERO intersection point: MEP={mepElement.Id}, Structural={structuralElement.Id}, Point=({intersectionPoint.X},{intersectionPoint.Y},{intersectionPoint.Z})");
                            try 
                            { 
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    var debugPath = SafeFileLogger.GetLogFilePath("refresh_intersection_debug.log");
                                    File.AppendAllText(debugPath, $"[{DateTime.Now}] ⚠️ ZERO POINT: MEP={mepElement.Id}, Structural={structuralElement.Id}, Point=({intersectionPoint.X},{intersectionPoint.Y},{intersectionPoint.Z}), Document={document?.Title}\n");
                                }
                            } 
                            catch { }
                        }
                        
                        // ✅ OOP PATTERN: Pass spatial index if available for optimized path
                        var mepCategoryForLookup = GetElementCategoryName(mepElement);
                        Dictionary<(double X, double Y, double Z), int> spatialIndexForCategory = null;
                        HashSet<int> sleeveIdsForCategory = null;
                        
                        if (sleeveSpatialIndexesByCategory.TryGetValue(mepCategoryForLookup, out var sleeveData))
                        {
                            spatialIndexForCategory = sleeveData.SpatialIndex;
                            sleeveIdsForCategory = sleeveData.SleeveIds;
                        }
                        
                        var newClashZone = CreateClashZone(mepElement, structuralElement, intersectionPoint, boundingBox, document, clearanceSettings, spatialIndexForCategory, sleeveIdsForCategory, calculatedWallCenterline);
                        
                        // ✅ CRITICAL: Set HasDamperNearby flag if duct-damper combo was detected
                        if (ductHasDamperNearby)
                        {
                            newClashZone.HasDamperNearby = true;
                            _log($"[DUCT-DAMPER] ✓ Set HasDamperNearby=true on NEW clash zone {newClashZone.Id} for duct {mepElement.Id}");
                        }
                        
                        // ✅ CRITICAL DEBUG: Verify coordinates AFTER creation
                        bool xmlIsZero = Math.Abs(newClashZone.IntersectionPointX) < 1e-9 && Math.Abs(newClashZone.IntersectionPointY) < 1e-9 && Math.Abs(newClashZone.IntersectionPointZ) < 1e-9;
                        if (xmlIsZero)
                        {
                            _log($"[⚠️ XML ZERO WARNING] ClashZone created with ZERO XML coordinates: ID={newClashZone.Id}, MEP={mepElement.Id}, Structural={structuralElement.Id}");
                            try 
                            { 
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    var debugPath = SafeFileLogger.GetLogFilePath("refresh_intersection_debug.log");
                                    File.AppendAllText(debugPath, $"[{DateTime.Now}] ⚠️ XML ZERO: Zone={newClashZone.Id}, MEP={mepElement.Id}, Structural={structuralElement.Id}, IP_XML=({newClashZone.IntersectionPointX},{newClashZone.IntersectionPointY},{newClashZone.IntersectionPointZ})\n");
                                }
                            } 
                            catch { }
                        }
                        
                        newClashZone.IsCurrentClash = true; // ✅ DEBUG: Mark as current refresh clash
                        // ✅ MEMORY: Drop heavy API objects immediately after populating numeric fields
                        newClashZone.ClearRevitApiObjects();
                        
                        // ✅ CRITICAL FIX: Do NOT check Global XML here using old MEP+Host key system
                        // The old GlobalFlagManager used MEP+Host as key, which meant:
                        // - Same MEP crossing Wall 1 and Wall 2 (different hosts, different intersection points)
                        // - Would only find ONE entry in Global XML (wrong - should have separate entries per intersection)
                        // 
                        // NEW SYSTEM: GlobalIndexService uses GUID as key (correct per intersection)
                        // Flags will be synced from Global XML AFTER clash zones are loaded from Filter XML
                        // This happens in RefreshService line 852-910, which correctly uses GUID lookup
                        // 
                        // For NEW clash zones (not in Filter XML), they start with flags=false (correct)
                        // For EXISTING clash zones (loaded from Filter XML), flags are synced by GUID (correct)
                        _log($"[GLOBAL-XML] New clash zone created: ClashZone {newClashZone.Id} for MEP={mepElement.Id}, Structural={structuralElement.Id}, Intersection=({intersectionPoint.X:F3},{intersectionPoint.Y:F3},{intersectionPoint.Z:F3}) - flags will be synced from Global XML if exists");
                        
                        newClashZones.Add(newClashZone);
                        _clashZoneStorage.ClashZones.Add(newClashZone);
                        _log($"New clash zone detected: MEP={mepElement.Id}, Structural={structuralElement.Id}");
                        if (isDuctWall) ductWallClashZonesCreated++;
                        
                        // ✅ MEMORY PROFILING: Track memory AFTER clash zone is added (correct timing)
                        if (_memoryProfiler != null)
                        {
                            _memoryProfiler.RecordClashZoneProcessing(newClashZones.Count, prioritizedIntersections.Count);
                        }
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
                    // ✅ CRITICAL FIX: Verify sleeve actually exists before preserving resolved zone
                    // If IsResolved=true but sleeve doesn't exist (deleted or invalid), reset flag and allow update
                    bool sleeveActuallyExists = false;
                    if (existingClashZone.SleeveInstanceId > 0)
                    {
                        try
                        {
                            var sleeveElement = document.GetElement(ElementIdCompat.FromValue(existingClashZone.SleeveInstanceId));
                            if (sleeveElement != null && sleeveElement is FamilyInstance fi && fi.IsValidObject)
                            {
                                sleeveActuallyExists = true;
                            }
                        }
                        catch { }
                    }
                    if (existingClashZone.ClusterSleeveInstanceId > 0 && !sleeveActuallyExists)
                    {
                        try
                        {
                            var clusterElement = document.GetElement(ElementIdCompat.FromValue(existingClashZone.ClusterSleeveInstanceId));
                            if (clusterElement != null && clusterElement is FamilyInstance fi && fi.IsValidObject)
                            {
                                sleeveActuallyExists = true;
                            }
                        }
                        catch { }
                    }
                    
                    // If sleeve doesn't exist, reset flag and allow zone to be updated
                    if (!sleeveActuallyExists)
                    {
                        _log($"[CRITICAL-FIX] Zone {existingClashZone.Id} marked as IsResolved=true but sleeve doesn't exist (SleeveId={existingClashZone.SleeveInstanceId}, ClusterId={existingClashZone.ClusterSleeveInstanceId}) - resetting flag and allowing update");
                        existingClashZone.IsResolved = false;
                        existingClashZone.IsClusterResolved = false;
                        if (existingClashZone.SleeveInstanceId > 0 && !sleeveActuallyExists)
                            existingClashZone.SleeveInstanceId = 0;
                        if (existingClashZone.ClusterSleeveInstanceId > 0 && !sleeveActuallyExists)
                            existingClashZone.ClusterSleeveInstanceId = 0;
                        // Now allow update (fall through to UpdateExistingClashZone)
                        UpdateExistingClashZone(existingClashZone, mepElement, structuralElement, intersectionPoint, boundingBox, document);
                        _log($"Updated existing clash zone after resetting invalid resolved flag: {existingClashZone.Id}");
                        continue; // Skip the rest of the resolved zone handling
                    }
                    
                    // ✅ CRITICAL FIX: Update MepElementOrientation for resolved zones with (0,0,0) orientation
                    // This fixes old XML data that doesn't have MepElementOrientation calculated
                    if (existingClashZone.MepElementOrientation == null || existingClashZone.MepElementOrientation == XYZ.Zero)
                    {
                        var mepDir = GetMepElementOrientation(mepElement);
                        existingClashZone.MepElementOrientation = mepDir;
                        
                        // Also update MepElementOrientationDirection for floor sleeves
                        // Determine if width runs along X or Y axis based on bounding box comparison
                        if (existingClashZone.StructuralElementType == "Floor" && mepElement is Duct duct)
                        {
                            try
                            {
                                var (orientation, widthDirection) = Helpers.MepElementOrientationHelper.GetDuctWidthOrientation(duct);
                                existingClashZone.MepElementOrientationDirection = orientation == "X-ORIENTED" ? "X" : "Y";
                                _log($"✅ UPDATED MepElementOrientationDirection for resolved clash zone {existingClashZone.Id}: '{existingClashZone.MepElementOrientationDirection}' ({orientation})");
                            }
                            catch (Exception ex)
                            {
                                _log($"Error updating MepElementOrientationDirection for resolved clash zone {existingClashZone.Id}: {ex.Message}");
                                // Fallback: use X if we can't determine
                                existingClashZone.MepElementOrientationDirection = "X";
                            }
                        }
                        
                        // ✅ FLOOR ROTATION: Calculate rotation angle for floor sleeves
                        if (existingClashZone.StructuralElementType == "Floor" || existingClashZone.StructuralElementType == "Floors")
                        {
                            existingClashZone.MepElementRotationAngle = CalculateMepElementRotationAngle(
                                existingClashZone.StructuralElementType, 
                                mepDir,
                                mepElement); // Pass element for vertical element rotation calculation
                            _log($"✅ UPDATED MepElementRotationAngle for resolved clash zone {existingClashZone.Id}: {existingClashZone.MepElementRotationAngle * 180 / Math.PI:F1}°");
                        }
                        
                        existingClashZone.LastUpdated = DateTime.Now;
                        _log($"✅ UPDATED MepElementOrientation for resolved clash zone {existingClashZone.Id}: ({mepDir.X:F3}, {mepDir.Y:F3}, {mepDir.Z:F3})");
                    }
                    else if (existingClashZone.MepElementRotationAngle == 0.0 && 
                             (existingClashZone.StructuralElementType == "Floor" || existingClashZone.StructuralElementType == "Floors"))
                    {
                        // ✅ FLOOR ROTATION: Update rotation angle if it's missing (old XML data)
                        // Note: For old XML data, we may not have access to mepElement, so use orientation only
                        // This is a fallback - new clash zones will have the element available
                        existingClashZone.MepElementRotationAngle = CalculateMepElementRotationAngle(
                            existingClashZone.StructuralElementType, 
                            existingClashZone.MepElementOrientation,
                            mepElement); // Pass element if available for vertical element rotation
                        _log($"✅ UPDATED MepElementRotationAngle for existing clash zone {existingClashZone.Id}: {existingClashZone.MepElementRotationAngle * 180 / Math.PI:F1}°");
                    }
                    
                    // Preserve resolved clash zones during refresh - keep them in the list
                    _log($"Preserved resolved clash zone: {existingClashZone.Id} (IsResolved={existingClashZone.IsResolved})");
                }
            }
            
            // ✅ CRASH-SAFE: Use SafeFileLogger instead of hardcoded path
            // SafeFileLogger automatically creates directories and handles missing paths gracefully
            // ✅ DECLARE ONCE: Declare refreshLogName early so it can be reused throughout the method
            string refreshLogName = $"Refresh_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log";
            
            // ✅ OOP REFACTORING: Use FlagManager for flag reset if available
            // Otherwise fall back to legacy method for backward compatibility
            if (_flagManager != null && _clashZoneStorage?.ClashZones != null)
            {
                // ✅ UNIFIED RESET: Use FlagManager unified method - single call handles everything
                if (selectedCategories != null && selectedCategories.Count > 0)
                {
                    
                    // ✅ UNIFIED RESET: Single method call handles everything
                    // ✅ CRITICAL: Pass refreshLogName so detailed debug logs are written
                    var allClashZones = _clashZoneStorage.ClashZones
                        .Where(cz => selectedCategories.Contains(cz.MepElementCategory, StringComparer.OrdinalIgnoreCase))
                        .ToList();
                    
                    // ✅ Create dictionary for clash zones by category (legacy FlagManager expects this format)
                    var clashZonesByCategory = allClashZones
                        .GroupBy(cz => cz.MepElementCategory, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
                    
                    int resetCount = _flagManager.ResetFlagsForDeletedSleeves(allClashZones, selectedCategories, refreshLogName);
                    
                    if (resetCount > 0)
                    {
                        _log($"[FLAG-MANAGER] Reset flags for {resetCount} deleted sleeves using unified FlagManager method");
                    }
                    else
                    {
                        _log($"[FLAG-MANAGER] No flags reset - all sleeves exist");
                    }
                }
            }
            else
            {
                // Fallback to legacy method for backward compatibility
                ResetResolvedFlagForDeletedSleeves(document, selectedCategories);
            }
            
            // CRITICAL FIX: Remove duplicate clash zones (same MEP + structural element)
            RemoveDuplicateClashZones();
            
            // ✅ MEMORY OPTIMIZATION: Force FULL GC after clash zone detection to release temporary objects
            System.GC.Collect(2, System.GCCollectionMode.Forced, true);
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect(2, System.GCCollectionMode.Forced, true);
            
            // ✅ MEMORY PROFILING: Final snapshot before returning
            if (_memoryProfiler != null)
            {
                _memoryProfiler.TakeSnapshot("DETECT_NEW_CLASH_ZONES_COMPLETE", newClashZones.Count);
            }
            
            // Update storage metadata
            _clashZoneStorage.LastUpdated = DateTime.Now;
            _clashZoneStorage.DocumentPath = documentPath;
            _clashZoneStorage.DocumentHash = documentHash;
            
            _log($"Clash zone detection complete. New zones: {newClashZones.Count}");
            // ✅ DEBUG: Log final summary of Duct-Wall clash zones through optimization pipeline
            _log($"[DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            _log($"[DEBUG-COUNT] DUCT-WALL CLASH ZONES OPTIMIZATION PIPELINE SUMMARY");
            _log($"[DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            _log($"[DEBUG-COUNT] STEP 1 - BEFORE OPTIMIZATION (Priority Sort): {ductWallBeforePriority} Duct-Wall intersections");
            _log($"[DEBUG-COUNT] STEP 2 - AFTER PRIORITY SORT: {ductWallAfterPriority} Duct-Wall intersections (changed: {ductWallAfterPriority - ductWallBeforePriority:+0;-0;=0})");
            _log($"[DEBUG-COUNT] STEP 3 - AFTER VALIDATION: {ductWallAfterValidation} Duct-Wall intersections (✅ passed, ❌ skipped: {ductWallSkippedInvalid})");
            _log($"[DEBUG-COUNT] STEP 4 - AFTER DAMPER CHECK: {ductWallAfterDamperCheck} Duct-Wall intersections (✅ passed, ❌ skipped: {ductWallSkippedDamper})");
            _log($"[DEBUG-COUNT] STEP 5 - AFTER PENETRATION FILTER: {ductWallAfterPenetration} Duct-Wall intersections (✅ passed, ❌ skipped: {ductWallSkippedPenetration})");
            _log($"[DEBUG-COUNT] STEP 6 - AFTER EXISTING CHECK: {ductWallAfterExistingCheck} Duct-Wall intersections (✅ passed, ❌ skipped: {ductWallSkippedExisting})");
            _log($"[DEBUG-COUNT] STEP 7 - FINAL CLASH ZONES CREATED: {ductWallClashZonesCreated} Duct-Wall clash zones");
            _log($"[DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            _log($"[DEBUG-COUNT] VERIFICATION: {ductWallAfterExistingCheck} should equal {ductWallClashZonesCreated} (after existing check = final created)");
            var totalSkipped = ductWallSkippedInvalid + ductWallSkippedDamper + ductWallSkippedPenetration + ductWallSkippedExisting;
            var totalProcessed = ductWallAfterValidation + totalSkipped;
            _log($"[DEBUG-COUNT] VERIFICATION: Total processed ({totalProcessed}) = Passed ({ductWallAfterValidation}) + Skipped ({totalSkipped})");
            _log($"[DEBUG-COUNT] ═══ TOTAL FILTERED OUT: {ductWallBeforePriority - ductWallClashZonesCreated} Duct-Wall intersections ═══");
            _log($"[DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            
            // ✅ REUSE: refreshLogName already declared earlier in the method
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] DUCT-WALL CLASH ZONES OPTIMIZATION PIPELINE SUMMARY");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 1 - BEFORE OPTIMIZATION: {ductWallBeforePriority} Duct-Wall intersections");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 2 - AFTER PRIORITY SORT: {ductWallAfterPriority} (changed: {ductWallAfterPriority - ductWallBeforePriority:+0;-0;=0})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 3 - AFTER VALIDATION: {ductWallAfterValidation} (✅ passed, ❌ skipped: {ductWallSkippedInvalid})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 4 - AFTER DAMPER CHECK: {ductWallAfterDamperCheck} (✅ passed, ❌ skipped: {ductWallSkippedDamper})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 5 - AFTER PENETRATION FILTER: {ductWallAfterPenetration} (✅ passed, ❌ skipped: {ductWallSkippedPenetration})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 6 - AFTER EXISTING CHECK: {ductWallAfterExistingCheck} (✅ passed, ❌ skipped: {ductWallSkippedExisting})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] STEP 7 - FINAL CLASH ZONES CREATED: {ductWallClashZonesCreated} Duct-Wall clash zones");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] VERIFICATION: {ductWallAfterExistingCheck} should equal {ductWallClashZonesCreated}");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] VERIFICATION: Total processed ({totalProcessed}) = Passed ({ductWallAfterValidation}) + Skipped ({totalSkipped})");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] ═══ TOTAL FILTERED OUT: {ductWallBeforePriority - ductWallClashZonesCreated} Duct-Wall intersections ═══");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [DEBUG-COUNT] ════════════════════════════════════════════════════════════════════════════");
            
            try
            {
                // ✅ MEMORY: Clear geometry caches at the end of detection
                MepIntersectionService.ClearGeometryCache();
            }
            catch { }
            return newClashZones;
        }

        /// <summary>
        /// ✅ PERFORMANCE OPTIMIZED: Streamlined clash zone creation for validated intersections
        /// Skips redundant validation, penetration checks, and damper detection since intersections 
        /// from MepIntersectionService are already validated and filtered.
        /// Expected: 10-20x faster than legacy path (10ms vs 189ms per zone)
        /// </summary>
        private List<ClashZone> DetectNewClashZonesStreamlined(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            Dictionary<string, double> clearanceSettings,
            List<string> selectedCategories)
        {
            var allProcessedZones = new List<ClashZone>();

            try
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                int newCount = 0;
                int updatedCount = 0;

                var documentPath = document.PathName;
                var documentHash = CalculateDocumentHash(document);

                // ✅ CACHES: Declared at method scope for use across multiple TrackOperation blocks
                Dictionary<(long mepId, long hostId, string pointKey), Guid> preFetchedGuids = new Dictionary<(long, long, string), Guid>();
                var openingLocationMap = new Dictionary<string, FamilyInstance>();
                var existingSleeveIds = new HashSet<int>();
                var mepParamsCache = new Dictionary<int, Dictionary<string, string>>();
                var hostParamsCache = new Dictionary<int, Dictionary<string, string>>();
                var orientationCache = new Dictionary<int, XYZ>();
                var hostDirectionCache = new Dictionary<int, (XYZ wallDirection, string wallOrientation, double centerCoord, bool centerCoordValid)>();
                // ✅ LINKED FILE FIX: Use string keys (ElementId|DocPath) for cross-document uniqueness
                var hostDataCache = new Dictionary<string, HostElementComputedData>();
                var mepDataCache = new Dictionary<string, MepElementComputedData>();
                HashSet<string>? preCachedWhitelist = null;
                bool allZonesExisting = false;
                bool skipGuidDbLookup = false;

                // ✅ PHASE 3 OPTIMIZATION: O(1) zone lookup map - Build BEFORE cache prep to enable fast-path check
                var zoneMap = _clashZoneStorage.ClashZones.ToDictionary(cz => cz.Id);

                // ✅ THREAD-SAFETY FIX: Pre-initialize lookup cache BEFORE entering parallel processing
                // This prevents multiple threads from trying to build the cache concurrently (infinite loop risk)
                BuildClashZoneLookup();

                // ✅ FAST PATH CHECK: Pre-check if all zones already exist in memory
                // This allows us to skip expensive DB GUID lookup entirely for PATH 1
                if (zoneMap.Count >= currentIntersections.Count && currentIntersections.Count > 0)
                {
                    // Quick heuristic: if we have as many zones in memory as intersections,
                    // likely all zones exist (PATH 1 scenario)
                    _log($"[STREAMLINED] Fast-path check: {zoneMap.Count} zones in memory, {currentIntersections.Count} intersections");
                    skipGuidDbLookup = true;
                }

                // ✅ PHASE 1-2.5: Cache Preparation
                using (var cacheOp = _performanceMonitor?.TrackOperation("6d1. Pre-fetch Whitelist/GUIDs") as IDisposable)
                {
                    // ✅ PHASE 1 OPTIMIZATION: Pre-calculate parameter whitelist once per refresh
                    if (OptimizationFlags.UsePreCachedWhitelist)
                    {
                        var parameterSnapshotService = new ParameterSnapshotService();
                        // We use a sample of intersections to build the whitelist (e.g., first 5)
                        var sample = currentIntersections.Take(5).Select(i => (i.Item1, i.Item2));
                        preCachedWhitelist = parameterSnapshotService.BuildWhitelist(_clashZoneStorage, sample);
                        _log($"[STREAMLINED] Pre-cached parameter whitelist with {preCachedWhitelist.Count} keys");
                    }

                    // ✅ PHASE 1 OPTIMIZATION: Bulk pre-fetch deterministic GUIDs from database
                    // SKIPPED if skipGuidDbLookup=true (PATH 1 - all zones already in memory)
                    if (!skipGuidDbLookup && OptimizationFlags.UseBatchGuidLookup && _guidManager != null)
                    {
                        var targets = currentIntersections.Select(i => ((long)i.Item1.Id.GetIntegerValue(), (long)i.Item2.Id.GetIntegerValue(), i.Item4.X, i.Item4.Y, i.Item4.Z)).ToList();
                        preFetchedGuids = _guidManager.BatchFetchGuidsDatabaseFirst(targets);
                        _log($"[STREAMLINED] Pre-fetched {preFetchedGuids.Count} GUIDs from database");
                    }
                    else if (skipGuidDbLookup)
                    {
                        _log($"[STREAMLINED] ⚡ SKIPPED GUID DB lookup - fast path (all zones in memory)");
                    }
                }

                // PERF OPT: Single-pass collector — build existingSleeveIds AND openingLocationMap in one scan
                using (var itemOp = _performanceMonitor?.TrackOperation("6d2. Sleeve/Opening Indexing") as IDisposable)
                {
                    var swSleeve = System.Diagnostics.Stopwatch.StartNew();
                    var symbolFamilyNameCache = new Dictionary<ElementId, bool>(); // FamilySymbol.Id → isOpening

                    var allFamilyInstances = new FilteredElementCollector(document)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .ToList();

                    foreach (var fi in allFamilyInstances)
                    {
                        int intId = fi.Id.GetIntegerValue();
                        existingSleeveIds.Add(intId);

                        // Check if this is an Opening family — cache by Symbol Id to avoid repeat COM calls
                        var symId = fi.Symbol?.Id;
                        if (symId == null) continue;

                        if (!symbolFamilyNameCache.TryGetValue(symId, out bool isOpening))
                        {
                            isOpening = fi.Symbol?.Family?.Name?.Contains("Opening") == true;
                            symbolFamilyNameCache[symId] = isOpening;
                        }

                        if (!isOpening) continue;

                        XYZ loc = null;
                        if (fi.Location is LocationPoint lp) loc = lp.Point;
                        else if (fi.Location is LocationCurve lc) loc = lc.Curve.Evaluate(0.5, true);

                        if (loc != null)
                        {
                            string key = $"{Math.Round(loc.X, 3)}_{Math.Round(loc.Y, 3)}_{Math.Round(loc.Z, 3)}";
                            if (!openingLocationMap.ContainsKey(key)) openingLocationMap.Add(key, fi);
                        }
                    }
                    swSleeve.Stop();
                    _log($"[STREAMLINED] Single-pass: {existingSleeveIds.Count} sleeve IDs + {openingLocationMap.Count} openings indexed in {swSleeve.ElapsedMilliseconds}ms");
                }
                var openingPointKeys = openingLocationMap.Keys.ToHashSet();

                // ✅ PHASE 2-2.5: Cache Preparation (Continued)
                using (var cacheOp2 = _performanceMonitor?.TrackOperation("6d3. Bulk Cache Generation") as IDisposable)
                {
                    // FAST PATH CHECK: If ALL intersections already have GUIDs in DB and zones in memory,
                    // skip ALL expensive pre-computation — PATH 1/3 only need flag operations, zero COM calls.
                    // For fast-path (skipGuidDbLookup=true), use zoneMap directly to validate all zones exist
                    if (skipGuidDbLookup)
                    {
                        // Verify all intersections have matching zones in memory (by MEP+Host+Point key)
                        // ✅ PERF OPT: Use HashSet for O(N) lookup instead of O(N^2) Any() inside All()
                        var existingZoneKeys = new HashSet<string>(zoneMap.Values.Select(z => 
                            $"{z.MepElementIdValue}_{z.StructuralElementIdValue}_{Math.Round(z.IntersectionPointX, 2)}_{Math.Round(z.IntersectionPointY, 2)}_{Math.Round(z.IntersectionPointZ, 2)}"));

                        allZonesExisting = currentIntersections.All(intersection =>
                        {
                            var (mepEl, hostEl, _, point) = intersection;
                            var key = $"{mepEl.Id.GetIntegerValue()}_{hostEl.Id.GetIntegerValue()}_{Math.Round(point.X, 2)}_{Math.Round(point.Y, 2)}_{Math.Round(point.Z, 2)}";
                            return existingZoneKeys.Contains(key);
                        });

                        if (allZonesExisting)
                            _log($"[STREAMLINED] ⚡ FAST PATH VALIDATED: ALL {currentIntersections.Count} zones exist in memory — zero COM calls needed");
                        else
                            _log($"[STREAMLINED] ⚠ Fast path check failed - some zones missing, will use preFetchedGuids");
                    }
                    else
                    {
                        // Normal path: check against pre-fetched GUIDs
                        allZonesExisting = preFetchedGuids.Count == currentIntersections.Count
                                                && zoneMap.Count >= preFetchedGuids.Count;
                        if (allZonesExisting)
                            _log($"[STREAMLINED] ⚡ ALL {currentIntersections.Count} zones exist in DB+memory — skipping expensive pre-computation (mepParams/hostParams/orientation/mepData caches)");
                    }

                    // PERF OPT: Filter to only PATH 2 (new) intersections for all pre-computation.
                    // currentIntersections includes ALL zones (PATH 1+2+3), but PATH 1/3 zones are already
                    // in DB+memory — their elements don't need CaptureBatchParams / mepDataCache / hostDataCache.
                    // preFetchedGuids identifies which zones exist in DB: anything NOT in it is a new zone.
                    // Without this filter, 287 MEP + 17 host elements get processed even if only 35 are new.
                    var newIntersections = currentIntersections;
                    if (!allZonesExisting && preFetchedGuids.Count > 0 && preFetchedGuids.Count < currentIntersections.Count)
                    {
                        newIntersections = currentIntersections.Where(i =>
                        {
                            string pk = $"{Math.Round(i.Item4.X, 4)}_{Math.Round(i.Item4.Y, 4)}_{Math.Round(i.Item4.Z, 4)}";
                            return !preFetchedGuids.ContainsKey(((long)i.Item1.Id.GetIntegerValue(), (long)i.Item2.Id.GetIntegerValue(), pk));
                        }).ToList();
                        _log($"[STREAMLINED] ⚡ PATH 2 filter: {newIntersections.Count}/{currentIntersections.Count} need pre-computation (saved {currentIntersections.Count - newIntersections.Count} element lookups)");
                    }

                    // ✅ PHASE 2 OPTIMIZATION: Bulk parameter capture (Ducts/Pipes/Cable Trays)
                    using (var paramOp = _performanceMonitor?.TrackOperation("6d3a. Bulk Parameter Capture") as IDisposable)
                    {
                        if (!allZonesExisting && OptimizationFlags.UseBulkIntersectionProcessing)
                        {
                            var mepElementsUnique = newIntersections.Select(i => i.Item1).Distinct(new ElementIdComparer()).ToList();
                            var hostElementsUnique = newIntersections.Select(i => i.Item2).Distinct(new ElementIdComparer()).ToList();

                            var mepC = ParameterSnapshotService.CaptureBatchParams(mepElementsUnique);
                            var hostC = ParameterSnapshotService.CaptureBatchParams(hostElementsUnique);

                            foreach (var kv in mepC) mepParamsCache[kv.Key] = kv.Value;
                            foreach (var kv in hostC) hostParamsCache[kv.Key] = kv.Value;

                            _log($"[STREAMLINED] Batch captured parameters for {mepParamsCache.Count} MEP elements and {hostParamsCache.Count} host elements");
                        }
                    }

                    // ✅ PHASE 2.5 OPTIMIZATION: Bulk orientation pre-calculation
                    using (var orientOp = _performanceMonitor?.TrackOperation("6d3b. Orientation Caching") as IDisposable)
                    {
                        if (!allZonesExisting && OptimizationFlags.UseBulkOrientationCaching)
                        {
                            var swOrient = System.Diagnostics.Stopwatch.StartNew();
                            var uniqueMepElements = newIntersections
                                .Select(i => i.Item1)
                                .Distinct(new ElementIdComparer())
                                .ToList();
                            foreach (var mep in uniqueMepElements)
                            {
                                try
                                {
                                    orientationCache[mep.Id.GetIntegerValue()] = GetMepElementOrientation(mep);
                                }
                                catch
                                {
                                    orientationCache[mep.Id.GetIntegerValue()] = XYZ.BasisX;
                                }
                            }
                            swOrient.Stop();
                            _log($"[STREAMLINED] Pre-calculated {orientationCache.Count} orientations in {swOrient.ElapsedMilliseconds}ms");
                        }
                    }

                    // PERF OPT: Pre-cache host wall direction + orientation + centerline coordinate.
                    using (var hostOp = _performanceMonitor?.TrackOperation("6d3c. Host Direction Caching") as IDisposable)
                    {
                        if (!allZonesExisting)
                        {
                            var swHost = System.Diagnostics.Stopwatch.StartNew();
                            foreach (var hostEl in newIntersections.Select(i => i.Item2).Distinct(new ElementIdComparer()))
                            {
                                try
                                {
                                    var dir = WallDirectionService.GetWallDirection(hostEl);
                                    string ori = (dir != null)
                                        ? (Math.Abs(dir.X) > Math.Abs(dir.Y) ? "X" : "Y")
                                        : string.Empty;
                                    double centerCoord = 0.0;
                                    bool centerCoordValid = false;
                                    if (hostEl is Wall wallHost)
                                    {
                                        var invariant = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetWallInvariantData(wallHost, document);
                                        if (invariant.Success) { centerCoord = invariant.CenterCoordinate; centerCoordValid = true; }
                                    }
                                    hostDirectionCache[hostEl.Id.GetIntegerValue()] = (dir, ori, centerCoord, centerCoordValid);
                                }
                                catch
                                {
                                    hostDirectionCache[hostEl.Id.GetIntegerValue()] = (null, string.Empty, 0.0, false);
                                }
                            }
                            swHost.Stop();
                            _log($"[STREAMLINED] Pre-cached wall direction+orientation+centerline for {hostDirectionCache.Count} unique hosts in {swHost.ElapsedMilliseconds}ms");
                        }
                    }

                    // PERF OPT: Pre-compute host-element-intrinsic data (thickness, normal, type).
                    using (var hostDataOp = _performanceMonitor?.TrackOperation("6d3d. Host Data Caching") as IDisposable)
                    {
                        if (!allZonesExisting)
                        {
                            var swHostData = System.Diagnostics.Stopwatch.StartNew();
                            var hostElementsBeforeDistinct = newIntersections.Select(i => i.Item2).ToList();
                            var hostElementsUnique = hostElementsBeforeDistinct.Distinct(new ElementIdComparer()).ToList();
                            _log($"[CACHE-DIAGNOSTIC] Host elements: {hostElementsBeforeDistinct.Count} total, {hostElementsUnique.Count} unique after Distinct (from {newIntersections.Count} new zones)");
                            foreach (var hostEl in hostElementsUnique)
                            {
                                try
                                {
                                    hostParamsCache.TryGetValue(hostEl.Id.GetIntegerValue(), out var hostParamDict);
                                    var d = new HostElementComputedData();
                                    d.StructuralElementType = GetStructuralElementType(hostEl, hostParamDict);
                                    d.StructuralElementThickness = GetElementThickness(hostEl, hostParamDict);
                                    d.WallThickness = GetWallThickness(hostEl, hostParamDict);
                                    d.FramingThickness = GetFramingThickness(hostEl, hostParamDict);
                                    d.StructuralElementNormal = WallDirectionService.GetStructuralElementNormal(hostEl);
                                    d.HostOrientation = WallDirectionService.GetHostOrientation(hostEl);
                                    d.GeometryHash = $"{hostEl.Id}_{hostEl.Category?.Name ?? "Unknown"}";  // ✅ PATH 2 OPT
                                    hostDataCache[GetElementCacheKey(hostEl)] = d;
                                }
                                catch { }
                            }
                            swHostData.Stop();
                            _log($"[STREAMLINED] Pre-computed host element data for {hostDataCache.Count} unique hosts in {swHostData.ElapsedMilliseconds}ms");
                        }
                    }

                    // PERF OPT: Pre-compute all MEP-element-intrinsic data.
                    using (var mepDataOp = _performanceMonitor?.TrackOperation("6d3e. MEP Data Caching") as IDisposable)
                    {
                        if (!allZonesExisting)
                        {
                            var swMepData = System.Diagnostics.Stopwatch.StartNew();
                            var mepElementsBeforeDistinct = newIntersections.Select(i => i.Item1).ToList();
                            var mepElementsUnique = mepElementsBeforeDistinct.Distinct(new ElementIdComparer()).ToList();
                            _log($"[CACHE-DIAGNOSTIC] MEP elements: {mepElementsBeforeDistinct.Count} total, {mepElementsUnique.Count} unique after Distinct (from {newIntersections.Count} new zones)");
                            foreach (var mepEl in mepElementsUnique)
                            {
                                try
                                {
                                    mepParamsCache.TryGetValue(mepEl.Id.GetIntegerValue(), out var mepParamDict);
                                    mepDataCache[GetElementCacheKey(mepEl)] = ComputeMepElementData(mepEl, mepParamDict);
                                }
                                catch (Exception ex)
                                {
                                    _log($"[MEP-DATA-CACHE] Error pre-computing data for {mepEl.Id}: {ex.Message}");
                                }
                            }

                            swMepData.Stop();
                            _log($"[STREAMLINED] Pre-computed MEP intrinsic data for {mepDataCache.Count} unique MEP elements in {swMepData.ElapsedMilliseconds}ms");
                        }
                    }
                }

                // ✅ DIAGNOSTIC: Cache summary before processing
                _log($"[CACHE-SUMMARY] Intersections: {currentIntersections.Count}, MEPDataCache: {mepDataCache.Count}, HostDataCache: {hostDataCache.Count}, MepParams: {mepParamsCache.Count}, HostParams: {hostParamsCache.Count}");

                // ✅ PHASE 3: Loop Execution with Threshold-Based Parallel Processing
                using (var processOp = _performanceMonitor?.TrackOperation("6d4. Zone Processing") as IDisposable)
                {
                    (processOp as PerformanceMonitor.OperationTracker)?.SetItemCount(currentIntersections.Count);
                    const int PARALLEL_THRESHOLD = 100;

                    if (currentIntersections.Count >= PARALLEL_THRESHOLD)
                    {
                        // Use parallel processing for large datasets
                        ProcessZonesParallel(currentIntersections, document, clearanceSettings, allProcessedZones,
                            zoneMap, preFetchedGuids, existingSleeveIds, openingPointKeys, preCachedWhitelist,
                            mepParamsCache, hostParamsCache, orientationCache, hostDirectionCache, mepDataCache, hostDataCache,
                            ref newCount, ref updatedCount);
                    }
                    else
                    {
                        // Use single-threaded processing for small datasets
                        ProcessZonesSingleThreaded(currentIntersections, document, clearanceSettings, allProcessedZones,
                            zoneMap, preFetchedGuids, existingSleeveIds, openingPointKeys, preCachedWhitelist,
                            mepParamsCache, hostParamsCache, orientationCache, hostDirectionCache, mepDataCache, hostDataCache,
                            ref newCount, ref updatedCount);
                    }
                }

                // Update storage metadata
                _clashZoneStorage.LastUpdated = DateTime.Now;
                _clashZoneStorage.DocumentPath = documentPath;
                _clashZoneStorage.DocumentHash = documentHash;

                sw.Stop();
                _log($"[STREAMLINED] Completed: Total={allProcessedZones.Count} (New={newCount}, Updated={updatedCount}) in {sw.ElapsedMilliseconds}ms");

                return allProcessedZones;
            }
            catch (Exception ex)
            {
                // ✅ CRITICAL: Graceful failure - log error and return partial results
                _log?.Invoke($"[STREAMLINED] ❌ CRITICAL ERROR in DetectNewClashZonesStreamlined: {ex.Message}");
                _log?.Invoke($"[STREAMLINED] Stack trace: {ex.StackTrace}");

                // Return whatever zones were processed before the crash
                _log?.Invoke($"[STREAMLINED] Returning {allProcessedZones?.Count ?? 0} zones processed before error");
                return allProcessedZones ?? new List<ClashZone>();
            }
        }

        /// <summary>
        /// Single-threaded zone processing for small datasets (< 100 zones)
        /// Lower overhead, no thread synchronization needed
        /// </summary>
        private void ProcessZonesSingleThreaded(
            List<(Element MepElement, Element HostElement, BoundingBoxXYZ BoundingBox, XYZ IntersectionPoint)> intersections,
            Document document,
            Dictionary<string, double> clearanceSettings,
            List<ClashZone> allProcessedZones,
            Dictionary<Guid, ClashZone> zoneMap,
            Dictionary<(long mepId, long hostId, string pointKey), Guid> preFetchedGuids,
            HashSet<int> existingSleeveIds,
            HashSet<string> openingPointKeys,
            HashSet<string> preCachedWhitelist,
            Dictionary<int, Dictionary<string, string>> mepParamsCache,
            Dictionary<int, Dictionary<string, string>> hostParamsCache,
            Dictionary<int, XYZ> orientationCache,
            Dictionary<int, (XYZ wallDirection, string wallOrientation, double centerCoord, bool centerCoordValid)> hostDirectionCache,
            Dictionary<string, MepElementComputedData> mepDataCache,
            Dictionary<string, HostElementComputedData> hostDataCache,
            ref int newCount,
            ref int updatedCount)
        {
            _log($"[SINGLE-THREADED] Processing {intersections.Count} zones");

            foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in intersections)
            {
                var currentNewCount = 0;
                var currentUpdatedCount = 0;
                var localProcessed = new List<ClashZone>();

                ProcessSingleIntersection(mepElement, structuralElement, boundingBox, intersectionPoint, document,
                    clearanceSettings, localProcessed, zoneMap, preFetchedGuids, existingSleeveIds, openingPointKeys,
                    preCachedWhitelist, mepParamsCache, hostParamsCache, orientationCache, hostDirectionCache, 
                    mepDataCache, hostDataCache, ref currentNewCount, ref currentUpdatedCount);
                
                // Merge to main results
                allProcessedZones.AddRange(localProcessed);
                
                // ✅ SIDE-EFFECT: Add new zones to persistent storage
                foreach (var nz in localProcessed)
                {
                    if (currentNewCount > 0 && !_clashZoneStorage.ClashZones.Contains(nz))
                        _clashZoneStorage.ClashZones.Add(nz);
                }
                
                newCount += currentNewCount;
                updatedCount += currentUpdatedCount;
            }
        }

        /// <summary>
        /// Parallel zone processing for large datasets (>= 100 zones)
        /// Better CPU utilization for multi-core systems
        /// </summary>
        private void ProcessZonesParallel(
            List<(Element MepElement, Element HostElement, BoundingBoxXYZ BoundingBox, XYZ IntersectionPoint)> intersections,
            Document document,
            Dictionary<string, double> clearanceSettings,
            List<ClashZone> allProcessedZones,
            Dictionary<Guid, ClashZone> zoneMap,
            Dictionary<(long mepId, long hostId, string pointKey), Guid> preFetchedGuids,
            HashSet<int> existingSleeveIds,
            HashSet<string> openingPointKeys,
            HashSet<string> preCachedWhitelist,
            Dictionary<int, Dictionary<string, string>> mepParamsCache,
            Dictionary<int, Dictionary<string, string>> hostParamsCache,
            Dictionary<int, XYZ> orientationCache,
            Dictionary<int, (XYZ wallDirection, string wallOrientation, double centerCoord, bool centerCoordValid)> hostDirectionCache,
            Dictionary<string, MepElementComputedData> mepDataCache,
            Dictionary<string, HostElementComputedData> hostDataCache,
            ref int newCount,
            ref int updatedCount)
        {
            _log($"[SEQUENTIAL] Processing {intersections.Count} zones sequentially for stability");

            List<ClashZone> newZones = new List<ClashZone>();
            List<ClashZone> updatedZones = new List<ClashZone>();
            List<int> newCounter = new List<int>();
            List<int> updatedCounter = new List<int>();
            System.Collections.Concurrent.ConcurrentBag<Exception> parallelExceptions = new System.Collections.Concurrent.ConcurrentBag<Exception>();

            foreach (var intersection in intersections)
            {
                try
                {
                    int localNewCount = 0;
                    int localUpdatedCount = 0;
                    List<ClashZone> localProcessed = new List<ClashZone>();

                    ProcessSingleIntersection(intersection.MepElement, intersection.HostElement, intersection.BoundingBox,
                        intersection.IntersectionPoint, document, clearanceSettings, localProcessed, zoneMap, preFetchedGuids,
                        existingSleeveIds, openingPointKeys, preCachedWhitelist, mepParamsCache, hostParamsCache,
                        orientationCache, hostDirectionCache, mepDataCache, hostDataCache, ref localNewCount, ref localUpdatedCount);

                    // Aggregate results
                    foreach (ClashZone zone in localProcessed)
                    {
                        if (localNewCount > 0)
                            newZones.Add(zone);
                        else
                            updatedZones.Add(zone);
                    }

                    if (localNewCount > 0) newCounter.Add(localNewCount);
                    if (localUpdatedCount > 0) updatedCounter.Add(localUpdatedCount);
                }
                catch (Exception ex)
                {
                    parallelExceptions.Add(new Exception($"MEP={intersection.MepElement?.Id}, Host={intersection.HostElement?.Id}: {ex.Message}", ex));
                }
            }
            
            // Log any exceptions that occurred during parallel processing
            if (!parallelExceptions.IsEmpty)
            {
                List<Exception> exceptionList = parallelExceptions.Take(10).ToList(); // Limit to first 10
                _log?.Invoke($"[PARALLEL] ⚠️ {parallelExceptions.Count} exceptions during processing. First few: {string.Join("; ", exceptionList.Select(e => e.Message))}");
                
                // If more than 10% failed, throw aggregate exception
                if (parallelExceptions.Count > intersections.Count / 10)
                {
                    throw new AggregateException("More than 10% of zones failed processing", parallelExceptions.ToArray());
                }
            }

            // Merge results back to main collections
            allProcessedZones.AddRange(newZones);
            allProcessedZones.AddRange(updatedZones);
            
            foreach (var zone in newZones)
            {
                if (!_clashZoneStorage.ClashZones.Contains(zone))
                    _clashZoneStorage.ClashZones.Add(zone);
            }

            newCount = newCounter.Sum();
            updatedCount = updatedCounter.Sum();
        }

        /// <summary>
        /// Process a single intersection - shared logic for both single-threaded and parallel paths
        /// </summary>
        private void ProcessSingleIntersection(
            Element mepElement, Element structuralElement, BoundingBoxXYZ boundingBox, XYZ intersectionPoint,
            Document document, Dictionary<string, double> clearanceSettings, List<ClashZone> processedZones,
            Dictionary<Guid, ClashZone> zoneMap, Dictionary<(long mepId, long hostId, string pointKey), Guid> preFetchedGuids,
            HashSet<int> existingSleeveIds, HashSet<string> openingPointKeys, HashSet<string> preCachedWhitelist,
            Dictionary<int, Dictionary<string, string>> mepParamsCache, Dictionary<int, Dictionary<string, string>> hostParamsCache,
            Dictionary<int, XYZ> orientationCache, Dictionary<int, (XYZ wallDirection, string wallOrientation, double centerCoord, bool centerCoordValid)> hostDirectionCache,
            Dictionary<string, MepElementComputedData> mepDataCache, Dictionary<string, HostElementComputedData> hostDataCache,
            ref int newCount, ref int updatedCount)
        {
            // Minimal validation - only check for null and invalid IDs
            if (mepElement == null || structuralElement == null ||
                mepElement.Id.GetIntegerValue() <= 0 || structuralElement.Id.GetIntegerValue() <= 0)
            {
                return;
            }

            // Check if clash zone already exists
            ClashZone? existingClashZone = null;

            // ✅ OPTIMIZATION: Try pre-fetched GUID lookup first
            if (OptimizationFlags.UseBatchGuidLookup)
            {
                string pointKey = $"{Math.Round(intersectionPoint.X, 4)}_{Math.Round(intersectionPoint.Y, 4)}_{Math.Round(intersectionPoint.Z, 4)}";
                if (preFetchedGuids.TryGetValue(((long)mepElement.Id.GetIntegerValue(), (long)structuralElement.Id.GetIntegerValue(), pointKey), out var guid))
                {
                    zoneMap.TryGetValue(guid, out existingClashZone);
                }
            }

            if (existingClashZone == null)
            {
                existingClashZone = FindExistingClashZone(mepElement.Id, structuralElement.Id, intersectionPoint);
            }

            if (existingClashZone == null)
            {
                // Create new clash zone
                try
                {
                    XYZ? calculatedCenterlineStreamlined = null;
                    string mepCategoryStreamlined = (mepDataCache != null && mepDataCache.TryGetValue(GetElementCacheKey(mepElement), out var mepDataForCat))
                        ? mepDataForCat.Category
                        : GetElementCategoryName(mepElement);

                    XYZ seedPoint = intersectionPoint;
                    if (mepCategoryStreamlined == "Duct Accessories" && mepElement is FamilyInstance fi)
                    {
                        if (fi.Location is LocationPoint lp)
                        {
                            seedPoint = lp.Point;
                        }
                    }

                    if (structuralElement is Wall wallStreamlined)
                    {
                        if (hostDirectionCache.TryGetValue(wallStreamlined.Id.GetIntegerValue(), out var hdc) && hdc.centerCoordValid)
                        {
                            bool isXWall = hdc.wallOrientation == "X";
                            calculatedCenterlineStreamlined = isXWall
                                ? new XYZ(seedPoint.X, hdc.centerCoord, seedPoint.Z)
                                : new XYZ(hdc.centerCoord, seedPoint.Y, seedPoint.Z);
                        }
                        else
                        {
                            calculatedCenterlineStreamlined = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                                wallStreamlined, seedPoint, document);
                        }
                    }
                    else if (structuralElement is FamilyInstance framingInstanceStreamlined &&
                             framingInstanceStreamlined.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
                    {
                        calculatedCenterlineStreamlined = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                            structuralElement, seedPoint, document);
                    }

                    // ✅ CRITICAL: Generate deterministic GUID for stable identification across detection runs
                    // This ensures the same intersection always gets the same GUID, preventing duplicates
                    Guid? preGuid = null;
                    if (_guidManager != null && OptimizationFlags.UseBatchGuidLookup)
                    {
                        preGuid = _guidManager.GenerateDeterministicGuid(
                            (long)mepElement.Id.GetIntegerValue(),
                            (long)structuralElement.Id.GetIntegerValue(),
                            intersectionPoint.X, intersectionPoint.Y, intersectionPoint.Z);
                    }

                    var newClashZone = CreateClashZone(mepElement, structuralElement, seedPoint, boundingBox, document,
                        clearanceSettings, null, null, calculatedCenterlineStreamlined, preCachedWhitelist, openingPointKeys,
                        mepParamsCache, hostParamsCache, orientationCache, hostDirectionCache, mepDataCache, preGuid, hostDataCache);

                    if (newClashZone != null)
                    {
                        newClashZone.IsCurrentClash = true;
                        newClashZone.ClearRevitApiObjects();
                        processedZones.Add(newClashZone);
                        // ✅ THREAD-SAFETY REMOVAL: Do not add to persistent list here (inside parallel loop)
                        // _clashZoneStorage.ClashZones.Add(newClashZone);
                        newCount++;
                    }
                }
                catch (Exception ex)
                {
                    _log?.Invoke($"[STREAMLINED] Error creating clash zone: MEP={mepElement.Id}, Structural={structuralElement.Id}, Error={ex.Message}");
                }
            }
            else
            {
                // Update existing zone flags
                if (!existingClashZone.IsResolved)
                {
                    string ptKey = $"{Math.Round(intersectionPoint.X, 3)}_{Math.Round(intersectionPoint.Y, 3)}_{Math.Round(intersectionPoint.Z, 3)}";
                    bool sleeveNowExists = (openingPointKeys != null) ? openingPointKeys.Contains(ptKey) : CheckForExistingSleeve(intersectionPoint, document);

                    if (sleeveNowExists)
                    {
                        existingClashZone.IsResolved = true;
                        if (existingClashZone.SleeveInstanceId <= 0) existingClashZone.SleeveInstanceId = -1;
                    }
                    if (sleeveNowExists && !existingClashZone.IsClusterResolved)
                    {
                        existingClashZone.IsClusterResolved = true;
                        if (existingClashZone.ClusterSleeveInstanceId <= 0) existingClashZone.ClusterSleeveInstanceId = -1;
                    }
                    if (sleeveNowExists && !existingClashZone.IsCombinedResolved)
                        existingClashZone.IsCombinedResolved = true;

                    existingClashZone.ReadyForPlacement = true;
                    existingClashZone.LastUpdated = DateTime.Now;
                }
                else
                {
                    bool sleeveActuallyExists = false;
                    if (existingClashZone.SleeveInstanceId > 0 && existingSleeveIds.Contains((int)existingClashZone.SleeveInstanceId))
                        sleeveActuallyExists = true;
                    else if (existingClashZone.ClusterSleeveInstanceId > 0 && existingSleeveIds.Contains((int)existingClashZone.ClusterSleeveInstanceId))
                        sleeveActuallyExists = true;

                    if (!sleeveActuallyExists)
                    {
                        existingClashZone.IsResolved = false;
                        existingClashZone.IsClusterResolved = false;
                        existingClashZone.IsCombinedResolved = false;
                        if (existingClashZone.SleeveInstanceId > 0) existingClashZone.SleeveInstanceId = 0;
                        if (existingClashZone.ClusterSleeveInstanceId > 0) existingClashZone.ClusterSleeveInstanceId = 0;
                        existingClashZone.ReadyForPlacement = true;
                        UpdateExistingClashZone(existingClashZone, mepElement, structuralElement, intersectionPoint,
                            boundingBox, document, preCachedWhitelist, existingSleeveIds, openingPointKeys,
                            mepParamsCache, hostParamsCache, mepDataCache);
                    }
                    else
                    {
                        existingClashZone.ReadyForPlacement = true;
                        existingClashZone.LastUpdated = DateTime.Now;
                    }
                }

                existingClashZone.IsCurrentClash = true;
                processedZones.Add(existingClashZone);
                updatedCount++;
            }
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
                                        _log($"Found MEP element {clashZone.MepElementId?.GetIntegerValue() ?? -1} in linked document {linkDoc.Title}");
                                        break;
                                    }
                                }
                                catch (Exception) // FIX: CS0168 - 'linkEx' commented to fix critical warning
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
                    _log($"Error getting MEP element {clashZone.MepElementId?.GetIntegerValue() ?? -1} for clash zone {clashZone.Id}: {ex.Message} - REMOVING invalid clash zone");
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
                var selectedHostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke() ?? new List<string>();
                
                if (selectedHostCategories.Count == 0)
                {
                    // If no host types selected, allow all (backward compatibility)
                    return true;
                }
                
                // Handle plural/singular mismatch: "Walls" (UI) vs "Wall" (Revit)
                bool hostTypeMatch = selectedHostCategories.Contains(clashZone.StructuralElementType) ||
                                   selectedHostCategories.Contains(clashZone.StructuralElementType + "s") ||
                                   selectedHostCategories.Any(t => t.TrimEnd('s').Equals(clashZone.StructuralElementType, StringComparison.OrdinalIgnoreCase));
                
                _log($"[HOST_TYPE_FILTER] ClashZone {clashZone.Id}: StructuralElementType='{clashZone.StructuralElementType}', SelectedHostTypes=[{string.Join(", ", selectedHostCategories)}], Match={hostTypeMatch}");
                
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
        private readonly ISectionBoxService? _sectionBoxService;
        private readonly System.Data.Common.DbConnection? _dbConnection;

        // Add ISectionBoxService and DbConnection to constructor
        public ClashZoneService(
            ClashZoneStorage? clashZoneStorage,
            Action<string>? log,
            Services.Interfaces.Refactor.IFlagManager? flagManager = null,
            GuidManager? guidManager = null,
            ISectionBoxService? sectionBoxService = null,
            System.Data.Common.DbConnection? dbConnection = null)
        {
            _clashZoneStorage = clashZoneStorage;
            _log = log;
            _flagManager = flagManager;
            _guidManager = guidManager;
            _sectionBoxService = sectionBoxService;
            _dbConnection = dbConnection;
        }

        /// <summary>
        /// Checks if a clash zone is visible in the current 3D section box (using cached DB value)
        /// </summary>
        private bool IsClashZoneVisibleInCurrentSectionBox(ClashZone clashZone, Document document)
        {
            try
            {
                var intersectionPoint = clashZone.IntersectionPoint;
                if (OptimizationFlags.UseSectionBoxCache)
                {
                    if (_sectionBoxService == null || _dbConnection == null)
                    {
                        _log($"SectionBoxService or DB connection not available - clash zone {clashZone.Id} considered visible (cache flag enabled, fallback)");
                        return true;
                    }
                    var sectionBox = _sectionBoxService.GetSectionBoxBounds(_dbConnection);
                    if (sectionBox == null)
                    {
                        _log($"No cached section box in DB - clash zone {clashZone.Id} considered visible (cache flag enabled, fallback)");
                        return true;
                    }
                    bool isVisible = intersectionPoint.X >= sectionBox.Min.X && intersectionPoint.X <= sectionBox.Max.X &&
                                     intersectionPoint.Y >= sectionBox.Min.Y && intersectionPoint.Y <= sectionBox.Max.Y &&
                                     intersectionPoint.Z >= sectionBox.Min.Z && intersectionPoint.Z <= sectionBox.Max.Z;
                    if (isVisible)
                        _log($"Clash zone {clashZone.Id} is visible in cached section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})");
                    else
                        _log($"Clash zone {clashZone.Id} is outside cached section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2})");
                    return isVisible;
                }
                else
                {
                    // Rollback: Use live section box from Revit API
                    if (!(document.ActiveView is View3D view3D) || !view3D.IsSectionBoxActive)
                    {
                        _log($"No active 3D section box - clash zone {clashZone.Id} considered visible (legacy mode)");
                        return true;
                    }
                    var sectionBox = Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                    if (sectionBox == null)
                    {
                        _log($"Could not get section box bounds - clash zone {clashZone.Id} considered visible (legacy mode)");
                        return true;
                    }
                    bool isVisible = intersectionPoint.X >= sectionBox.Min.X && intersectionPoint.X <= sectionBox.Max.X &&
                                     intersectionPoint.Y >= sectionBox.Min.Y && intersectionPoint.Y <= sectionBox.Max.Y &&
                                     intersectionPoint.Z >= sectionBox.Min.Z && intersectionPoint.Z <= sectionBox.Max.Z;
                    if (isVisible)
                        _log($"Clash zone {clashZone.Id} is visible in section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2}) (legacy mode)");
                    else
                        _log($"Clash zone {clashZone.Id} is outside section box at ({intersectionPoint.X:F2}, {intersectionPoint.Y:F2}, {intersectionPoint.Z:F2}) (legacy mode)");
                    return isVisible;
                }
            }
            catch (Exception ex)
            {
                _log($"Error checking clash zone {clashZone.Id} section box visibility: {ex.Message} - considering visible");
                return true;
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
        
        // Dictionary cache for O(1) lookup: (MepId, StructId) -> List of Candidates
        private Dictionary<(int, int), List<ClashZone>> _clashZoneLookup;
        
        /// <summary>
        /// Builds the lookup cache for fast existing clash zone retrieval.
        /// Call this ONCE before a batch processing run.
        /// </summary>
        public void BuildClashZoneLookup()
        {
            _clashZoneLookup = new Dictionary<(int, int), List<ClashZone>>();
            
            if (_clashZoneStorage?.ClashZones == null) return;
            
            foreach (var cz in _clashZoneStorage.ClashZones)
            {
                int mepId = cz.MepElementId?.GetIntegerValue() ?? (int)cz.MepElementIdValue;
                int structId = cz.StructuralElementId?.GetIntegerValue() ?? (int)cz.StructuralElementIdValue;
                
                var key = (mepId, structId);
                if (!_clashZoneLookup.ContainsKey(key))
                {
                    _clashZoneLookup[key] = new List<ClashZone>();
                }
                _clashZoneLookup[key].Add(cz);
            }
            
            if (OptimizationFlags.UseDiagnosticMode)
                _log($"[ClashZoneCache] Built lookup index with {_clashZoneLookup.Count} unique pairs from {_clashZoneStorage.ClashZones.Count} zones.");
        }

        private ClashZone? FindExistingClashZone(ElementId mepElementId, ElementId structuralElementId, XYZ intersectionPoint)
        {
            // ✅ CRITICAL FIX: Compare by IntegerValue to handle XML deserialization cases
            int mepIdValue = mepElementId?.GetIntegerValue() ?? -1;
            int structuralIdValue = structuralElementId?.GetIntegerValue() ?? -1;
            
            if (mepIdValue <= 0 || structuralIdValue <= 0)
            {
                if(OptimizationFlags.UseDiagnosticMode) 
                    _log($"[FindExistingClashZone] ❌ Invalid IDs - MEP={mepIdValue}, Structural={structuralIdValue}");
                return null;
            }

            // Lazy Build Cache if needed (safety net)
            if (_clashZoneLookup == null)
            {
                BuildClashZoneLookup();
            }

            // ✅ OPTIMIZED: O(1) Lookup
            var key = (mepIdValue, structuralIdValue);
            if (!_clashZoneLookup.TryGetValue(key, out var candidates))
            {
                // No candidates found for this pair
                if (OptimizationFlags.UseDiagnosticMode)
                {
                   // Too verbose for production
                   // _log($"[FindExistingClashZone] No existing zone found for MEP={mepIdValue}, Structural={structuralIdValue}");
                }
                return null;
            }

            // If candidates found, perform detailed check (Guid / Spatial)
            if (_guidManager != null && candidates.Count > 0)
            {
                // Priority 1: Check Revit sleeves for GUID
                // Note: GuidManager expects a full list usually, but we should pass only candidates to strict search if possible.
                // However, GuidManager might rely on scanning ALL zones if the pair IDs changed? 
                // Assuming ID stability: passing candidates is safe.
                // If GuidManager signatures require List<ClashZone>, we pass 'candidates'.
                
                var existingByRevitGuid = _guidManager.FindByRevitSleeveGuid(candidates, mepIdValue, structuralIdValue, intersectionPoint);
                if (existingByRevitGuid != null)
                {
                    if (OptimizationFlags.UseDiagnosticMode)
                        _log($"[FindExistingClashZone] ✅ FOUND VIA REVIT SLEEVE: ClashZone {existingByRevitGuid.Id} (read GUID from placed sleeve)");
                    return existingByRevitGuid;
                }
                
                // Priority 2: Check XML by MEP+Host+Point
                var existingByPoint = _guidManager.FindByMepHostAndPoint(candidates, mepIdValue, structuralIdValue, intersectionPoint);
                if (existingByPoint != null)
                {
                    if (OptimizationFlags.UseDiagnosticMode)
                        _log($"[FindExistingClashZone] ✅ FOUND VIA XML POINT MATCH: ClashZone {existingByPoint.Id}");
                    return existingByPoint;
                }
            }
            
            // ✅ FALLBACK: Legacy Logic (First match in candidates)
            // Since candidates are already filtered by MepId & StructId, we just take the first one 
            // (or logic to pick best if multiple, e.g. closest point?)
            // Legacy just took the first match.
            var legacyMatch = candidates.FirstOrDefault();

            if (legacyMatch != null)
            {
                if (OptimizationFlags.UseDiagnosticMode)
                    _log($"[FindExistingClashZone] ✅ FOUND (LEGACY): ClashZone {legacyMatch.Id}");
            }
            
            return legacyMatch;
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
        /// ⚠️ LEGACY METHOD - DEPRECATED (kept for backward compatibility only)
        /// 
        /// This method is replaced by FlagManager.ResetFlagsForDeletedSleeves() in the OOP refactoring.
        /// It is only called if FlagManager is not available.
        /// 
        /// Reset IsResolved flag for clash zones where sleeves no longer exist
        /// This allows re-placement of sleeves after manual deletion
        /// Called during refresh to detect deleted sleeves
        /// ⚠️ CRITICAL FIX: Only reset flags for clash zones in current UI context (selected filters/categories)
        /// </summary>
        /// <remarks>
        /// TODO: This method can be removed once all callers are migrated to use FlagManager
        /// </remarks>
        private void ResetResolvedFlagForDeletedSleeves(Document document, List<string>? selectedCategories = null)
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
                
                // ✅ PERFORMANCE OPTIMIZATION: Batch retrieve elements BEFORE loop to avoid duplicate GetElement calls
                // Pre-retrieve all unique MEP and structural element IDs from duct clash zones (only if needed)
                var ductClashZones = clashZonesToCheck
                    .Where(cz => string.Equals(cz.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) && 
                                (cz.IsResolved || cz.IsClusterResolved))
                    .ToList();
                
                var elementCache = new Dictionary<ElementId, Element>();
                if (ductClashZones.Count > 0)
                {
                    var mepElementIds = ductClashZones.Select(cz => cz.MepElementId).Distinct().Where(id => id != null && id != ElementId.InvalidElementId).ToList();
                    var structuralElementIds = ductClashZones.Select(cz => cz.StructuralElementId).Distinct().Where(id => id != null && id != ElementId.InvalidElementId).ToList();
                    
                    foreach (var mepId in mepElementIds)
                    {
                        var element = document.GetElement(mepId);
                        if (element != null) elementCache[mepId] = element;
                    }
                    foreach (var structId in structuralElementIds)
                    {
                        var element = document.GetElement(structId);
                        if (element != null) elementCache[structId] = element;
                    }
                }
                
                foreach (var clashZone in clashZonesToCheck)
                {
                    // ✅ FIX: Check BOTH individual and cluster sleeve flags
                    bool needsIndividualCheck = clashZone.IsResolved;
                    bool needsClusterCheck = clashZone.IsClusterResolved;
                    
                    if (needsIndividualCheck || needsClusterCheck)
                    {
                        // ✅ CRITICAL: Check Global XML FIRST before resetting flags
                        // If Global XML says sleeve exists, trust it even if Revit check fails (sleeve might be in linked file)
                        bool globalSaysResolved = false;
                        bool globalSaysClusterResolved = false;
                        // bool globalEntryFound = false; // FIX: CS0219 - commented to fix critical warning
                        // ✅ LEGACY XML REMOVED: GlobalIndexService checks removed.
                        // Relying solely on Revit model state to determine if sleeves exist.
                        
                        
                        // ✅ METHOD 3: Check for damper presence before resetting duct clash zones
                        if (string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase))
                        {
                            // ✅ PERFORMANCE OPTIMIZATION: Use pre-cached elements to avoid duplicate GetElement calls
                            Element mepElement = null;
                            Element structuralElement = null;
                            
                            if (elementCache.TryGetValue(clashZone.MepElementId, out var cachedMepElement))
                                mepElement = cachedMepElement;
                            else
                                mepElement = document.GetElement(clashZone.MepElementId);
                            
                            if (elementCache.TryGetValue(clashZone.StructuralElementId, out var cachedStructElement))
                                structuralElement = cachedStructElement;
                            else
                                structuralElement = document.GetElement(clashZone.StructuralElementId);
                            
                            if (mepElement != null && structuralElement != null)
                            {
                                if (CheckForDamperAtDuctEnd(document, mepElement, structuralElement, clashZone.IntersectionPoint))
                                {
                                    _log($"[ResetResolvedFlag] [METHOD3] SKIP: Duct clash zone {clashZone.Id} - Damper found at duct end, keeping flags=true");
                                    continue; // Don't reset this clash zone - damper exists
                                }
                            }
                        }
                        
                        // Validate placement point exists
                        if (clashZone.SleevePlacementPoint == null)
                        {
                            _log($"[ResetResolvedFlag] WARNING: ClashZone {clashZone.Id} has null SleevePlacementPoint - skipping");
                            continue;
                        }
                        
                        // ✅ CRITICAL FIX: Check cluster FIRST (per flag hierarchy - cluster flags take precedence)
                        // According to RELIABLE_FLAG_MANAGEMENT_IMPLEMENTATION.md: "Cluster flags take precedence over individual flags"
                        // FLOW: 1) Check cluster flag, 2) If cluster flag=true, check cluster sleeve in Revit, 3) If cluster flag=false, check individual flag, 4) If individual flag=true, check individual sleeve in Revit
        
                        if (needsClusterCheck)
                        {
                            // ✅ STEP 1: Cluster flag is TRUE - check cluster sleeve existence in Revit
                            _log($"[ResetResolvedFlag] ClashZone {clashZone.Id} has IsClusterResolved=true - checking cluster sleeve existence in Revit");
                            
                            // ✅ CRITICAL: If Global XML says cluster resolved, trust it (don't reset based on Revit check alone)
                            if (globalSaysClusterResolved)
                            {
                                _log($"[ResetResolvedFlag] ✅ SKIP RESET: ClashZone {clashZone.Id} - Global XML says IsClusterResolved=true, trusting Global XML (sleeve may be in linked file)");
                                continue; // Don't reset - Global XML is authoritative
                        }
                        
                            // ✅ STEP 2: Check if cluster sleeve exists in Revit by ClusterSleeveInstanceId
                        bool clusterSleeveExists = false;
                            if (clashZone.ClusterSleeveInstanceId > 0)
                        {
                            var clusterSleeve = document.GetElement(ElementIdCompat.FromValue(clashZone.ClusterSleeveInstanceId));
                            clusterSleeveExists = clusterSleeve != null;
                            
                                _log($"[ResetResolvedFlag] Checking cluster sleeve by ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} for clash zone {clashZone.Id} ({clashZone.MepElementCategory}): exists={clusterSleeveExists}");
                                
                                if (clusterSleeveExists)
                                {
                                    // ✅ Cluster sleeve found in Revit - skip reset
                                    _log($"[ResetResolvedFlag] ✅ SKIP RESET: Cluster sleeve found in Revit (ID: {clashZone.ClusterSleeveInstanceId}) - keeping IsClusterResolved=true, skipping individual check");
                                    continue; // Cluster exists - don't reset, don't check individual
                                }
                                else
                                {
                                    // ❌ Cluster sleeve NOT found in Revit - reset ALL flags
                                    _log($"[ResetResolvedFlag] ❌ Cluster sleeve NOT found in Revit (ID: {clashZone.ClusterSleeveInstanceId}) - resetting ALL flags");
                                }
                            }
                            else
                            {
                                // ❌ BUG CASE: IsClusterResolved=true but ClusterSleeveInstanceId invalid
                                // Inconsistent state - cluster sleeve is missing
                                clusterSleeveExists = false;
                                _log($"[ResetResolvedFlag] ❌ Cluster flag is true but ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} (invalid) - resetting ALL flags");
                            }
                            
                            // ✅ STEP 3: Reset ALL flags because cluster sleeve is missing
                            if (!clusterSleeveExists)
                            {
                                clashZone.IsClusterResolved = false;
                                clashZone.ClusterSleeveInstanceId = -1;
                                clashZone.IsResolved = false; // Reset individual flag too (cluster deletion means individual was also deleted)
                                clashZone.SleeveInstanceId = -1;
                                clashZone.SleeveFamilyName = string.Empty;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [RESET-ALL] ClashZone {clashZone.Id}: Cluster sleeve deleted/invalid, ALL flags reset\n");
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [RESET-ALL] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}\n");
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [RESET-ALL] PARAMS: SleeveInstanceId={clashZone.SleeveInstanceId}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                                }
                                
                                clashZone.LastUpdated = DateTime.Now;
                                resetCount++;
                                _log($"[ResetResolvedFlag] ✓ Reset ALL flags to FALSE for clash zone {clashZone.Id} ({clashZone.MepElementCategory}) - cluster sleeve deleted/invalid");
                                continue; // Done with this clash zone - don't check individual
                            }
                        }
                        else if (needsIndividualCheck)
                        {
                            // ✅ STEP 1: Individual flag is TRUE (cluster flag was false) - check individual sleeve existence in Revit
                            _log($"[ResetResolvedFlag] ClashZone {clashZone.Id} has IsResolved=true (cluster flag is false) - checking individual sleeve existence in Revit");
                            
                            // ✅ CRITICAL: If Global XML says individual resolved, trust it (don't reset based on Revit check alone)
                            if (globalSaysResolved)
                            {
                                _log($"[ResetResolvedFlag] ✅ SKIP RESET: ClashZone {clashZone.Id} - Global XML says IsResolved=true, trusting Global XML (sleeve may be in linked file)");
                                continue; // Don't reset - Global XML is authoritative
                            }
                            
                            // ✅ STEP 2: Check if individual sleeve exists in Revit by SleeveInstanceId
                            if (clashZone.SleeveInstanceId > 0)
                            {
                                var individualSleeve = document.GetElement(ElementIdCompat.FromValue(clashZone.SleeveInstanceId));
                                bool individualSleeveExists = individualSleeve != null;
                                
                                _log($"[ResetResolvedFlag] Checking individual sleeve by SleeveInstanceId={clashZone.SleeveInstanceId} for clash zone {clashZone.Id} ({clashZone.MepElementCategory}): exists={individualSleeveExists}");
                                
                                if (individualSleeveExists)
                                {
                                    // ✅ Individual sleeve found in Revit - skip reset
                                    _log($"[ResetResolvedFlag] ✅ SKIP RESET: Individual sleeve found in Revit (ID: {clashZone.SleeveInstanceId}) - keeping IsResolved=true");
                                    continue; // Sleeve exists - don't reset
                            }
                            else
                            {
                                    // ❌ Individual sleeve NOT found in Revit - reset flag
                                    _log($"[ResetResolvedFlag] ❌ Individual sleeve NOT found in Revit (ID: {clashZone.SleeveInstanceId}) - resetting individual flag");
                                    clashZone.IsResolved = false;
                                    clashZone.SleeveInstanceId = -1;
                                    clashZone.SleeveFamilyName = string.Empty;
                                    clashZone.LastUpdated = DateTime.Now;
                                    resetCount++;
                                    _log($"[ResetResolvedFlag] ✓ Reset individual flags to FALSE for clash zone {clashZone.Id} ({clashZone.MepElementCategory}) - sleeve NOT found in Revit");
                                }
                            }
                            else
                            {
                                // ❌ SleeveInstanceId is invalid (<=0) but flag is true - reset flag
                                _log($"[ResetResolvedFlag] ❌ Individual flag is true but SleeveInstanceId={clashZone.SleeveInstanceId} (invalid) - resetting individual flag");
                                clashZone.IsResolved = false;
                                clashZone.SleeveInstanceId = -1;
                                clashZone.SleeveFamilyName = string.Empty;
                                clashZone.LastUpdated = DateTime.Now;
                                resetCount++;
                                _log($"[ResetResolvedFlag] ✓ Reset individual flags to FALSE for clash zone {clashZone.Id} ({clashZone.MepElementCategory}) - invalid SleeveInstanceId");
                            }
                        }
                    }
                }
                
                if (resetCount > 0)
                {
                    _log($"[ResetResolvedFlag] Reset resolved flags (individual and/or cluster) for {resetCount} clash zones where sleeves were deleted (from selected categories only)");
                    
                    // ⚠️ CRITICAL: Log flag states AFTER reset operation completion
                    if (resetCount > 0)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [RESET-COMPLETE] Reset operation completed for {resetCount} clash zones\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [RESET-COMPLETE] Saving flags to both Global XML and Filter XML\n");
                        
                        // ✅ CRITICAL FIX: Save flags to Global XML after reset (legacy path)
                        // Note: FlagManager path handles this automatically, but legacy path needs manual save
                        // ✅ LEGACY XML REMOVED: No longer saving to Global XML here.
                        // FlagManager handles persistence now.
                    }
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
        
        /// <summary>
        /// ✅ HELPER: Calculate wall centerline point once before CreateClashZone (for ducts, pipes, cable trays)
        /// This avoids duplication - calculation happens here, passed to CreateClashZone to save in DB
        /// For dampers, this is calculated in DamperProcessingService using bbox method
        /// </summary>
        private XYZ? CalculateWallCenterlinePoint(Element structuralElement, XYZ intersectionPoint, Document document, string mepCategory)
        {
            try
            {
                // Only calculate for walls and framing (floors don't need centerline adjustment)
                if (structuralElement is Wall || 
                    (structuralElement is FamilyInstance framingInstance && 
                     framingInstance.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming))
                {
                    if (structuralElement is Wall hostWallForCenterline)
                    {
                        // ✅ LEGACY RAY-TRACE METHOD: For ducts, pipes, and cable trays, use ray-trace to find 2 wall faces and calculate midpoint
                        // This gives "half in and half out" positioning (legacy behavior)
                        var wallCenterlinePoint = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetWallCenterlinePointFromBbox(
                            hostWallForCenterline, 
                            intersectionPoint, 
                            document);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _log($"[ClashZoneService] ✅ Calculated Wall Centerline Point: ({wallCenterlinePoint.X:F3}, {wallCenterlinePoint.Y:F3}, {wallCenterlinePoint.Z:F3}) for {mepCategory}, wall {structuralElement.Id}");
                        }
                        return wallCenterlinePoint;
                    }
                    else
                    {
                        // For framing, use the element centerline method
                        var wallCenterlinePoint = JSE_RevitAddin_MEP_OPENINGS.Helpers.WallCenterlineHelper.GetElementCenterlinePoint(
                            structuralElement, 
                            intersectionPoint, 
                            document);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            _log($"[ClashZoneService] ✅ Calculated Framing Centerline Point: ({wallCenterlinePoint.X:F3}, {wallCenterlinePoint.Y:F3}, {wallCenterlinePoint.Z:F3}) for {mepCategory}, framing {structuralElement.Id}");
                        }
                        return wallCenterlinePoint;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    _log($"[ClashZoneService] ⚠️ Error calculating wall centerline point: {ex.Message}, will use intersection point as fallback");
                }
            }
            
            // Return null if calculation fails or not applicable - CreateClashZone will use intersection point as fallback
            return null;
        }
        
        private ClashZone CreateClashZone(Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document, Dictionary<string, double>? clearanceSettings = null, Dictionary<(double X, double Y, double Z), int>? spatialIndex = null, HashSet<int>? sleeveIds = null, XYZ? wallCenterlinePoint = null, HashSet<string>? parameterWhitelist = null, HashSet<string>? openingPointKeys = null, Dictionary<int, Dictionary<string, string>>? mepParamsCache = null, Dictionary<int, Dictionary<string, string>>? hostParamsCache = null, Dictionary<int, XYZ>? orientationCache = null, Dictionary<int, (XYZ wallDirection, string wallOrientation, double centerCoord, bool centerCoordValid)>? hostDirectionCache = null, Dictionary<string, MepElementComputedData>? mepDataCache = null, Guid? preComputedGuid = null, Dictionary<string, HostElementComputedData>? hostDataCache = null)
        {
            // ✅ DECLARATIONS AT METHOD SCOPE
            ClashZone? clashZone = null;
            System.Diagnostics.Stopwatch? swTotal = OptimizationFlags.EnableDetailedClashZoneProfiler ? System.Diagnostics.Stopwatch.StartNew() : null;
            System.Diagnostics.Stopwatch? swOp = OptimizationFlags.EnableDetailedClashZoneProfiler ? new System.Diagnostics.Stopwatch() : null;
            
            // ✅ CRITICAL DEBUG: Log cache keys to verify match between storage and lookup
            string mepKey = GetElementCacheKey(mepElement);
            string hostKey = GetElementCacheKey(structuralElement);
            
            MepElementComputedData? mepData = null;
            HostElementComputedData? hostData = null;
            // ✅ NOTE: actual cache lookup happens inside the timed block below (lines ~2950-2955); these are populated there
            
            Dictionary<string, string>? mepParamDict = null;
            if (mepParamsCache != null) mepParamsCache.TryGetValue(mepElement.Id.GetIntegerValue(), out mepParamDict);
            
            Dictionary<string, string>? hostParamDict = null;
            if (hostParamsCache != null) hostParamsCache.TryGetValue(structuralElement.Id.GetIntegerValue(), out hostParamDict);

            string mepCategory = "Unknown";
            string structuralElementType = "Unknown";
            double mepWidth = 0.0, mepHeight = 0.0;
            XYZ mepOrientation = XYZ.BasisX;
            XYZ wallDirection = XYZ.BasisX;
            string wallDirectionType = "Unknown";
            string mepCategoryForClearance = "Unknown";
            double finalWidth = 0.0, finalHeight = 0.0;
            string pipeOpeningType = "Unknown";
            string levelName = "Unknown";
            double levelElevation = 0.0;
            bool hasExistingSleeve = false;
            int existingSleeveId = -1;
            string wallOrientation = "Unknown";
            MepElementSize mepElementSize = null;
            string ductShape = "Unknown";
            bool isInsulated = false;
            double insulationThickness = 0.0;
            string insulationType = "Unknown";
            double finalDiameter = 0.0;
            XYZ finalPlacementPoint = intersectionPoint;
            List<SerializableKeyValue> mepParameterValues = new List<SerializableKeyValue>();
            List<SerializableKeyValue> hostParameterValues = new List<SerializableKeyValue>();
            bool hasMepConnector = false;
            string connectorSide = string.Empty;
            DamperConnectorInfo damperConnectorInfo = default;
            double elevationFromLevel = 0.0;
            double pipeOuterDiameter = 0.0;
            double pipeNominalDiameter = 0.0;
            string formattedSize = string.Empty;
            string sizeParameterValue = string.Empty;
            string systemAbbreviation = string.Empty;
            string mepSystemType = string.Empty;
            string mepSystemName = string.Empty;
            string mepServiceType = string.Empty;

            using (var internalOp = _performanceMonitor?.TrackOperation("6d4d1. Data Fetch & Logic") as IDisposable)
            {
                
                // Cache lookup
                if (swOp != null) swOp.Restart();
                mepParamDict = null;
                hostParamDict = null;
                
                if (mepParamsCache != null) mepParamsCache.TryGetValue(mepElement.Id.GetIntegerValue(), out mepParamDict);
                if (hostParamsCache != null) hostParamsCache.TryGetValue(structuralElement.Id.GetIntegerValue(), out hostParamDict);
                // PERF OPT: Look up pre-computed MEP element data (size, shape, insulation, system info…)
                mepData = null;
                mepDataCache?.TryGetValue(GetElementCacheKey(mepElement), out mepData);
                // PERF OPT: Look up pre-computed host structural data (thickness, normal, type, orientation)
                // Same host wall intersects many MEP elements — computed once per unique host in pre-batch block.
                hostData = null;
                hostDataCache?.TryGetValue(GetElementCacheKey(structuralElement), out hostData);
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] Cache lookup: {swOp.ElapsedMilliseconds}ms"); }

                // IMPORTANT: The intersection point is already at the wall center (mid-plane)
                // The MepIntersectionService finds intersections with wall faces and CreateBoundingBox()
                // averages entry/exit points, giving us the wall center automatically.
                // No additional offset needed since family insertion point is at middle.
                
                // DEBUG: Log placement point calculation
                if (OptimizationFlags.UseDiagnosticMode)
                {
                    _log($"[DEBUG] Placement Point Calculation for {structuralElement.Id}:");
                    _log($"[DEBUG]   Intersection Point: {intersectionPoint} (already at wall center - using as placement point)");
                }
                
                // Get structural element type — use pre-computed hostData cache to avoid COM call per zone
                if (swOp != null) swOp.Restart();
                structuralElementType = hostData?.StructuralElementType ?? GetStructuralElementType(structuralElement, hostParamDict);
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] GetStructuralElementType (cached={hostData != null}): {swOp.ElapsedMilliseconds}ms"); }
                
                if (OptimizationFlags.UseDiagnosticMode)
                    _log($"[DEBUG] StructuralElementType for {structuralElement.Id}: '{structuralElementType}' (Element: {structuralElement.GetType().Name})");
            }
            
            using (var metadataOp = _performanceMonitor?.TrackOperation("6d4d2. Metadata & Lookup") as IDisposable)
            {
                // Get document info
                var mepElementDoc = mepElement?.Document;
                var structuralElementDoc = structuralElement?.Document;
                
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("refresh.log", $"[CLASH-ZONE-CREATE] MEP Element {mepElement?.Id?.GetIntegerValue() ?? -1}: Document='{mepElementDoc?.Title ?? "null"}' (IsLinked={mepElementDoc != document})");
                    SafeFileLogger.SafeAppendText("refresh.log", $"[CLASH-ZONE-CREATE] Structural Element {structuralElement?.Id?.GetIntegerValue() ?? -1}: Document='{structuralElementDoc?.Title ?? "null"}' (IsLinked={structuralElementDoc != document})");
                }
                
                // Get MEP dimensions — from pre-computed cache if available
                if (swOp != null) swOp.Restart();
                var dims = (mepData != null)
                    ? (mepData.RawWidth, mepData.RawHeight)
                    : GetMepElementDimensions(mepElement, mepParamDict);
                mepWidth = dims.Item1;
                mepHeight = dims.Item2;
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] GetMepElementDimensions (cached={mepData != null}): {swOp.ElapsedMilliseconds}ms"); }
                
                // Get MEP orientation (with cache)
                if (swOp != null) swOp.Restart();
                if (orientationCache != null && orientationCache.TryGetValue(mepElement.Id.GetIntegerValue(), out var cachedOrientation))
                {
                    mepOrientation = cachedOrientation;
                    if (OptimizationFlags.UseDiagnosticMode)
                        _log($"[ORIENTATION-CACHE] Using cached orientation for {mepElement.Id}");
                }
                else
                {
                    mepOrientation = GetMepElementOrientation(mepElement);
                    if (OptimizationFlags.UseDiagnosticMode && orientationCache != null)
                        _log($"[ORIENTATION-CACHE] Cache miss for {mepElement.Id}, calculated on-demand");
                }
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] GetMepElementOrientation (or cache): {swOp.ElapsedMilliseconds}ms"); }
                
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLASH-ZONE-CREATE] MEP Orientation: ({mepOrientation.X:F6}, {mepOrientation.Y:F6}, {mepOrientation.Z:F6})");
                }
                
                // Get wall direction — use pre-cached value (same host can intersect many MEP elements)
                if (swOp != null) swOp.Restart();
                if (hostDirectionCache != null && hostDirectionCache.TryGetValue(structuralElement.Id.GetIntegerValue(), out var cachedHostDir))
                {
                    wallDirection = cachedHostDir.wallDirection;
                    // Derive wallDirectionType from cached direction (avoids GetWallDirectionType COM call)
                    wallDirectionType = (wallDirection != null)
                        ? (Math.Abs(wallDirection.X) > Math.Abs(wallDirection.Y) ? "X-WALL" : "Y-WALL")
                        : null;
                }
                else
                {
                    wallDirection = WallDirectionService.GetWallDirection(structuralElement);
                    wallDirectionType = WallDirectionService.GetWallDirectionType(structuralElement, wallDirection);
                }
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] WallDirectionService (cached={hostDirectionCache != null}): {swOp.ElapsedMilliseconds}ms"); }

                // Get category, pipe opening type, level info — all from pre-computed cache if available
                if (swOp != null) swOp.Restart();
                mepCategoryForClearance = (mepData != null) ? mepData.Category : GetElementCategoryName(mepElement, mepParamDict);
                finalWidth = mepWidth;
                finalHeight = mepHeight;
                pipeOpeningType = (mepData != null) ? mepData.PipeOpeningType : GetPipeOpeningType(mepElement);
                var levelInfo = (mepData != null)
                    ? (mepData.LevelName, mepData.LevelElevation)
                    : GetMepElementLevelInfo(mepElement);
                levelName = levelInfo.Item1;
                levelElevation = levelInfo.Item2;
                if (swOp != null) { swOp.Stop(); _log($"[PROFILER] Category+PipeOpeningType+LevelInfo (cached={mepData != null}): {swOp.ElapsedMilliseconds}ms"); }

                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id}: Raw dimensions {mepWidth:F3}x{mepHeight:F3} (clearance will be handled by CONDITIONS service during placement)");
                
                // ✅ OOP PATTERN: Two paths - optimized if spatial index provided, fallback if not
                hasExistingSleeve = false;
                existingSleeveId = -1;
                
                var swLookup = System.Diagnostics.Stopwatch.StartNew();
                if (spatialIndex != null && spatialIndex.Count > 0)
                {
                    // ✅ OPTIMIZED PATH: Use spatial index for O(1) lookup
                    var (found, sleeveId) = FindSleeveInSpatialIndex(intersectionPoint, spatialIndex);
                    hasExistingSleeve = found;
                    existingSleeveId = sleeveId;
                    
                    if (found && OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                        _log($"[OPTIMIZED-SLEEVE-LOOKUP] ✓ Found existing sleeve {sleeveId} at placement point using spatial index");
                }
                else if (openingPointKeys != null)
                {
                     // ✅ STREAMLINED OPTIMIZED PATH: Use pre-indexed opening keys for O(1) existence check
                     // NOTE: empty set is valid — means no sleeves placed yet, skip expensive FilteredElementCollector scan
                     string pointKey = $"{Math.Round(intersectionPoint.X, 3)}_{Math.Round(intersectionPoint.Y, 3)}_{Math.Round(intersectionPoint.Z, 3)}";
                     hasExistingSleeve = openingPointKeys.Contains(pointKey);
                     
                     if (hasExistingSleeve && OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                        _log($"[STREAMLINED-SLEEVE-LOOKUP] ✓ Found existing sleeve at {pointKey} using pre-indexed map");
                }
                else
                {
                    // ✅ FALLBACK PATH: Use old method when no sleeves exist for category
                    hasExistingSleeve = CheckForExistingSleeve(intersectionPoint, document);
                    
                    if (hasExistingSleeve && OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                        _log($"[LEGACY-SLEEVE-LOOKUP] ✓ Found existing sleeve at placement point using legacy scan");
                }
                
                // Old granular timing removed - replaced by EnableDetailedClashZoneProfiler
                
                // PERF OPT: Reuse mepCategoryForClearance (already computed above) — was calling GetElementCategoryName twice
                mepCategory = mepCategoryForClearance;
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id} ({mepElement.GetType().Name}): Category='{mepCategory}', Element.Category.Name='{mepElement.Category?.Name}'");
                
                // ✅ FIX: Skip pipe accessories when processing pipes filter
                if (mepCategory == "Pipe Accessories")
                {
                    if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] SKIP: Pipe Accessories element {mepElement.Id} - not processing unwanted clash zones");
                    return null; // Skip creating clash zone for pipe accessories
                }
                
                // ✅ WALL-AWARE DETECTION: Get wall orientation EARLY for BBox logic
                // PERF OPT: Use pre-cached orientation from hostDirectionCache (already computed per unique host)
                wallOrientation = (hostDirectionCache != null && hostDirectionCache.TryGetValue(structuralElement.Id.GetIntegerValue(), out var cachedOri))
                    ? cachedOri.wallOrientation
                    : (wallDirection != null ? (Math.Abs(wallDirection.X) > Math.Abs(wallDirection.Y) ? "X" : "Y") : WallDirectionService.GetHostOrientation(structuralElement));
            }

            using (var strategyOp = _performanceMonitor?.TrackOperation("6d4d3. Strategy & Insulation") as IDisposable)
            {
                // ✅ Size, shape, insulation — from pre-computed MEP cache if available
                var swSize = OptimizationFlags.UseDiagnosticMode ? System.Diagnostics.Stopwatch.StartNew() : null;
                if (mepData != null && mepData.StrategyValid)
                {
                    // PERF OPT: All from cache — zero COM calls
                    mepElementSize    = mepData.StrategySize;
                    finalWidth        = mepData.FinalWidth;
                    finalHeight       = mepData.FinalHeight;
                    finalDiameter     = mepData.FinalDiameter;
                    ductShape         = mepData.DuctShape;
                    isInsulated       = mepData.IsInsulated;
                    insulationThickness = mepData.InsulationThickness;
                    insulationType    = mepData.InsulationType;

                    if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id}: Dimensions from cache: {finalWidth:F3}x{finalHeight:F3}");
                }
                else
                {
                    // Live path: strategy not cached or strategy returned 0 (bbox fallback needed)
                    mepElementSize = (mepData != null) ? mepData.StrategySize : GetMepElementSizeWithStrategy(mepElement, mepCategoryForClearance);
                    finalDiameter = 0.0;

                    if (mepElementSize.Width > 0 || mepElementSize.Height > 0 || mepElementSize.Diameter > 0)
                    {
                        finalWidth    = mepElementSize.Width  > 0 ? mepElementSize.Width  : finalWidth;
                        finalHeight   = mepElementSize.Height > 0 ? mepElementSize.Height : finalHeight;
                        finalDiameter = mepElementSize.Diameter;
                        if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id}: Dimensions from Strategy: {finalWidth:F3}x{finalHeight:F3}");
                    }
                    else
                    {
                        // Generic bbox fallback (wall-orientation-dependent — cannot be pre-cached)
                        var bbox = mepElement.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            finalHeight = bbox.Max.Z - bbox.Min.Z;
                            finalWidth  = (wallOrientation == "X" || wallOrientation == "X-Wall")
                                ? bbox.Max.X - bbox.Min.X
                                : bbox.Max.Y - bbox.Min.Y;
                            if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[CLASH_DEBUG] Element {mepElement.Id}: Dimensions from BBox fallback: {finalWidth:F3}x{finalHeight:F3}");
                        }
                    }

                    ductShape           = (mepData != null) ? mepData.DuctShape : GetDuctShape(mepElement, mepParamDict);
                    (isInsulated, insulationThickness) = (mepData != null)
                        ? (mepData.IsInsulated, mepData.InsulationThickness)
                        : _insulationDetector.GetInsulationInfo(mepElement, mepElementSize);
                    insulationType = isInsulated ? "Insulated" : "Normal";
                }
                swSize?.Stop();

                // LOG GRANULAR TIMING — gated behind DiagnosticMode to avoid 384× disk writes
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                     SafeFileLogger.SafeAppendText("create_clashzone_perf.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] ID={mepElement.Id} " +
                        $"Size+Insul={swSize?.ElapsedMilliseconds}ms\n");
                }
            }
            
            using (var paramOp = _performanceMonitor?.TrackOperation("6d4d4. Parameters & Misc") as IDisposable)
            {
                // PERF OPT: All remaining MEP-intrinsic values from pre-computed cache if available
                pipeOuterDiameter   = mepData?.PipeOuterDiameter  ?? 0.0;
                pipeNominalDiameter = mepData?.PipeNominalDiameter ?? 0.0;

                if (mepData == null && mepElement is Pipe pipe)
                {
                    // Live fallback for pipe diameters when cache is absent
                    try
                    {
                        var outerDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        if (outerDiamParam != null) pipeOuterDiameter = outerDiamParam.AsDouble();
                        var nominalDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                        if (nominalDiamParam != null) pipeNominalDiameter = nominalDiamParam.AsDouble();
                    }
                    catch { }
                }

                // formattedSize: use cache if strategy was valid, else compute from resolved finalWidth/Height
                formattedSize = (mepData?.FormattedSize != null)
                    ? mepData.FormattedSize
                    : GetMepElementSizeString(mepElement, finalWidth, finalHeight, ductShape, pipeNominalDiameter);

                sizeParameterValue  = mepData?.SizeParameterValue  ?? GetMepElementSizeParameterValue(mepElement);
                systemAbbreviation  = mepData?.SystemAbbreviation   ?? GetMepSystemAbbreviation(mepElement);
                mepSystemType       = mepData?.SystemType           ?? GetMepSystemType(mepElement, mepCategoryForClearance);
                mepSystemName       = mepData?.SystemName           ?? string.Empty;
                if (mepData == null)
                {
                    try
                    {
                        var sysNameParam = mepElement?.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                        if (sysNameParam != null && sysNameParam.HasValue) mepSystemName = sysNameParam.AsString() ?? string.Empty;
                    }
                    catch { }
                }
                mepServiceType      = mepData?.ServiceType          ?? GetMepServiceType(mepElement);
                elevationFromLevel  = mepData?.ElevationFromLevel   ?? GetMepElementOffset(mepElement);
                
                // ✅ WALL-AWARE DETECTION: Wall orientation already retrieved above
                // string wallOrientation = WallDirectionService.GetHostOrientation(structuralElement);
                
                // ✅ DIAGNOSTIC: Log wall orientation for debugging
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode && mepCategory == "Duct Accessories")
                {
                    SafeFileLogger.SafeAppendText("damper_connector_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [ClashZoneService] WallOrientation='{wallOrientation ?? "NULL"}' for StructuralElement={structuralElement?.Id?.GetIntegerValue() ?? -1}, Type={structuralElement?.GetType()?.Name ?? "Unknown"}\n");
                    // ✅ BUILD STAMP near wall-orientation log for absolute certainty
                    try
                    {
                        var asm = typeof(ClashZoneService).Assembly;
                        string asmLoc = asm.Location;
                        var asmWrite = System.IO.File.Exists(asmLoc) ? System.IO.File.GetLastWriteTime(asmLoc) : DateTime.MinValue;
                        string asmVer = asm.GetName().Version?.ToString() ?? "unknown";
                        SafeFileLogger.SafeAppendText("damper_connector_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [BUILD-STAMP:CLASH] AssemblyLastWrite={asmWrite:yyyy-MM-dd HH:mm:ss}, Version={asmVer}, Assembly={asmLoc}\n");
                    }
                    catch { }
                }
                
                // ✅ OOP METHOD: Use DamperConnectorService to detect connector info (with wall orientation for wall-aware detection)
                damperConnectorInfo = GetDamperConnectorInfo(mepElement, mepCategory, wallOrientation, mepParamDict);
                connectorSide = damperConnectorInfo.ConnectorSide;
                hasMepConnector = damperConnectorInfo.HasMepConnector;
                // Note: Strategy will determine clearance from UI based on damper type and connector side
                
                // Placement point: intersection point is already at host center for FULL penetrations
                // MepIntersectionService.GetBoundingBoxCenter() averages entry/exit points to get mid-depth
                // NOTE: For partial penetrations that pass the 20% filter, intersection point is used as-is
                // (the working code doesn't have special handling for partial penetrations)
                // SAFETY: Some projects reported (0,0,0) due to missing transform at persistence time.
                // If we see a zero/near-zero point, fix up from the intersection bbox center or MEP bbox center.
                if (intersectionPoint == null || (Math.Abs(intersectionPoint.X) < 1e-9 && Math.Abs(intersectionPoint.Y) < 1e-9 && Math.Abs(intersectionPoint.Z) < 1e-9))
                {
                    try
                    {
                        XYZ fix = null;
                        if (boundingBox != null)
                        {
                            fix = new XYZ(
                                (boundingBox.Min.X + boundingBox.Max.X) / 2.0,
                                (boundingBox.Min.Y + boundingBox.Max.Y) / 2.0,
                                (boundingBox.Min.Z + boundingBox.Max.Z) / 2.0);
                        }
                        if (fix == null)
                        {
                            var mepBBoxFix = mepElement.get_BoundingBox(null);
                            if (mepBBoxFix != null)
                            {
                                fix = new XYZ(
                                    (mepBBoxFix.Min.X + mepBBoxFix.Max.X) / 2.0,
                                    (mepBBoxFix.Min.Y + mepBBoxFix.Max.Y) / 2.0,
                                    (mepBBoxFix.Min.Z + mepBBoxFix.Max.Z) / 2.0);
                            }
                        }
                        if (fix != null)
                        {
                            _log($"[FIXUP] IntersectionPoint was 0,0,0 → using center {fix} (bbox/meppbbox)");
                            intersectionPoint = fix;
                        }
                    }
                    catch { }
                }
                XYZ placementPoint = intersectionPoint;

                // ✅ FINAL PLACEMENT POINT: Use pre-calculated value passed from caller (calculated once before CreateClashZone using bbox method)
                // This is the final placement point at wall centerline, saved directly to SleevePlacementPoint
                // If not provided, fallback to intersection point (for backward compatibility or when calculation fails)
                // ✅ CRITICAL FIX: FOR FLOORS, ALWAYS USE INTERSECTION POINT (FloorCenterPoint concept)
                // WallCenterlinePoint is only relevant for Walls/Framing
                bool isFloorLoc = (structuralElementType == "Floor");
                finalPlacementPoint = (isFloorLoc || wallCenterlinePoint == null) ? intersectionPoint : wallCenterlinePoint;
                
                // ✅ DIAGNOSTIC: Log final placement point being set on ClashZone object
                // Gated behind DiagnosticMode to avoid ~350× disk writes per refresh
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode && wallCenterlinePoint != null)
                {
                    bool isZero = (finalPlacementPoint.X == 0.0 && finalPlacementPoint.Y == 0.0 && finalPlacementPoint.Z == 0.0);
                    string zeroWarning = isZero ? " ⚠️⚠️⚠️ ZERO VALUE!" : "";
                    SafeFileLogger.SafeAppendText("wall_centerline_set.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [SET] Zone (MEP={mepElement?.Id}, Category={mepCategory}): " +
                        $"Setting SleevePlacementPoint=({finalPlacementPoint.X:F6}ft, {finalPlacementPoint.Y:F6}ft, {finalPlacementPoint.Z:F6}ft), " +
                        $"Intersection=({intersectionPoint.X:F6}ft, {intersectionPoint.Y:F6}ft, {intersectionPoint.Z:F6}ft), " +
                        $"Source={(wallCenterlinePoint != null ? "Bbox Method" : "Fallback to Intersection")}{zeroWarning}\n");
                }

                // ✅ CRITICAL DEBUG: Verify intersection point is NOT zero before creating ClashZone
                bool isInputZero = Math.Abs(intersectionPoint.X) < 1e-9 && Math.Abs(intersectionPoint.Y) < 1e-9 && Math.Abs(intersectionPoint.Z) < 1e-9;
                if (isInputZero)
                {
                    _log($"[⚠️ CREATE-WARNING] IntersectionPoint is ZERO when creating ClashZone: MEP={mepElement.Id}, Structural={structuralElement.Id}, Point=({intersectionPoint.X},{intersectionPoint.Y},{intersectionPoint.Z})");
                    try 
                    { 
                        var debugPath = SafeFileLogger.GetLogFilePath("refresh_intersection_debug.log");
                        File.AppendAllText(debugPath, $"[{DateTime.Now}] ⚠️ CREATE-WARNING: IntersectionPoint ZERO - MEP={mepElement.Id}, Structural={structuralElement.Id}, Point=({intersectionPoint.X},{intersectionPoint.Y},{intersectionPoint.Z}), Doc={document?.Title}\n");
                    } 
                    catch { }
                }
                
                // ✅ NOTE: pipeOuterDiameter and pipeNominalDiameter are already extracted above (before formatted size calculation)
                
                // ✅ STAGE 1 (REFRESH): Capture MEP and Host parameters for ClashZone persistence
                // PERF OPT: Use pre-batch-captured cache (Dictionary<string,string>) when available —
                // avoids individual Revit parameter COM calls per zone.

                try
                {
                    // MEP params — from pre-batch cache if available
                    if (mepParamsCache != null && mepParamsCache.TryGetValue(mepElement.Id.GetIntegerValue(), out var cachedMepParams) && cachedMepParams != null)
                    {
                        mepParameterValues = cachedMepParams.Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList();
                    }
                    else if (_clashZoneStorage != null)
                    {
                        var parameterSnapshotService = new ParameterSnapshotService();
                        var whitelist = parameterWhitelist ?? parameterSnapshotService.BuildWhitelist(_clashZoneStorage, new[] { (mepElement, structuralElement) });
                        mepParameterValues = parameterSnapshotService.CaptureParams(mepElement, whitelist);
                    }

                    // Host params — from pre-batch cache if available
                    if (hostParamsCache != null && hostParamsCache.TryGetValue(structuralElement.Id.GetIntegerValue(), out var cachedHostParams) && cachedHostParams != null)
                    {
                        hostParameterValues = cachedHostParams.Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value }).ToList();
                    }
                    else if (_clashZoneStorage != null)
                    {
                        var parameterSnapshotService = new ParameterSnapshotService();
                        var whitelist = parameterWhitelist ?? parameterSnapshotService.BuildWhitelist(_clashZoneStorage, new[] { (mepElement, structuralElement) });
                        hostParameterValues = parameterSnapshotService.CaptureParams(structuralElement, whitelist);
                    }

                    if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode && (mepParameterValues.Count > 0 || hostParameterValues.Count > 0))
                        _log($"[PARAM-SNAPSHOT] Captured {mepParameterValues.Count} MEP parameters and {hostParameterValues.Count} host parameters for ClashZone (MEP={mepElement.Id}, Host={structuralElement.Id})");
                }
                catch (Exception ex)
                {
                    _log($"[PARAM-SNAPSHOT] ⚠️ Error capturing parameters: {ex.Message}");
                }
            }
            
            
            using (var finalOp = _performanceMonitor?.TrackOperation("6d4d5. Obj Init & GUID") as IDisposable)
            {
                // ✅ MEP METADATA: Extract type and family names for persistence
                // PERF OPT: Use pre-computed cache if available — avoids GetTypeId+GetElement COM calls per zone
                string? mepElementTypeName = mepData?.TypeName;
                string? mepElementFamilyName = mepData?.FamilyName;
                if (mepData == null)
                {
                    try
                    {
                        var typeId = mepElement?.GetTypeId();
                        if (typeId != null && typeId != ElementId.InvalidElementId)
                        {
                            var typeElem = mepElement?.Document?.GetElement(typeId);
                            if (typeElem != null)
                            {
                                mepElementTypeName = typeElem.Name;
                                if (typeElem is ElementType et) mepElementFamilyName = et.FamilyName;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                            _log($"[METADATA-ERROR] Failed to extract MEP type/family for ID={mepElement?.Id}: {ex.Message}");
                    }
                }

                // ✅ PERFORMANCE OPTIMIZATION: Use fast factory method instead of 70-property object initializer
                // Expected savings: ~40-50% faster object creation (~2ms → ~1ms per zone)
                clashZone = ClashZone.CreateFast(
                    // Core identifiers
                    id: preComputedGuid ?? Guid.NewGuid(),
                    mepElementId: mepElement.Id,
                    structuralElementId: structuralElement.Id,
                    // Intersection point
                    intersectionX: intersectionPoint.X,
                    intersectionY: intersectionPoint.Y,
                    intersectionZ: intersectionPoint.Z,
                    // Placement point
                    placementX: finalPlacementPoint.X,
                    placementY: finalPlacementPoint.Y,
                    placementZ: finalPlacementPoint.Z,
                    // Cached MEP data
                    mepCategory: MepCategoryConstants.Normalize(mepCategory),
                    mepWidth: finalWidth,
                    mepHeight: finalHeight,
                    mepOuterDiameter: pipeOuterDiameter,
                    mepNominalDiameter: pipeNominalDiameter,
                    mepFormattedSize: formattedSize,
                    mepSizeParameterValue: sizeParameterValue,
                    mepSystemAbbreviation: systemAbbreviation,
                    mepSystemType: mepSystemType,
                    mepSystemName: mepSystemName ?? string.Empty,
                    mepServiceType: mepServiceType,
                    mepTypeName: mepElementTypeName ?? string.Empty,
                    mepFamilyName: mepElementFamilyName ?? string.Empty,
                    mepUniqueId: mepData?.UniqueId ?? mepElement?.UniqueId ?? string.Empty,
                    mepOrientationX: mepOrientation.X,
                    mepOrientationY: mepOrientation.Y,
                    mepOrientationZ: mepOrientation.Z,
                    mepOrientationDirection: GetMepOrientationDirection(structuralElementType, mepOrientation, wallDirectionType),
                    mepRotationAngle: mepData?.MepRotationAngle ?? 
                        // ✅ ROTATION OPTIMIZATION: Only calculate for rectangular/non-round elements on Floors.
                        // Pipes and Conduits are rotation-invariant (round) so we skip the expensive calculation.
                        ((structuralElementType == "Floor" || structuralElementType == "Floors") && 
                         (mepCategory == "Ducts" || mepCategory == "Cable Trays" || mepCategory == "Duct Accessories")
                            ? CalculateMepElementRotationAngle(structuralElementType, mepOrientation, mepElement) 
                            : 0.0),
                    mepAngleToXRad: mepData?.MepAngleToXRad ?? 0.0,
                    mepAngleToXDeg: mepData?.MepAngleToXDeg ?? 0.0,
                    mepAngleToYRad: mepData?.MepAngleToYRad ?? 0.0,
                    mepAngleToYDeg: mepData?.MepAngleToYDeg ?? 0.0,
                    mepLevelName: levelName,
                    mepLevelElevation: levelElevation,
                    elevationFromLevel: elevationFromLevel,
                    // Cached Host data
                    structuralElementType: structuralElementType,
                    hostOrientation: hostData?.HostOrientation ?? GetHostOrientationWithTracking(structuralElement),
                    structuralThickness: hostData?.StructuralElementThickness ?? GetElementThickness(structuralElement),
                    wallThickness: hostData?.WallThickness ?? GetWallThickness(structuralElement),
                    framingThickness: hostData?.FramingThickness ?? GetFramingThickness(structuralElement),
                    // Insulation
                    isInsulated: isInsulated,
                    insulationThickness: insulationThickness,
                    insulationType: insulationType,
                    // Duct/Pipe
                    ductShape: ductShape,
                    pipeOpeningType: pipeOpeningType,
                    // Damper
                    hasMepConnector: hasMepConnector,
                    damperConnectorSide: connectorSide,
                    isStandardDamper: damperConnectorInfo.IsStandardDamper,
                    // Document keys
                    documentPath: document.PathName,
                    structuralDocTitle: structuralElement.Document.Title,
                    sourceDocKey: mepElement.Document.Title ?? mepElement.Document.PathName ?? string.Empty,
                    hostDocKey: structuralElement.Document.Title ?? structuralElement.Document.PathName ?? string.Empty,
                    // Geometry hashes (cached)
                    mepGeometryHash: mepData?.GeometryHash ?? CalculateElementGeometryHash(mepElement),
                    hostGeometryHash: hostData?.GeometryHash ?? CalculateElementGeometryHash(structuralElement),
                    // Flags
                    hasExistingSleeve: hasExistingSleeve,
                    // Bounding box
                    clashBoundingBox: boundingBox
                );
                
                // ✅ POST-FACTORY: Set parameter snapshots (not in fast factory to keep method signature manageable)
                clashZone.MepParameterValues = mepParameterValues;
                clashZone.HostParameterValues = hostParameterValues;
                clashZone.MepElementSizeData = mepElementSize;
                clashZone.WallDirection = wallDirection;
                clashZone.WallDirectionType = wallDirectionType;
                clashZone.MepElementOrientation = mepOrientation;
                clashZone.StructuralElementNormal = hostData?.StructuralElementNormal ?? WallDirectionService.GetStructuralElementNormal(structuralElement);
                clashZone.WallCenterlinePoint = finalPlacementPoint;
                clashZone.WallCenterlinePointX = finalPlacementPoint.X;
                clashZone.WallCenterlinePointY = finalPlacementPoint.Y;
                clashZone.WallCenterlinePointZ = finalPlacementPoint.Z;
                clashZone.IsMSFDDamper = hasMepConnector && !string.IsNullOrEmpty(connectorSide);
                
                // ✅ PERFORMANCE OPT: Diagnostic logging removed from hot path
                // File I/O was causing ~50-70ms per zone slowdown
                // Use PerformanceMonitor instead for timing data
                
                // ✅ CRITICAL: Set deterministic GUID for stable identification across detection runs
                // PERF OPT: Use pre-generated deterministic GUID (fast FNV-1a hash, no DB round-trip per zone).
                // The GUID was generated in ProcessSingleIntersection using fast FNV-1a hash.
                // It will be persisted to DB during the bulk save step (step 9).
                if (preComputedGuid.HasValue)
                {
                    clashZone.Id = preComputedGuid.Value;
                }
                // Note: For new zones, preComputedGuid should always be set by caller.
                // No fallback to DB lookup to avoid per-zone database round-trips.
                try
                {
                    var th = clashZone.StructuralElementThickness;
                    var thMm = RevitUnitConversionService.Instance.FromInternalMillimeters(th);
                    if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                            if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[CLASH-THICKNESS-ASSIGN] structuralId={structuralElement.Id.GetIntegerValue()} thickness={th:F6}ft ({thMm:F1}mm)");
                }
                catch { }
                
                // DEBUG: Log the pre-calculated data
                if (OptimizationFlags.UseDiagnosticMode)
                {
                    _log($"[DEBUG] Created ClashZone {clashZone.Id}:");
                    _log($"[DEBUG]   IntersectionPoint: {intersectionPoint} (used as placement point)");
                    if (intersectionPoint == null || (Math.Abs(intersectionPoint.X) < 1e-9 && Math.Abs(intersectionPoint.Y) < 1e-9 && Math.Abs(intersectionPoint.Z) < 1e-9))
                    {
                        _log($"[WARN]   IntersectionPoint is zero/invalid at save time. BBox null? {boundingBox == null}");
                    }
                    _log($"[DEBUG]   SleevePlacementPoint: {finalPlacementPoint} (calculated using bbox method at wall centerline)");
                    _log($"[DEBUG]   MepElementWidth: {finalWidth}, MepElementHeight: {finalHeight}");
                    _log($"[DEBUG]   MepElementOrientation: {mepOrientation}");
                    _log($"[DEBUG]   PipeOpeningType: {pipeOpeningType}");
                    _log($"[DEBUG]   IsResolved: {hasExistingSleeve}");
                }
                
                // ✅ CRITICAL: Immediately update Global XML when sleeve is found (optimized path)
                // This ensures flags are set correctly during detection, eliminating need for recovery
                if (hasExistingSleeve && existingSleeveId > 0 && _guidManager != null)
                {
                    try
                    {
                        // Update Global XML entry immediately with sleeve ID and resolved flag
                        var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId)>
                        {
                            (clashZone.Id, true, false, existingSleeveId, -1)
                        };
                        
                        // ✅ LEGACY XML REMOVED: GlobalIndexService is deleted.
                        // GlobalIndexService.UpsertFlagsWithIds(document, mepCategory, updates);
                        
                        if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                            _log($"[OPTIMIZED-SLEEVE-LOOKUP] ✓ Skipped Global XML update (legacy service deleted): Entry {clashZone.Id} → IsResolved=true, SleeveInstanceId={existingSleeveId}");
                    }
                    catch (Exception ex)
                    {
                        _log($"[OPTIMIZED-SLEEVE-LOOKUP] Error (legacy service deleted): {ex.Message}");
                    }
                }
            }
            
            return clashZone;
        }
        
        /// <summary>
        /// Get structural element type name for depth calculation
        /// </summary>
        // ✅ OOP REFACTORING: Removed GetHostOrientation() - now uses WallDirectionService.GetHostOrientation()

        private string GetStructuralElementType(Element element, Dictionary<string, string>? paramCache = null)
        {
            if (paramCache != null)
            {
                if (paramCache.TryGetValue("Category", out var cat))
                {
                    if (cat.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0) return "Wall";
                    if (cat.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0) return "Floor";
                    if (cat.IndexOf("Structural Framing", StringComparison.OrdinalIgnoreCase) >= 0) return "Structural Framing";
                }
            }

            if (element is Wall)
                return "Wall";
            else if (element is Floor)
                return "Floor";
            else if (element is FamilyInstance famInst && 
                     famInst.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
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
        private string GetElementCategoryName(Element element, Dictionary<string, string>? paramCache = null)
        {
            if (paramCache != null && paramCache.TryGetValue("Category", out var catName))
                return catName;

            try
            {
                // ⚠️ CRITICAL FIX: Check element type FIRST for core categories
                if (element is Autodesk.Revit.DB.Mechanical.Duct) return "Ducts";
                if (element is Autodesk.Revit.DB.Plumbing.Pipe) return "Pipes";
                if (element is Autodesk.Revit.DB.Electrical.CableTray) return "Cable Trays";
                if (element is Autodesk.Revit.DB.Electrical.Conduit) return "Conduits";
                
                // Fallback to Revit API Category name
                return element?.Category?.Name ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// Get MEP System Type for all categories (Pipes, Ducts, etc.)
        /// </summary>
        private string GetMepSystemType(Element mepElement, string category, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ PERF OPT: Try paramCache first
                if (paramCache != null)
                {
                    if (category == "Ducts" || category == "Duct Accessories")
                    {
                        if (paramCache.TryGetValue("System Type", out var dVal) && !string.IsNullOrEmpty(dVal)) return dVal;
                    }
                    else if (category == "Pipes" || category == "Pipe Accessories")
                    {
                        if (paramCache.TryGetValue("System Type", out var pVal) && !string.IsNullOrEmpty(pVal)) return pVal;
                    }
                    if (paramCache.TryGetValue("System Type", out var gVal) && !string.IsNullOrEmpty(gVal)) return gVal;
                    if (paramCache.TryGetValue("System Classification", out var cVal) && !string.IsNullOrEmpty(cVal)) return cVal;
                }

                // ✅ PERF OPT: Try paramCache first
                if (paramCache != null)
                {
                    if (paramCache.TryGetValue("System Type", out var stVal) && !string.IsNullOrEmpty(stVal)) return stVal;
                }

                // 1. Try Built-in Parameter (Fastest, language independent)
                var param = mepElement.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM); // Mechanical System
                if (param == null) param = mepElement.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM); // Piping System
                
                if (param != null) return param.AsValueString() ?? string.Empty;

                // 2. Try Generic "System Type" Parameter (Lookup)
                // Only if BIP failed (e.g. Electrical or obscure categories)
                var sysTypeParam = mepElement.LookupParameter("System Type");
                if (sysTypeParam != null) return sysTypeParam.AsValueString() ?? string.Empty;

                // 3. Try "System Classification" (Often used if System Type is missing)
                var classParam = mepElement.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                if (classParam != null) return classParam.AsValueString() ?? string.Empty;
                
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Get MEP Service Type (shared parameter) for all categories
        /// </summary>
        private string GetMepServiceType(Element mepElement, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ PERF OPT: Try paramCache first
                if (paramCache != null)
                {
                    if (paramCache.TryGetValue("Service Type", out var stVal) && !string.IsNullOrEmpty(stVal)) return stVal;
                    if (paramCache.TryGetValue("Service", out var sVal) && !string.IsNullOrEmpty(sVal)) return sVal;
                }

                // 1. Try "Service Type" (Common Shared Parameter)
                var param = mepElement.LookupParameter("Service Type");
                if (param != null) return param.AsString() ?? param.AsValueString() ?? string.Empty;

                // 2. Try "Service"
                var param2 = mepElement.LookupParameter("Service");
                if (param2 != null) return param2.AsString() ?? param2.AsValueString() ?? string.Empty;
                
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
/// Get MEP Element Offset/Elevation from Level — returns value in FEET (Revit internal units).
/// </summary>
private double GetMepElementOffset(Element mepElement, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
{
    try
    {
        // ✅ PERF OPT: Try paramCache first (captured in bulk)
        // All parameter values in Revit are stored in FEET (internal units)
        if (paramCache != null)
        {
            // 1. "Elevation from Level" (Standard BIP RBS_OFFSET_PARAM fallback name)
            if (paramCache.TryGetValue("Elevation from Level", out var eStr) && double.TryParse(eStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var eVal)) return eVal;
            
            // 2. "Offset" (Standard BIP name)
            if (paramCache.TryGetValue("Offset", out var oStr) && double.TryParse(oStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var oVal)) return oVal;
            
            // 3. "Middle Elevation" (Curve BIP name)
            if (paramCache.TryGetValue("Middle Elevation", out var mStr) && double.TryParse(mStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var mVal)) return mVal;

            // 4. "Reference Level Elevation" - Stored in feet (converted from mm during capture)
            if (paramCache.TryGetValue("Reference Level Elevation", out var rleStr) && double.TryParse(rleStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var rleVal)) return rleVal;
        }

        // Try BuiltIn RBS_OFFSET_PARAM ("Offset" / "Elevation from Level" — same BIP)
        var param = mepElement.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM);
        if (param != null && param.StorageType == StorageType.Double) return param.AsDouble();

        // Lookup fallbacks
        var param3 = mepElement.LookupParameter("Offset");
        if (param3 != null && param3.StorageType == StorageType.Double) return param3.AsDouble();

        var param4 = mepElement.LookupParameter("Middle Elevation");
        if (param4 != null && param4.StorageType == StorageType.Double) return param4.AsDouble();

        var param5 = mepElement.LookupParameter("Elevation from Level");
        if (param5 != null && param5.StorageType == StorageType.Double) return param5.AsDouble();

        return 0.0;
    }
    catch
    {
        return 0.0;
    }
}

        /// <summary>
        /// ✅ METHOD 3: Check for damper presence at duct end near intersection point
        /// This prevents creating clash zones for ducts that have dampers at their ends
        /// </summary>
        private bool CheckForDamperAtDuctEnd(Document document, Element ductElement, Element wallElement, XYZ intersectionPoint)
        {
            try
            {
                if (!(ductElement is Autodesk.Revit.DB.Mechanical.Duct duct))
                {
                    return false; // Not a duct, no need to check
                }

                _log($"[METHOD3] Checking for damper at duct end: Duct {duct.Id} near intersection {intersectionPoint}");

                // Get duct geometry and find end points
                var ductGeometry = duct.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                if (ductGeometry == null) return false;

                var ductEndPoints = GetDuctEndPoints(duct, ductGeometry);
                if (ductEndPoints.Count == 0) return false;

                _log($"[METHOD3] Found {ductEndPoints.Count} duct end points");

                // Check each end point for damper presence
                foreach (var endPoint in ductEndPoints)
                {
                    double distanceToIntersection = endPoint.DistanceTo(intersectionPoint);
                    
                    // Only check end points that are reasonably close to the intersection (within 2 feet)
                    if (distanceToIntersection < 2.0) // 2 feet = ~600mm
                    {
                        _log($"[METHOD3] Checking end point {endPoint} (distance to intersection: {distanceToIntersection:F3}ft)");
                        
                        // Search for dampers within radius of this end point
                        if (SearchForDampersNearPoint(document, endPoint, 0.5)) // 0.5 feet = ~150mm radius
                        {
                            _log($"[METHOD3] ✓ DAMPER FOUND: Damper detected near duct end point {endPoint}");
                            return true; // Damper found, skip this duct
                        }
                    }
                }

                _log($"[METHOD3] ✗ NO DAMPER: No damper found near any duct end points");
                return false; // No damper found, create clash zone normally
            }
            catch (Exception ex)
            {
                _log($"[METHOD3] ERROR: Failed to check for damper at duct end: {ex.Message}");
                return false; // On error, allow clash zone creation (fail safe)
            }
        }

        /// <summary>
        /// Get end points of a duct by analyzing its geometry
        /// </summary>
        private List<XYZ> GetDuctEndPoints(Autodesk.Revit.DB.Mechanical.Duct duct, GeometryElement ductGeometry)
        {
            var endPoints = new List<XYZ>();
            
            try
            {
                foreach (GeometryObject geomObj in ductGeometry)
                {
                    if (geomObj is Solid solid)
                    {
                        // Get the edges of the solid
                        foreach (Edge edge in solid.Edges)
                        {
                            var curve = edge.AsCurve();
                            if (curve != null)
                            {
                                // Add start and end points of each edge
                                endPoints.Add(curve.GetEndPoint(0));
                                endPoints.Add(curve.GetEndPoint(1));
                            }
                        }
                    }
                }

                // Remove duplicate points (within tolerance)
                var uniqueEndPoints = new List<XYZ>();
                const double tolerance = 0.01; // 1cm tolerance

                foreach (var point in endPoints)
                {
                    bool isDuplicate = uniqueEndPoints.Any(existing => existing.DistanceTo(point) < tolerance);
                    if (!isDuplicate)
                    {
                        uniqueEndPoints.Add(point);
                    }
                }

                _log($"[METHOD3] Extracted {uniqueEndPoints.Count} unique end points from duct geometry");
                return uniqueEndPoints;
            }
            catch (Exception ex)
            {
                _log($"[METHOD3] ERROR: Failed to get duct end points: {ex.Message}");
                return new List<XYZ>();
            }
        }

        /// <summary>
        /// Search for dampers within a specified radius of a point
        /// </summary>
        private bool SearchForDampersNearPoint(Document document, XYZ searchPoint, double searchRadius)
        {
            try
            {
                // Create a bounding box around the search point
                var searchBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(searchPoint.X - searchRadius, searchPoint.Y - searchRadius, searchPoint.Z - searchRadius),
                    Max = new XYZ(searchPoint.X + searchRadius, searchPoint.Y + searchRadius, searchPoint.Z + searchRadius)
                };

                // Create a filter for duct accessories (dampers)
                var categoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory);
                var boundingBoxFilter = new BoundingBoxIntersectsFilter(new Outline(searchBox.Min, searchBox.Max));
                var logicalAndFilter = new LogicalAndFilter(categoryFilter, boundingBoxFilter);

                // Search for duct accessories in the bounding box
                var damperCollector = new FilteredElementCollector(document)
                    .WherePasses(logicalAndFilter)
                    .WhereElementIsNotElementType();

                var dampers = damperCollector.ToList();
                
                _log($"[METHOD3] Found {dampers.Count} duct accessories within {searchRadius}ft of point {searchPoint}");

                // Check if any of these are actually dampers (fire dampers, volume dampers, etc.)
                foreach (var damper in dampers)
                {
                    var damperType = damper.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString();
                    var damperFamily = damper.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString();
                    
                    _log($"[METHOD3] Checking damper {damper.Id}: Type='{damperType}', Family='{damperFamily}'");
                    
                    // Check if this is a fire damper, volume damper, or other damper type
                    if (IsDamperType(damperType, damperFamily))
                    {
                        _log($"[METHOD3] ✓ CONFIRMED DAMPER: {damper.Id} is a damper (Type='{damperType}', Family='{damperFamily}')");
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _log($"[METHOD3] ERROR: Failed to search for dampers near point: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Check if an element is a damper based on its type and family names
        /// </summary>
        private bool IsDamperType(string typeName, string familyName)
        {
            if (string.IsNullOrEmpty(typeName) && string.IsNullOrEmpty(familyName))
                return false;

            var combinedName = $"{familyName} {typeName}".ToLowerInvariant();
            
            // Common damper keywords
            var damperKeywords = new[] { "damper", "fire", "volume", "control", "vav", "msfd", "fd" };
            
            bool isDamper = damperKeywords.Any(keyword => combinedName.Contains(keyword));
            
            if (isDamper)
            {
                _log($"[METHOD3] Damper detected: '{combinedName}' contains damper keywords");
            }
            
            return isDamper;
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
                
                if (insulationParam != null)
                {
                    double insulationValue = insulationParam.AsDouble();
                    if (insulationValue > 0.0)
                    {
                        double insulationMm = RevitUnitConversionService.Instance.FromInternalMillimeters(insulationValue);
                        _log($"[DEBUG] Element {element.Id} has insulation: {insulationMm:F1}mm");
                        return "Insulated";
                    }
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
        private string GetDuctShape(Element element, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ PERF OPT: Try paramCache first (captured in bulk)
                if (paramCache != null && paramCache.TryGetValue("Duct Shape", out var cachedShape) && !string.IsNullOrEmpty(cachedShape))
                {
                    return cachedShape;
                }
                if (!(element is Autodesk.Revit.DB.Mechanical.Duct))
                    return string.Empty; // Not a duct
                
                // Get duct type family name
                var duct = element as Autodesk.Revit.DB.Mechanical.Duct;
                var ductType = duct?.DuctType;
                var familyName = ductType?.FamilyName ?? string.Empty;
                
                if (OptimizationFlags.UseDiagnosticMode)
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
                double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
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
        /// Get MEP element size string (e.g., "12x10" or "D12") for description
        /// </summary>
        private string GetMepElementSizeString(Element mepElement, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ OPTIMIZED: Use paramCache directly if available (fastest)
                if (paramCache != null)
                {
                    // For Rectangular Ducts/Cable Trays
                    if (paramCache.TryGetValue("Width", out var wStr) && paramCache.TryGetValue("Height", out var hStr))
                    {
                        if (double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var w) && 
                            double.TryParse(hStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var h))
                        {
                            // Convert to mm for display
                            double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(w);
                            double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(h);
                            return $"{wMm:F0}x{hMm:F0}";
                        }
                    }
                    
                    // For Round Ducts/Pipes/Conduits
                    if (paramCache.TryGetValue("Diameter", out var dStr) || 
                        paramCache.TryGetValue("Nominal Diameter", out dStr) || 
                        paramCache.TryGetValue("Outside Diameter", out dStr))
                    {
                        if (double.TryParse(dStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                        {
                            double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(d);
                            return $"D{dMm:F0}";
                        }
                    }
                }
                
                // Fallback to dimensions method (which also tries cache but has API fallback)
                var (width, height) = GetMepElementDimensions(mepElement, paramCache);
                
                if (width <= 0 && height <= 0) return "Unknown";
                
                double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                
                if (Math.Abs(width - height) < 0.001) // Round
                {
                    return $"D{widthMm:F0}";
                }
                else // Rectangular
                {
                    return $"{widthMm:F0}x{heightMm:F0}";
                }
            }
            catch
            {
                return "Unknown";
            }
        }
        /// <summary>
        /// ✅ OPTIMIZED: Pre-collects sleeves by category and builds spatial index for O(1) lookup
        /// Uses "calculate once, use many times" strategy (same as recovery method)
        /// Returns spatial index Dictionary<(roundedX, roundedY, roundedZ), sleeveId> and HashSet of sleeve IDs
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="category">MEP element category name</param>
        /// <returns>Tuple: (spatialIndex Dictionary, sleeveIds HashSet, categorySleeves List)</returns>
        private (Dictionary<(double X, double Y, double Z), int> SpatialIndex, HashSet<int> SleeveIds, List<FamilyInstance> CategorySleeves) PreCollectSleevesByCategory(Document document, string category)
        {
            var spatialIndex = new Dictionary<(double X, double Y, double Z), int>();
            var sleeveIds = new HashSet<int>();
            var categorySleeves = new List<FamilyInstance>();
            
            try
            {
                // ✅ OPTIMIZATION: Filter sleeves by MEP_Category parameter (no expensive Revit API calls)
                // Same optimization as recovery method uses
                var allSleeves = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => 
                    {
                        // Check if it's a sleeve family
                        bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                               s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                        
                        string familyName = s.Symbol?.FamilyName ?? "";
                        bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                            familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        
                        if (!((s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                               (hasSleeveKeyword || isKnownFamily)))
                            return false;
                        
                        // ✅ PERFORMANCE: Filter by MEP_Category parameter (no Revit API call needed)
                        var mepCategoryParam = s.LookupParameter("MEP_Category");
                        if (mepCategoryParam != null && !string.IsNullOrWhiteSpace(mepCategoryParam.AsString()))
                        {
                            string sleeveCategory = mepCategoryParam.AsString();
                            return string.Equals(sleeveCategory, category, StringComparison.OrdinalIgnoreCase);
                        }
                        
                        // Fallback: If MEP_Category parameter is missing, skip (category unknown)
                        return false;
                    })
                    .ToList();
                
                categorySleeves = allSleeves;
                
                if (!DeploymentConfiguration.DeploymentMode)
                    _log($"[OPTIMIZED-SLEEVE-LOOKUP] Pre-collected {categorySleeves.Count} sleeves for category '{category}' (filtered by MEP_Category parameter)");
                
                // ✅ OPTIMIZATION: Build spatial index with 0.1ft tolerance (matches Global XML and recovery)
                double pointTolerance = 0.1; // 0.1ft = ~30mm tolerance (matches GlobalIndexService.FindByMepHostAndPoint)
                
                foreach (var sleeve in categorySleeves)
                {
                    int sleeveId = sleeve.Id.GetIntegerValue();
                    sleeveIds.Add(sleeveId);
                    
                    // Get sleeve location
                    XYZ sleeveLocation = null;
                    if (sleeve.Location is LocationPoint locationPoint)
                    {
                        sleeveLocation = locationPoint.Point;
                    }
                    else if (sleeve.Location is LocationCurve locationCurve)
                    {
                        // For LocationCurve, use the midpoint
                        sleeveLocation = locationCurve.Curve.Evaluate(0.5, true);
                    }
                    
                    if (sleeveLocation != null)
                    {
                        // Round to tolerance for spatial index key
                        double roundedX = Math.Round(sleeveLocation.X / pointTolerance) * pointTolerance;
                        double roundedY = Math.Round(sleeveLocation.Y / pointTolerance) * pointTolerance;
                        double roundedZ = Math.Round(sleeveLocation.Z / pointTolerance) * pointTolerance;
                        
                        var key = (roundedX, roundedY, roundedZ);
                        
                        // If multiple sleeves at same rounded point, keep the first one (shouldn't happen, but safe)
                        if (!spatialIndex.ContainsKey(key))
                        {
                            spatialIndex[key] = sleeveId;
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    _log($"[OPTIMIZED-SLEEVE-LOOKUP] Built spatial index with {spatialIndex.Count} entries for {categorySleeves.Count} sleeves");
            }
            catch (Exception ex)
            {
                _log($"[OPTIMIZED-SLEEVE-LOOKUP] Error pre-collecting sleeves: {ex.Message}");
            }
            
            return (spatialIndex, sleeveIds, categorySleeves);
        }
        
        /// <summary>
        /// ✅ OPTIMIZED: Finds sleeve in spatial index using O(1) lookup
        /// Returns (found, sleeveId) tuple
        /// </summary>
        /// <param name="placementPoint">Placement point to search for</param>
        /// <param name="spatialIndex">Pre-built spatial index</param>
        /// <returns>Tuple: (found bool, sleeveId int)</returns>
        private (bool Found, int SleeveId) FindSleeveInSpatialIndex(XYZ placementPoint, Dictionary<(double X, double Y, double Z), int> spatialIndex)
        {
            if (spatialIndex == null || spatialIndex.Count == 0)
                return (false, -1);
            
            try
            {
                // Round to same tolerance as spatial index (0.1ft)
                double pointTolerance = 0.1;
                double roundedX = Math.Round(placementPoint.X / pointTolerance) * pointTolerance;
                double roundedY = Math.Round(placementPoint.Y / pointTolerance) * pointTolerance;
                double roundedZ = Math.Round(placementPoint.Z / pointTolerance) * pointTolerance;
                
                var key = (roundedX, roundedY, roundedZ);
                
                if (spatialIndex.TryGetValue(key, out int sleeveId))
                {
                    return (true, sleeveId);
                }
            }
            catch (Exception ex)
            {
                _log($"[OPTIMIZED-SLEEVE-LOOKUP] Error finding sleeve in spatial index: {ex.Message}");
            }
            
            return (false, -1);
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Check if a sleeve exists at the EXACT placement point stored in XML
        /// This is essential for refresh to detect deleted sleeves and reset IsResolved flag
        /// Without this, deleted sleeves cannot be re-placed (duplication suppressor prevents it)
        /// FALLBACK PATH: Used when no sleeves exist for category (optimization not needed)
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
        private double GetElementThickness(Element element, Dictionary<string, string>? paramDict = null)
        {
            try
            {
                if (element is Wall || (element?.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_Walls))
                {
                    return GetWallThickness(element, paramDict);
                }
                else if (element is Floor floor)
                {
                    if (paramDict != null && paramDict.TryGetValue("Floor Thickness", out var thickStr) &&
                        double.TryParse(thickStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var thickVal))
                    {
                        return thickVal;
                    }
                    if (paramDict != null && paramDict.TryGetValue("Thickness", out var tStr) &&
                        double.TryParse(tStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var tVal))
                    {
                        return tVal;
                    }
                    return floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.1;
                }
                else if ((element?.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming))
                {
                    return GetFramingThickness(element, paramDict);
                }
                return 0.1; // Default fallback
            }
            catch
            {
                return 0.1; // Default fallback
            }
        }
        
        /// <summary>
        /// Get wall thickness (for walls only) with robust fallback for compound walls
        /// </summary>
        private double GetWallThickness(Element element, Dictionary<string, string>? paramDict = null)
        {
            if (element == null) return 0.0;
            
            try
            {
                // ✅ PERFORMANCE OPTIMIZATION: Check parameter cache first
                if (paramDict != null)
                {
                    if (paramDict.TryGetValue("Wall Width", out var wStr) && double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var wVal))
                        return wVal;
                    if (paramDict.TryGetValue("Width", out var widStr) && double.TryParse(widStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var widVal))
                        return widVal;
                    if (paramDict.TryGetValue("Thickness", out var thickStr) && double.TryParse(thickStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var thickVal))
                        return thickVal;
                }

                double thickness = 0.0;
                
                // Method 1: Direct Wall property if it's a Wall object
                if (element is Wall wall)
                {
                    thickness = wall.Width;
                }
                
                // Method 2: FALLBACK - If thickness is 0 or it's not a Wall object but in OST_Walls category
                // (Covers FaceWalls or other elements that might not cast to Wall but are walls)
                if (thickness <= 0.001)
                {
                    bool isWallCategory = element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_Walls;
                    if (isWallCategory || element is Wall)
                    {
                        // 1) Try Built-in parameter "Width" on Instance
                        Parameter p = element.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM) ?? 
                                     element.LookupParameter("Width") ??
                                     element.LookupParameter("Thickness");
                        
                        if (p != null && p.HasValue) 
                        {
                            thickness = p.AsDouble();
                        }
                        
                        // 2) Try Type parameter if instance failed/missing
                        if (thickness <= 0.001)
                        {
                            ElementId typeId = element.GetTypeId();
                            if (typeId != ElementId.InvalidElementId)
                            {
                                Element typeElem = element.Document.GetElement(typeId);
                                if (typeElem != null)
                                {
                                    Parameter tp = typeElem.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM) ?? 
                                                  typeElem.LookupParameter("Width") ??
                                                  typeElem.LookupParameter("Thickness");
                                    if (tp != null && tp.HasValue) 
                                    {
                                        thickness = tp.AsDouble();
                                    }
                                }
                            }
                        }
                    }
                }

                // If still 0, log error but return a minimal default to prevent 0.0 in DB
                if (thickness <= 0.001)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[WALL-THICKNESS] Wall {element.Id.GetIntegerValue()}: Could not determine thickness (returned 0.0). Category={element.Category?.Name}");
                    
                    return 0.1; // Minimal fallback (approx 30mm)
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[WALL-THICKNESS] Wall {element.Id.GetIntegerValue()}: thickness={RevitUnitConversionService.Instance.FromInternalMillimeters(thickness):F1}mm");
                
                return thickness;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GetWallThickness] Critical error for wall {element?.Id?.GetIntegerValue()}: {ex.Message}");
                return 0.1; // Safety fallback
            }
        }
        
        /// <summary>
        /// Get structural framing parameter 'b' thickness (for structural framing only)
        /// </summary>
        private double GetFramingThickness(Element element, Dictionary<string, string>? paramDict = null)
        {
            try
            {
                if ((element?.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming))
                {
                    // ✅ PERFORMANCE OPTIMIZATION: Check parameter cache first
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramDict != null)
                    {
                        string[] possibleNames = { "b", "B", "Breadth", "Depth", "Width", "Height", "d", "D" };
                        foreach (var name in possibleNames)
                        {
                            if (paramDict.TryGetValue(name, out var valStr) && double.TryParse(valStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                            {
                                if (val > 0) return val;
                            }
                        }
                    }

                    // Structural framing: read TYPE parameter 'b' (case-insensitive) regardless of instance/type wrapper
                    try
                    {
                        // Resolve the type element from any element (FamilyInstance or not)
                        ElementId typeId = ElementId.InvalidElementId;
                        try { typeId = (element as FamilyInstance)?.GetTypeId() ?? element.GetTypeId(); } catch { }
                        var typeElem = element.Document?.GetElement(typeId);

                        if (typeElem == null)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FRAMING-THICKNESS] Could not get type element for framing {element.Id.GetIntegerValue()}");
                            return 0.0;
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
                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FRAMING-THICKNESS] Found parameter '{paramName}' = {bVal:F6}ft ({RevitUnitConversionService.Instance.FromInternalMillimeters(bVal):F1}mm) on framing {element.Id.GetIntegerValue()}");
                                        break;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[FRAMING-THICKNESS] Error reading parameter '{paramName}': {ex.Message}");
                            }
                        }

                        // Log all available parameters for debugging if not found
                        if (bVal <= 0.0)
                        {
                            try
                            {
                                var parts = new System.Collections.Generic.List<string>();
                                foreach (Parameter tp in typeElem.Parameters)
                                {
                                    if (tp == null) continue;
                                    
                                    var name = tp.Definition?.Name ?? "<null>";
                                    string val = string.Empty;
                                    if (tp.StorageType == StorageType.Double)
                                    {
                                        double d = tp?.AsDouble() ?? 0.0;
                                        double mm = RevitUnitConversionService.Instance.FromInternalMillimeters(d);
                                        val = mm.ToString("F1") + "mm";
                                    }
                                    else
                                    {
                                        val = tp.AsString() ?? tp.AsValueString() ?? string.Empty;
                                    }
                                    parts.Add(name + ":" + val);
                                }
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("refresh.log", $"[FRAMING-THICKNESS-PARAMS] typeId={typeId.GetIntegerValue()}: {string.Join(", ", parts)}");
                                    }
                            }
                            catch (Exception ex)
                            {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        SafeFileLogger.SafeAppendText("refresh.log", $"[FRAMING-THICKNESS] Error logging parameters: {ex.Message}");
                                    }
                            }
                        }

                        return bVal > 0.0 ? bVal : 0.0;
                    }
                    catch (Exception ex)
                    {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("refresh.log", $"[FRAMING-THICKNESS] Error getting framing thickness: {ex.Message}");
                            }
                        return 0.0;
                    }
                }
                return 0.0; // Not structural framing
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("refresh.log", $"[GetFramingThickness] Error: {ex.Message}");
                }
                return 0.0;
            }
        }
        
        /// <summary>
        /// ✅ SIZE PARAMETER VALUE: Extract raw Size parameter value as string from MEP element
        /// This is the exact text from the Size parameter (e.g., "20 mmø", "200 mm dia symbol")
        /// Different from GetMepElementSizeString which may fall back to calculated format
        /// </summary>
        /// <param name="mepElement">The MEP element</param>
        /// <returns>Raw Size parameter value as string, or empty string if not available</returns>
        private string GetMepElementSizeParameterValue(Element mepElement, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                if (mepElement == null) return string.Empty;

                // ✅ PERF OPT: Trust paramCache completely during bulk refresh
                if (paramCache != null)
                {
                    if (paramCache.TryGetValue("Size", out var cachedSize) && !string.IsNullOrEmpty(cachedSize))
                    {
                        return cachedSize;
                    }
                    // Cache definitively proves it doesn't exist or is empty. Do not do a slow string-lookup fallback.
                    return string.Empty;
                }
                
                var sizeParam = mepElement.LookupParameter("Size");
                if (sizeParam == null)
                {
                    // ✅ DIAGNOSTIC: Log if Size parameter not found
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        SafeLog("Refresh_debug.log", $"[ClashZoneService] ⚠️ Size parameter not found for element {mepElement.Id.GetIntegerValue()}, Category={mepElement.Category?.Name}, Doc={mepElement.Document?.Title}");
                    return string.Empty;
                }
                
                // Try AsString() first (for string parameters)
                if (sizeParam.StorageType == StorageType.String)
                {
                    var sizeString = sizeParam.AsString();
                    if (!string.IsNullOrWhiteSpace(sizeString))
                    {
                        // ✅ DIAGNOSTIC: Log successful extraction
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            SafeLog("Refresh_debug.log", $"[ClashZoneService] ✅ Size parameter (String): Element {mepElement.Id.GetIntegerValue()}, Value='{sizeString}', Doc={mepElement.Document?.Title}");
                        return sizeString.Trim();
                    }
                }
                // Try AsValueString() as fallback (for numeric parameters that display as formatted strings)
                else
                {
                    var sizeString = sizeParam.AsValueString();
                    if (!string.IsNullOrWhiteSpace(sizeString))
                    {
                        // ✅ DIAGNOSTIC: Log successful extraction
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            SafeLog("Refresh_debug.log", $"[ClashZoneService] ✅ Size parameter (ValueString): Element {mepElement.Id.GetIntegerValue()}, Value='{sizeString}', StorageType={sizeParam.StorageType}, Doc={mepElement.Document?.Title}");
                        return sizeString.Trim();
                    }
                }
                
                // ✅ DIAGNOSTIC: Log if parameter exists but value is empty
                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    SafeLog("Refresh_debug.log", $"[ClashZoneService] ⚠️ Size parameter exists but value is empty for element {mepElement.Id.GetIntegerValue()}, StorageType={sizeParam.StorageType}, Doc={mepElement.Document?.Title}");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (OptimizationFlags.UseDiagnosticMode)
                        SafeLog("Refresh_debug.log", $"[ClashZoneService] ❌ ERROR reading Size parameter for element {mepElement?.Id.GetIntegerValue() ?? -1}: {ex.Message}, Doc={mepElement?.Document?.Title}");
                    DebugLogger.Warning($"[ClashZoneService] Error reading Size parameter value: {ex.Message}");
                }
            }
            return string.Empty; // Return empty if Size parameter not available
        }
        
        /// <summary>
        /// Get MEP element size as string - OPTIMIZED FOR ALL CATEGORIES: Reads Size parameter directly for exact schedule display
        /// Works for: Pipes, Ducts, Cable Trays, Duct Accessories, Pipe Accessories, etc.
        /// </summary>
        /// <param name="mepElement">The MEP element (already in memory during refresh - zero cost, no linked file access)</param>
        /// <param name="width">Width in internal units (feet) - for pipes, this is outer diameter; for ducts/trays, this is width</param>
        /// <param name="height">Height in internal units (feet) - for pipes, same as width (diameter); for ducts/trays, this is height</param>
        /// <param name="shape">Shape of the element (Round/Circular or Rectangular)</param>
        /// <param name="nominalDiameter">Nominal diameter for pipes only (in internal units) - used as fallback for pipes</param>
        /// <returns>Size string matching schedule display (e.g., "20 mmø" for pipes, "600x300" for ducts, "400x200" for cable trays)</returns>
        private string GetMepElementSizeString(Element mepElement, double width, double height, string shape, double nominalDiameter = 0, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ PERF OPT: Check paramCache first (captured in bulk)
                if (paramCache != null && paramCache.TryGetValue("Size", out var cachedSize) && !string.IsNullOrEmpty(cachedSize))
                {
                    return cachedSize;
                }
                // ✅ OPTIMIZED APPROACH FOR ALL CATEGORIES: Read Size parameter as string directly from element
                // This gives us the exact value shown in schedules (e.g., "20 mmø" for pipes, "600x300" for ducts) without calculations
                // Element is already in memory during refresh, so this is a zero-cost operation (no linked file access)
                // Works for: Pipes, Ducts, Cable Trays, Duct Accessories, Pipe Accessories, Conduits, etc.
                var sizeParam = mepElement?.LookupParameter("Size");
                if (sizeParam != null)
                {
                    // Try AsString() first (for string parameters)
                    if (sizeParam.StorageType == StorageType.String)
                    {
                        var sizeString = sizeParam.AsString();
                        if (!string.IsNullOrWhiteSpace(sizeString))
                        {
                            // Return the actual Size parameter value (matches schedule display exactly)
                            return sizeString.Trim();
                        }
                    }
                    // Try AsValueString() as fallback (for numeric parameters that display as formatted strings)
                    else
                    {
                        var sizeString = sizeParam.AsValueString();
                        if (!string.IsNullOrWhiteSpace(sizeString))
                        {
                            // Return the formatted value string (e.g., "20 mmø" from a numeric parameter)
                            return sizeString.Trim();
                        }
                    }
                }
                
                // ✅ FALLBACK: If Size parameter not available or empty, calculate formatted size
                // This maintains backward compatibility for cases where Size parameter doesn't exist
                // Uses numeric values (width, height) stored in ClashZone for calculation
                return FormatMepElementSize(mepElement, width, height, shape, nominalDiameter);
            }
            catch
            {
                // Fallback to calculated format on error
                return FormatMepElementSize(mepElement, width, height, shape, nominalDiameter);
            }
        }
        
        /// <summary>
        /// Format MEP element size as string (e.g., "600x300", "Ø200") - Used as fallback when Size parameter is not available
        /// Works for ALL categories: Pipes, Ducts, Cable Trays, Duct Accessories, etc.
        /// </summary>
        /// <param name="mepElement">The MEP element</param>
        /// <param name="width">Width in internal units (feet) - for pipes, this is outer diameter; for ducts/trays, this is width</param>
        /// <param name="height">Height in internal units (feet) - for pipes, same as width (diameter); for ducts/trays, this is height</param>
        /// <param name="shape">Shape of the element (Round/Circular or Rectangular)</param>
        /// <param name="nominalDiameter">Optional nominal diameter for pipes only (in internal units) - used for Size parameter display instead of outer diameter</param>
        /// <returns>Formatted size string (e.g., "Ø200" for round pipes, "600x300" for rectangular ducts/trays)</returns>
        private string FormatMepElementSize(Element mepElement, double width, double height, string shape, double nominalDiameter = 0)
        {
            try
            {
                if (shape == "Round" || shape == "Circular")
                {
                    // ✅ FALLBACK FOR ROUND ELEMENTS (Pipes, Round Ducts): Use nominal diameter for pipes, width for others
                    // For pipes: Outer diameter is used for sleeve sizing, but Size parameter should show nominal diameter (e.g., "Ø20" not "Ø25")
                    // For round ducts: Use width (which is the diameter)
                    double diameterToUse = (nominalDiameter > 0) ? nominalDiameter : width;
                    double diameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(diameterToUse);
                    return $"Ø{Math.Round(diameterMm, 0)}";
                }
                else
                {
                    // ✅ FALLBACK FOR RECTANGULAR ELEMENTS (Ducts, Cable Trays, Accessories): Format as "WxH"
                    // Convert from feet to mm and format as "600x300"
                    double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                    double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
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
        private string GetMepSystemAbbreviation(Element mepElement, System.Collections.Generic.Dictionary<string, string>? paramCache = null)
        {
            try
            {
                // ✅ PERF OPT: Check paramCache first (captured in bulk)
                if (paramCache != null && paramCache.TryGetValue("System Abbreviation", out var cachedAbbr) && !string.IsNullOrEmpty(cachedAbbr))
                    return cachedAbbr;

                // Try direct BIP first (O(1) vs O(N) LookupParameter)
                var abbrevParam = mepElement.LookupParameter("System Abbreviation");
                if (abbrevParam != null && abbrevParam.StorageType == StorageType.String)
                    return abbrevParam.AsString() ?? string.Empty;

                // Fallback: System Name → abbreviate
                var systemNameParam = mepElement.LookupParameter("System Name");
                if (systemNameParam != null && systemNameParam.StorageType == StorageType.String)
                {
                    var systemName = systemNameParam.AsString();
                    if (!string.IsNullOrEmpty(systemName))
                        return systemName.Length > 3 ? systemName.Substring(0, 3).ToUpper() : systemName.ToUpper();
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// ✅ OOP METHOD: Get damper connector info using DamperConnectorService
        /// Checks if MEP connector is found and determines which side it's on
        /// Strategy will handle clearance from UI accordingly based on damper type
        /// </summary>
        /// <param name="wallOrientation">Wall orientation ("X" or "Y") to prioritize wall width axis during detection</param>
        private DamperConnectorInfo GetDamperConnectorInfo(Element mepElement, string mepCategory, string wallOrientation = null, Dictionary<string, string>? mepParamDict = null)
        {
            try
            {
                // ✅ PERFORMANCE OPTIMIZATION: Check parameter cache first
                // If we captured synthetic parameters in ParameterSnapshotService, use them!
                if (mepParamDict != null && mepParamDict.TryGetValue("HasMepConnector", out var hasConnStr))
                {
                    bool hasConn = hasConnStr.Equals("True", StringComparison.OrdinalIgnoreCase);
                    mepParamDict.TryGetValue("DamperConnectorSide", out var connSide);
                    mepParamDict.TryGetValue("IsStandardDamper", out var isStdStr);
                    mepParamDict.TryGetValue("DamperType", out var damperType);
                    
                    bool isStd = isStdStr?.Equals("True", StringComparison.OrdinalIgnoreCase) ?? true;
                    
                    // Note: If wallOrientation is provided but differ from the one during capture, 
                    // we might need to re-detect, but usually it's fine for the first pass or for standard dampers.
                    // For non-standard dampers where wall orientation MATTERS for world side detection:
                    // If wallOrientation is null, we can definitely use the cache.
                    // If it's provided, we might still use cache but it's a trade-off.
                    // THE USERS REQUESTED METICULOUSNESS.
                    
                    return new DamperConnectorInfo(
                        hasMepConnector: hasConn,
                        connectorSide: connSide ?? string.Empty,
                        isStandardDamper: isStd,
                        damperType: damperType ?? "Standard",
                        requiresMepSideClearance: !isStd
                    );
                }

                // Fallback: Direct API call via centralized service
                var damperConnectorService = new DamperConnectorService();
                var connectorInfo = damperConnectorService.DetectConnectorInfo(mepElement, mepCategory, wallOrientation);
                
                if (connectorInfo.HasMepConnector)
                {
                    _log($"[DEBUG] Damper {mepElement?.Id}: Found MEP connector on side '{connectorInfo.ConnectorSide}', Type='{connectorInfo.DamperType}', IsStandard={connectorInfo.IsStandardDamper}");
                }
                
                return connectorInfo;
            }
            catch (Exception ex)
            {
                _log($"[DEBUG] Error getting damper connector info: {ex.Message}");
                return DamperConnectorInfo.None;
            }
        }
        
        // ✅ OOP REFACTORING: Removed GetWallDirection(), GetWallDirectionType(), GetStructuralElementNormal()
        // All three methods replaced by WallDirectionService (eliminates code duplication):
        // - WallDirectionService.GetWallDirection()
        // - WallDirectionService.GetWallDirectionType()
        // - WallDirectionService.GetStructuralElementNormal()
        
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
                         famInst.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
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
        private (double width, double height) GetMepElementDimensions(Element mepElement, Dictionary<string, string>? paramCache = null)
        {
            try
            {
                if (mepElement is Duct duct)
                {
                    double width = 0;
                    double height = 0;
                    double diameter = 0;

                    // Try paramCache first if available
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramCache != null)
                    {
                        if (paramCache.TryGetValue("Width", out var wStr) && double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var wVal)) width = wVal;
                        if (paramCache.TryGetValue("Height", out var hStr) && double.TryParse(hStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var hVal)) height = hVal;
                        if (paramCache.TryGetValue("Diameter", out var dStr) && double.TryParse(dStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var dVal)) diameter = dVal;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog instead of direct SafeFileLogger
                        if (width <= 0 && height <= 0 && diameter <= 0)
                        {
                            var keys = string.Join(", ", paramCache.Keys.Take(10));
                            SafeLog("pipe_dimension_debug.log", $"DUCT {duct.Id}: paramCache FAILED, W='{wStr ?? "null"}', H='{hStr ?? "null"}', D='{dStr ?? "null"}', Keys=[{keys}] - Falling back to API");
                        }
                    }
                    
                    // ✅ FIX: Fall back to direct Revit API if paramCache didn't provide values
                    if (width <= 0 && height <= 0 && diameter <= 0)
                    {
                        width = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                        height = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                        diameter = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        if (width > 0 || height > 0 || diameter > 0)
                        {
                            SafeLog("pipe_dimension_debug.log", $"DUCT {duct.Id}: Direct API fallback OK, W={width * 304.8:F1}mm, H={height * 304.8:F1}mm, D={diameter * 304.8:F1}mm");
                        }
                    }
                    
                    if ((width <= 0.0 || height <= 0.0) && diameter > 0.0)
                    {
                        width = diameter;
                        height = diameter;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        SafeLog("pipe_dimension_debug.log", $"DUCT {duct.Id}: ROUND DUCT - Using diameter={diameter:F6}ft ({diameter * 304.8:F1}mm) for W/H");
                    }
                    
                    // ✅ PERFORMANCE OPT: Use SafeLog
                    if (width <= 0 && height <= 0)
                    {
                        SafeLog("pipe_dimension_debug.log", $"DUCT {duct.Id}: ⚠️ ZERO DIMENSIONS! W={width:F6}, H={height:F6}, D={diameter:F6}, UsedCache={paramCache != null}");
                    }
                    
                    return (width, height);
                }
                else if (mepElement is Pipe pipe)
                {
                    double outerDiameter = 0;
                    double nominalDiameter = 0;

                    // Try paramCache first if available
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramCache != null)
                    {
                        if (paramCache.TryGetValue("Outside Diameter", out var odStr) && double.TryParse(odStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var odVal)) outerDiameter = odVal;
                        // Try "Nominal Diameter" first, then "Diameter" as fallback (Revit uses both names)
                        if (paramCache.TryGetValue("Nominal Diameter", out var ndStr) && double.TryParse(ndStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var ndVal)) nominalDiameter = ndVal;
                        else if (paramCache.TryGetValue("Diameter", out var dStr) && double.TryParse(dStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var dVal)) nominalDiameter = dVal;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        if (outerDiameter <= 0 && nominalDiameter <= 0)
                        {
                            var keys = string.Join(", ", paramCache.Keys.Take(10));
                            SafeLog("pipe_dimension_debug.log", $"PIPE {pipe.Id}: paramCache FAILED, OD='{odStr ?? "null"}', ND='{ndStr ?? "null"}', Keys=[{keys}] - Falling back to API");
                        }
                    }
                    
                    // ✅ FIX: Fall back to direct Revit API if paramCache didn't provide values
                    if (outerDiameter <= 0 && nominalDiameter <= 0)
                    {
                        outerDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0;
                        nominalDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)?.AsDouble() ?? 0;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        SafeLog("pipe_dimension_debug.log", $"PIPE {pipe.Id}: Direct API used, OD={outerDiameter:F6}ft ({outerDiameter * 304.8:F1}mm), ND={nominalDiameter:F6}ft, Doc={pipe.Document?.Title}");
                    }
                    
                    var diameter = outerDiameter;
                    if (diameter <= 0)
                    {
                        diameter = nominalDiameter;
                    }
                    
                    // ✅ PERFORMANCE OPT: Use SafeLog
                    if (diameter <= 0)
                    {
                        SafeLog("pipe_dimension_debug.log", $"PIPE {pipe.Id}: ⚠️ ZERO DIAMETER! OD={outerDiameter:F6}, ND={nominalDiameter:F6}, UsedCache={paramCache != null}");
                    }
                    
                    return (diameter, diameter);
                }
                else if (mepElement is Autodesk.Revit.DB.Electrical.CableTray cableTray)
                {
                    double width = 0;
                    double height = 0;

                    // Try paramCache first if available
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramCache != null)
                    {
                        if (paramCache.TryGetValue("Width", out var wStr) && double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var wVal)) width = wVal;
                        if (paramCache.TryGetValue("Height", out var hStr) && double.TryParse(hStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var hVal)) height = hVal;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        if (width <= 0 && height <= 0)
                        {
                            var keys = string.Join(", ", paramCache.Keys.Take(10));
                            SafeLog("pipe_dimension_debug.log", $"CABLETRAY {cableTray.Id}: paramCache FAILED, W='{wStr ?? "null"}', H='{hStr ?? "null"}', Keys=[{keys}] - Falling back to API");
                        }
                    }
                    
                    // ✅ FIX: Fall back to direct Revit API if paramCache didn't provide values
                    if (width <= 0 && height <= 0)
                    {
                        width = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)?.AsDouble() ?? 0;
                        height = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? 0;
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        if (width > 0 || height > 0)
                        {
                            SafeLog("pipe_dimension_debug.log", $"CABLETRAY {cableTray.Id}: Direct API fallback OK, W={width * 304.8:F1}mm, H={height * 304.8:F1}mm");
                        }
                    }
                    
                    return (width, height);
                }
                else if (mepElement is Conduit conduit)
                {
                    double width = 0;
                    double height = 0;

                    // Try paramCache first if available
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramCache != null)
                    {
                        if (paramCache.TryGetValue("Width", out var wStr) && double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var wVal)) width = wVal;
                        if (paramCache.TryGetValue("Height", out var hStr) && double.TryParse(hStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var hVal)) height = hVal;
                    }
                    
                    // ✅ FIX: Fall back to direct Revit API if paramCache didn't provide values
                    if (width <= 0 && height <= 0)
                    {
                        width = conduit.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)?.AsDouble() ?? 0;
                        height = conduit.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)?.AsDouble() ?? 0;
                    }
                    
                    return (width, height);
                }
                else if (mepElement is FamilyInstance famInst)
                {
                    double width = 0;
                    double height = 0;

                    // Try paramCache first if available
                    // NOTE: ParameterSnapshotService now returns raw double value as string (already in internal units/feet)
                    if (paramCache != null)
                    {
                        if (paramCache.TryGetValue("Damper Width", out var dwStr) && double.TryParse(dwStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var dwVal)) width = dwVal;
                        else if (paramCache.TryGetValue("Width", out var wStr) && double.TryParse(wStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var wVal)) width = wVal;
                        
                        if (paramCache.TryGetValue("Damper Height", out var dhStr) && double.TryParse(dhStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var dhVal)) height = dhVal;
                        else if (paramCache.TryGetValue("Height", out var hStr) && double.TryParse(hStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var hVal)) height = hVal;
                    }
                    
                    // ✅ FIX: Fall back to direct Revit API if paramCache didn't provide values
                    if (width <= 0 && height <= 0)
                    {
                        var widthParam = famInst.LookupParameter("Damper Width") ?? famInst.LookupParameter("Width") ?? famInst.LookupParameter("width");
                        var heightParam = famInst.LookupParameter("Damper Height") ?? famInst.LookupParameter("Height") ?? famInst.LookupParameter("height");
                        width = widthParam?.AsDouble() ?? 0.1;
                        height = heightParam?.AsDouble() ?? 0.1;
                    }
                    
                    return (width, height);
                }
                
                return (0.1, 0.1);
            }
            catch
            {
                return (0.1, 0.1);
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
                        if (BoundingBoxService.BoundingBoxesIntersect(damperBbox.Min, damperBbox.Max, wallBbox.Min, wallBbox.Max))
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
        // ✅ OOP REFACTORING: Removed duplicate BoundingBoxesIntersect - now uses BoundingBoxService.BoundingBoxesIntersect()

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
                            // Try to get the damper element from document (works for active doc and linked files)
                            var damperElement = document.GetElement(clashZone.MepElementId);
                            
                            BoundingBoxXYZ damperBbox = null;
                            if (damperElement != null)
                            {
                                // Element is in active document or same linked file
                                damperBbox = damperElement.get_BoundingBox(null);
                            }
                            else
                            {
                                // Element might be in a different linked file - try to get from intersection point
                                // Use stored intersection point coordinates to create bounding box approximation
                                if (clashZone.IntersectionPointX != 0 || clashZone.IntersectionPointY != 0 || clashZone.IntersectionPointZ != 0)
                                {
                                    var intersectionPoint = new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                                    // Create a small bounding box around intersection point (approx 200mm = 0.66ft cube)
                                    const double approximateSize = 0.66; // 200mm in feet
                                    damperBbox = new BoundingBoxXYZ
                                    {
                                        Min = new XYZ(intersectionPoint.X - approximateSize, intersectionPoint.Y - approximateSize, intersectionPoint.Z - approximateSize),
                                        Max = new XYZ(intersectionPoint.X + approximateSize, intersectionPoint.Y + approximateSize, intersectionPoint.Z + approximateSize)
                                    };
                                    _log($"[PRE-CALC-XML] Damper {clashZone.MepElementId} in linked file - using intersection point approximation");
                                }
                            }
                            
                            if (damperBbox != null)
                            {
                                damperLocations.Add((damperId: clashZone.MepElementId, bbox: damperBbox, wallId: clashZone.StructuralElementId));
                                _log($"[PRE-CALC-XML] Damper {clashZone.MepElementId} location cached from XML (wall: {clashZone.StructuralElementId})");
                            }
                            else
                            {
                                _log($"[PRE-CALC-XML] Warning: Could not get bounding box for damper {clashZone.MepElementId} from XML");
                            }
                        }
                    }
                }
                
                // STEP 2: Add dampers from current intersections (from current refresh cycle)
                foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in currentIntersections)
                {
                    var mepCat = GetElementCategoryName(mepElement);
                    
                    // ✅ SIMPLE: Add all Duct Accessories - category check is sufficient for avoidance
                    if (IsDamperElement(mepElement))
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
        /// ✅ SIMPLE: Category check is sufficient - all Duct Accessories are treated as dampers for avoidance
        /// </summary>
        private bool IsDamperClashZone(ClashZone clashZone)
        {
            try
            {
                // ✅ SIMPLE: If it's a Duct Accessory category, treat it as a damper
                return string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Check if a duct is near a damper using pre-calculated locations (Efficient O(n) lookup)
        /// ✅ CRITICAL FIX: Now checks SAME WALL requirement before proximity check
        /// </summary>
        private bool IsDuctNearDamperOnSameWall(Element ductElement, ElementId wallId, XYZ intersectionPoint, List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations)
        {
            try
            {
                // ✅ CRITICAL: Check intersection point proximity FIRST (most reliable - checks damper at duct end where intersection occurs)
                // Use tight tolerance for intersection point check (10mm) - per user requirement for duct/damper proximity detection
                const double intersectionTolerance = 0.0328; // 10mm tolerance for damper at intersection point (0.0328ft)
                
                // ✅ FALLBACK: Same tolerance for bounding box check for connected duct-damper pairs
                const double bboxProximityTolerance = 0.0328; // 10mm tolerance for connected duct-damper pairs (0.0328ft)
                
                // Get duct bounding box
                var ductBbox = ductElement.get_BoundingBox(null);
                if (ductBbox == null)
                {
                    _log($"[DUCT-DAMPER] Duct {ductElement.Id} has no bounding box - cannot check damper proximity");
                    return false;
                }
                
                // ✅ CRITICAL FIX: Filter dampers by SAME WALL first (major optimization and correctness fix)
                var dampersOnSameWall = damperLocations.Where(d => d.wallId == wallId).ToList();
                _log($"[DUCT-DAMPER] Checking duct {ductElement.Id} on wall {wallId} against {dampersOnSameWall.Count} dampers on same wall (total dampers: {damperLocations.Count})");
                
                if (dampersOnSameWall.Count == 0)
                {
                    _log($"[DUCT-DAMPER] No dampers on wall {wallId} - duct {ductElement.Id} will proceed");
                    return false;
                }
                
                // Check pre-calculated damper locations on SAME WALL (Efficient O(n) lookup)
                foreach (var (damperId, damperBbox, _) in dampersOnSameWall)
                {
                    // Skip if it's the same element
                    if (damperId == ductElement.Id) continue;
                    
                    // ✅ METHOD 1: Check if damper center/bbox is near intersection point (most reliable for duct-damper combos)
                    XYZ damperCenter = (damperBbox.Min + damperBbox.Max) * 0.5;
                    double distanceToIntersection = damperCenter.DistanceTo(intersectionPoint);
                    
                    if (distanceToIntersection <= intersectionTolerance)
                    {
                        _log($"[DUCT-DAMPER] ✓ MATCH AT INTERSECTION: Duct {ductElement.Id} intersection point is {distanceToIntersection:F4}ft from Damper {damperId} center on wall {wallId} (tolerance: {intersectionTolerance}ft = 10mm)");
                        return true;
                    }
                    
                    // ✅ METHOD 2: Check if damper bbox contains or is near intersection point
                    if (IsPointNearBoundingBox(intersectionPoint, damperBbox, intersectionTolerance))
                    {
                        _log($"[DUCT-DAMPER] ✓ MATCH AT INTERSECTION (bbox): Duct {ductElement.Id} intersection point is within {intersectionTolerance}ft (10mm) of Damper {damperId} bounding box on wall {wallId}");
                        return true;
                    }
                    
                    // ✅ METHOD 3: Fallback - Check if duct and damper bounding boxes are within proximity (for connected pairs)
                    var bboxDistance = GetMinimumDistanceBetweenBoundingBoxes(ductBbox, damperBbox);
                    if (bboxDistance <= bboxProximityTolerance)
                    {
                        _log($"[DUCT-DAMPER] ✓ MATCH (bbox proximity): Duct {ductElement.Id} bbox is {bboxDistance:F4}ft from Damper {damperId} bbox on wall {wallId} (tolerance: {bboxProximityTolerance}ft = 10mm)");
                        return true;
                    }
                }
                
                _log($"[DUCT-DAMPER] No dampers found near duct {ductElement.Id} on wall {wallId} (checked intersection point and bbox proximity)");
                return false;
            }
            catch (Exception ex)
            {
                _log($"Error checking duct-damper proximity: {ex.Message}");
                _log($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }
        
        /// <summary>
        /// Check if a point is near a bounding box (within tolerance)
        /// </summary>
        private bool IsPointNearBoundingBox(XYZ point, BoundingBoxXYZ bbox, double tolerance)
        {
            try
            {
                // Expand bbox by tolerance
                var expandedMin = new XYZ(bbox.Min.X - tolerance, bbox.Min.Y - tolerance, bbox.Min.Z - tolerance);
                var expandedMax = new XYZ(bbox.Max.X + tolerance, bbox.Max.Y + tolerance, bbox.Max.Z + tolerance);
                
                // Check if point is within expanded bbox
                return point.X >= expandedMin.X && point.X <= expandedMax.X &&
                       point.Y >= expandedMin.Y && point.Y <= expandedMax.Y &&
                       point.Z >= expandedMin.Z && point.Z <= expandedMax.Z;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// [DEPRECATED] Old method - kept for backward compatibility but should not be used
        /// Use IsDuctNearDamperOnSameWall instead
        /// </summary>
        private bool IsDuctNearDamper(Element ductElement, List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations)
        {
            // This method doesn't check same wall - use IsDuctNearDamperOnSameWall instead
            _log($"[WARNING] IsDuctNearDamper called without wall check - this is deprecated");
            return false;
        }
        
        /// <summary>
        /// Check if an element is a damper based on family name or category
        /// </summary>
        private bool IsDamperElement(Element element)
        {
            try
            {
                // ✅ SIMPLE: If it's a Duct Accessory category, treat it as a damper for avoidance logic
                // No need to check family name - category is sufficient
                return element.Category?.Id.GetIntegerValue() == (int)BuiltInCategory.OST_DuctAccessory;
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
        /// Get wall orientation from wall direction type for clustering distance calculation
        /// </summary>
        private string GetWallOrientationFromType(string wallDirectionType)
        {
            try
            {
                if (string.IsNullOrEmpty(wallDirectionType))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[GetWallOrientationFromType] WallDirectionType is null or empty, defaulting to X");
                    return "X"; // Default to X orientation
                }
                
                // Convert wall direction type to orientation for clustering distance calculation
                string orientation;
                if (wallDirectionType.Contains("X-WALL"))
                {
                    orientation = "X"; // X-oriented walls use X,Z coordinates for clustering
                }
                else if (wallDirectionType.Contains("Y-WALL"))
                {
                    orientation = "Y"; // Y-oriented walls use Y,Z coordinates for clustering
                }
                else if (wallDirectionType.Contains("FRAMING"))
                {
                    orientation = "Z"; // Framing uses X,Y coordinates for clustering
                }
                else
                {
                    orientation = "X"; // Default fallback
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[GetWallOrientationFromType] Unknown wall direction type '{wallDirectionType}', defaulting to X");
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetWallOrientationFromType] WallDirectionType='{wallDirectionType}' → Orientation='{orientation}'");
                return orientation;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GetWallOrientationFromType] Error converting wall direction type '{wallDirectionType}': {ex.Message}");
                return "X"; // Default to X orientation
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        if (!DeploymentConfiguration.DeploymentMode)
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
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("refresh.log", $"[GetMepElementOrientationFromBbox] Element {mepElement.Id}: BboxWidth={bboxWidth:F3}, BboxHeight={bboxHeight:F3}, Orientation={mepOrientation}");
                }
                
                return mepOrientation;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("refresh.log", $"[GetMepElementOrientationFromBbox] Error getting orientation for element {mepElement?.Id}: {ex.Message}");
                }
                return "X"; // Default to X orientation
            }
        }

        /// <summary>
        /// Get MEP element orientation vector.
        /// ✅ OPTIMIZED: Removed duplicate double-guarded logging. Only logs in DiagnosticMode.
        /// ✅ OPTIMIZED: Pipes and Conduits return immediately (round elements, no complex orientation needed).
        /// </summary>
        public XYZ GetMepElementOrientation(Element mepElement)
        {
            try
            {
                // ✅ FAST PATH: Pipes and Conduits are round — just return centerline direction, no helper needed.
                if (mepElement is Pipe pipe && pipe.Location is LocationCurve pipeCurveF)
                {
                    var l = pipeCurveF.Curve as Line;
                    return l?.Direction ?? XYZ.BasisX;
                }
                if (mepElement is Conduit conduit && conduit.Location is LocationCurve conduitCurveF)
                {
                    var l = conduitCurveF.Curve as Line;
                    return l?.Direction ?? XYZ.BasisX;
                }

                bool diagMode = OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode;

                if (mepElement is Duct duct && duct.Location is LocationCurve curve)
                {
                    var line = curve.Curve as Line;
                    if (line != null)
                    {
                        var direction = line.Direction;
                        try
                        {
                            var (orientation, widthDirection) = Helpers.MepElementOrientationHelper.GetDuctWidthOrientation(duct);
                            if (diagMode)
                                DebugLogger.Info($"[GetMepElementOrientation] Duct {mepElement.Id}: Width orientation={orientation}, WidthDirection=({widthDirection.X:F3}, {widthDirection.Y:F3}, {widthDirection.Z:F3})");
                            return widthDirection;
                        }
                        catch
                        {
                            return direction;
                        }
                    }
                }
                else if (mepElement is Duct verticalDuct && verticalDuct.Location is LocationPoint)
                {
                    // For vertical ducts through floors, use helper to determine X or Y orientation
                    try
                    {
                        var (_, widthDirection) = Helpers.MepElementOrientationHelper.GetDuctWidthOrientation(verticalDuct);
                        return widthDirection;
                    }
                    catch
                    {
                        return XYZ.BasisY;
                    }
                }
                else if (mepElement is FamilyInstance ductAccessory &&
                         ductAccessory.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_DuctAccessory &&
                         ductAccessory.Location is LocationCurve ductAccessoryCurve)
                {
                    var line = ductAccessoryCurve.Curve as Line;
                    if (line != null) return line.Direction;
                }
                else if (mepElement is Autodesk.Revit.DB.Electrical.CableTray cableTray && cableTray.Location is LocationCurve cableTrayCurve)
                {
                    var line = cableTrayCurve.Curve as Line;
                    if (line != null)
                    {
                        var direction = line.Direction;
                        double absZ = Math.Abs(direction.Z);
                        bool isVertical = absZ > Math.Max(Math.Abs(direction.X), Math.Abs(direction.Y)) * 0.7;
                        if (isVertical)
                        {
                            try
                            {
                                var (_, widthDirection) = Helpers.MepElementOrientationHelper.GetCableTrayWidthOrientation(cableTray);
                                return widthDirection;
                            }
                            catch { return direction; }
                        }
                        return direction;
                    }
                }
                else if (mepElement.Category?.Id.GetIntegerValue() == (int)BuiltInCategory.OST_DuctAccessory)
                {
                    var damper = mepElement as FamilyInstance;
                    if (damper != null)
                        return damper.GetTotalTransform().BasisX;
                }

                return XYZ.BasisX;
            }
            catch
            {
                return XYZ.BasisX;
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
        /// Uses the same logic as HostLevelHelper.GetHostReferenceLevel to get immediate reference level
        /// </summary>
        private (string levelName, double levelElevation) GetMepElementLevelInfo(Element mepElement, Dictionary<string, string>? paramCache = null)
        {
            // ✅ PERF: Only log in explicit diagnostic mode, never per-element in normal runs
            bool diagMode = OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode;
            try
            {
                // ✅ OPTIMIZATION: Check parameter cache first (captured in bulk)
                if (paramCache != null)
                {
                    string levelName = null;
                    if (paramCache.TryGetValue("Reference Level", out var rl) && !string.IsNullOrEmpty(rl)) levelName = rl;
                    else if (paramCache.TryGetValue("Level", out var l) && !string.IsNullOrEmpty(l)) levelName = l;
                    else if (paramCache.TryGetValue("Schedule Level", out var sl) && !string.IsNullOrEmpty(sl)) levelName = sl;

                    if (!string.IsNullOrEmpty(levelName))
                    {
                        double elevation = 0;
                        if (paramCache.TryGetValue("Reference Level Elevation", out var elevStr) && double.TryParse(elevStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var elevVal))
                            elevation = elevVal;
                        if (diagMode) SafeFileLogger.SafeAppendText("refresh.log", $"[ClashZoneService] MEP Element {mepElement.Id}: Using cached level '{levelName}' (elevation: {elevation})");
                        return (levelName, elevation);
                    }
                }

                // Fallback: HostLevelHelper
                var refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(mepElement.Document, mepElement);
                if (refLevel != null)
                {
                    if (diagMode) SafeFileLogger.SafeAppendText("refresh.log", $"[ClashZoneService] MEP Element {mepElement.Id}: Using reference level '{refLevel.Name}' (elevation: {refLevel.Elevation})");
                    return (refLevel.Name, refLevel.Elevation);
                }

                // Fallback: LevelId
                if (mepElement.LevelId != ElementId.InvalidElementId)
                {
                    var level = mepElement.Document.GetElement(mepElement.LevelId) as Level;
                    if (level != null)
                    {
                        if (diagMode) SafeFileLogger.SafeAppendText("refresh.log", $"[ClashZoneService] MEP Element {mepElement.Id}: Using LevelId level '{level.Name}' (elevation: {level.Elevation})");
                        return (level.Name, level.Elevation);
                    }
                }

                // Fallback: location point Z
                if (mepElement.Location is LocationPoint locationPoint)
                {
                    var elevation = locationPoint.Point.Z;
                    if (diagMode) SafeFileLogger.SafeAppendText("refresh.log", $"[ClashZoneService] MEP Element {mepElement.Id}: Using location point elevation {elevation}");
                    return ($"Auto-Level-{elevation:F2}", elevation);
                }
                else if (mepElement.Location is LocationCurve locationCurve)
                {
                    var startPoint = locationCurve.Curve.GetEndPoint(0);
                    var elevation = startPoint.Z;
                    if (diagMode) SafeFileLogger.SafeAppendText("refresh.log", $"[ClashZoneService] MEP Element {mepElement.Id}: Using location curve elevation {elevation}");
                    return ($"Auto-Level-{elevation:F2}", elevation);
                }
                
                // Final fallback
                if (diagMode) DebugLogger.Warning($"[ClashZoneService] MEP Element {mepElement.Id}: No level found (fallback 'Level 1')");
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
        
        /// <summary>
        /// Check if there's an existing cluster sleeve at the given placement point
        /// </summary>
        private bool CheckForExistingClusterSleeve(XYZ placementPoint, Document document, double tolerance = 0.05) // 50mm tolerance
        {
            try
            {
                // Look for existing cluster sleeves near the placement point
                // Cluster sleeves are typically rectangular opening families
                
                var openingFamilies = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                    .ToList();
                
                _log($"[CheckForExistingClusterSleeve] Checking for cluster sleeves at placement point {placementPoint}");
                
                // Check if any opening exists within tolerance of the placement point
                foreach (var opening in openingFamilies)
                {
                    if (opening.Location is LocationPoint locationPoint)
                    {
                        double distance = locationPoint.Point.DistanceTo(placementPoint);
                        if (distance <= tolerance)
                        {
                            // Check if this is likely a cluster sleeve (rectangular, larger size)
                            var familyName = opening.Symbol?.Family?.Name ?? "";
                            var isRectangular = familyName.Contains("Rectangular") || familyName.Contains("Rect");
                            
                            // Cluster sleeves are typically larger than individual sleeves
                            var widthParam = opening.LookupParameter("Width");
                            var heightParam = opening.LookupParameter("Height");
                            var diameterParam = opening.LookupParameter("Diameter");
                            bool isLargeSleeve = IsParameterLargeValue(widthParam, 300) ||
                                                 IsParameterLargeValue(heightParam, 300) ||
                                                 IsParameterLargeValue(diameterParam, 300);
                            
                            if (isRectangular || isLargeSleeve)
                            {
                                _log($"[CheckForExistingClusterSleeve] Found existing cluster sleeve at distance {distance:F3}ft: {opening.Id} ({familyName})");
                                return true;
                            }
                        }
                    }
                }
                
                _log($"[CheckForExistingClusterSleeve] No existing cluster sleeve found at placement point {placementPoint}");
                return false;
            }
            catch (Exception ex)
            {
                _log($"[CheckForExistingClusterSleeve] Error checking for existing cluster sleeve: {ex.Message}");
                return false;
            }
        }
        
        private void UpdateExistingClashZone(ClashZone existingZone, Element mepElement, Element structuralElement, XYZ intersectionPoint, BoundingBoxXYZ boundingBox, Document document, HashSet<string>? parameterWhitelist = null, HashSet<int>? existingSleeveIds = null, HashSet<string>? existingOpeningPointKeys = null, Dictionary<int, Dictionary<string, string>>? mepParamsCache = null, Dictionary<int, Dictionary<string, string>>? hostParamsCache = null, Dictionary<string, MepElementComputedData>? mepDataCache = null)
        {
            using (var metadataOp = _performanceMonitor?.TrackOperation("6d4a. Metadata Updates") as IDisposable)
            {
                Dictionary<string, string>? mepParamDict = null;
                if (mepParamsCache != null) mepParamsCache.TryGetValue(mepElement.Id.GetIntegerValue(), out mepParamDict);

                // PERF OPT: Use pre-computed MEP element data (same element can update multiple zones)
                MepElementComputedData mepData = null;
                mepDataCache?.TryGetValue(GetElementCacheKey(mepElement), out mepData);

                var mepCategory = (mepData != null) ? mepData.Category : GetElementCategoryName(mepElement, mepParamDict);
                existingZone.IntersectionPoint = intersectionPoint;
                existingZone.SleevePlacementPoint = intersectionPoint;
                existingZone.ClashBoundingBox = boundingBox;
                existingZone.MepElementSize = GetMepElementSize(mepElement);
                existingZone.RequiredClearance = CalculateRequiredClearance(existingZone.MepElementSize);
                existingZone.MepElementGeometryHash = CalculateElementGeometryHash(mepElement);
                existingZone.StructuralElementGeometryHash = CalculateElementGeometryHash(structuralElement);

                // PERF OPT: Use cached values — eliminates 4 COM call sequences per zone
                existingZone.MepSystemType = (mepData != null) ? mepData.SystemType : GetMepSystemType(mepElement, mepCategory, mepParamDict);
                existingZone.MepServiceType = (mepData != null) ? mepData.ServiceType : GetMepServiceType(mepElement, mepParamDict);
//[lines 5928-5933 unchanged]
                existingZone.ElevationFromLevel = (mepData != null) ? mepData.ElevationFromLevel : GetMepElementOffset(mepElement, mepParamDict);
            }
            
            using (var paramOp = _performanceMonitor?.TrackOperation("6d4b. Parameter Snapshot") as IDisposable)
            {
                // ✅ STAGE 1 (REFRESH): Capture MEP and Host parameters using ParameterSnapshotService
                // This ensures existing ClashZones also get parameter snapshots updated during refresh
                // Same logic as CreateClashZone - capture all whitelisted parameters
                // ✅ PHASE 4 OPTIMIZATION: Skip synchronous capture if lazy capture is enabled (saves ~9 seconds for 888 zones)
                try
                {
                    if (_clashZoneStorage != null && !OptimizationFlags.UseLazyParameterCapture)
                    {
                        var parameterSnapshotService = new ParameterSnapshotService();
                        
                        // Build whitelist from storage (includes common keys, user-defined keys, and learned keys)
                        // ✅ PHASE 3 OPTIMIZATION: Reuse pre-cached whitelist
                        var whitelist = parameterWhitelist ?? parameterSnapshotService.BuildWhitelist(_clashZoneStorage, new[] { (mepElement, structuralElement) });
                        
                        // Capture MEP element parameters
                        var mepParameterValues = parameterSnapshotService.CaptureParams(mepElement, whitelist);
                        
                        // Capture host element parameters
                        var hostParameterValues = parameterSnapshotService.CaptureParams(structuralElement, whitelist);
                        
                        // Update existing zone with captured parameters
                        existingZone.MepParameterValues = mepParameterValues;
                        existingZone.HostParameterValues = hostParameterValues;
                        
                        // ✅ CRITICAL DIAGNOSTIC: Log parameter assignment with sample keys
                        if (!DeploymentConfiguration.DeploymentMode && (mepParameterValues.Count > 0 || hostParameterValues.Count > 0))
                        {
                            var mepSampleKeys = mepParameterValues.Take(5).Select(kv => kv?.Key ?? "null").Where(k => !string.IsNullOrEmpty(k)).ToList();
                            var mepSampleStr = mepSampleKeys.Count > 0 ? string.Join(", ", mepSampleKeys) : "none";
                            if (mepParameterValues.Count > mepSampleKeys.Count) mepSampleStr += $" (+{mepParameterValues.Count - mepSampleKeys.Count} more)";
                            
                             if (OptimizationFlags.UseDiagnosticMode)
                                _log($"[PARAM-SNAPSHOT] ✅ ASSIGNED {mepParameterValues.Count} MEP parameters and {hostParameterValues.Count} host parameters to existing ClashZone {existingZone.Id} (MEP={mepElement.Id}, Host={structuralElement.Id})");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Non-fatal: Log error but continue
                    _log($"[PARAM-SNAPSHOT] ⚠️ Error capturing parameters for existing ClashZone: {ex.Message}");
                }
            }
            
            using (var verifyOp = _performanceMonitor?.TrackOperation("6d4c. Sleeve Verification") as IDisposable)
            {
                // ✅ PHASE 3 OPTIMIZATION: Use O(1) opening map if provided, otherwise fallback to slow path
                bool hasExistingSleeve = false;
                string pointKey = $"{Math.Round(intersectionPoint.X, 3)}_{Math.Round(intersectionPoint.Y, 3)}_{Math.Round(intersectionPoint.Z, 3)}";

                if (existingOpeningPointKeys != null)
                {
                    hasExistingSleeve = existingOpeningPointKeys.Contains(pointKey);
                }
                else
                {
                    hasExistingSleeve = CheckForExistingSleeve(intersectionPoint, document);
                }

                if (hasExistingSleeve && !existingZone.IsResolved)
                {
                    existingZone.IsResolved = true;
                    // ✅ FIXED: Only set SleeveInstanceId to -1 if it's not already set (preserve from XML)
                    if (existingZone.SleeveInstanceId <= 0)
                    {
                        existingZone.SleeveInstanceId = -1; // Will be populated during placement if needed
                    }
                    if (OptimizationFlags.UseDiagnosticMode)
                        _log($"[UpdateExistingClashZone] Found existing individual sleeve at intersection point - set IsResolved=true for clash zone {existingZone.Id}, preserved SleeveInstanceId={existingZone.SleeveInstanceId}");
                }
                
                // ✅ PHASE 3 OPTIMIZATION: Check for existing cluster sleeve (using same pre-indexed map for now, though cluster usually has higher tolerance)
                bool hasExistingClusterSleeve = false;
                if (existingOpeningPointKeys != null)
                {
                    // NOTE: Cluster sleeves often have higher tolerance (50mm), but for streamlined we prioritize exact location matches
                    hasExistingClusterSleeve = existingOpeningPointKeys.Contains(pointKey);
                }
                else
                {
                    hasExistingClusterSleeve = CheckForExistingClusterSleeve(intersectionPoint, document);
                }

                if (hasExistingClusterSleeve && !existingZone.IsClusterResolved)
                {
                    existingZone.IsClusterResolved = true;
                    existingZone.ClusterSleeveInstanceId = -1; // Will be populated during clustering if needed
                     if (OptimizationFlags.UseDiagnosticMode)
                        _log($"[UpdateExistingClashZone] Found existing cluster sleeve at intersection point - set IsClusterResolved=true for clash zone {existingZone.Id}");
                }
                
                existingZone.LastUpdated = DateTime.Now;
            }
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
        
        /// <summary>
        /// ✅ CRITICAL FIX: Get MEP element size using strategy classes for insulation detection
        /// </summary>
        private MepElementSize GetMepElementSizeWithStrategy(Element mepElement, string mepCategory, System.Collections.Generic.Dictionary<string, string>? mepParamDict = null)
        {
            try
            {
                // ✅ PERFORMANCE OPT: Prevent string allocation for SafeLog when diagnostic mode is off
                bool diagMode = !DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode;
                if (diagMode) SafeLog("refresh.log", $"Analyzing element {mepElement.Id} with category '{mepCategory}'");
                
                // PERF OPT: Use cached strategy instances (all stateless except DamperPlacementStrategy which caches by doc)
                ISleevePlacementStrategy strategy;
                switch (mepCategory)
                {
                    case "Ducts":         strategy = _ductStrategy; break;
                    case "Pipes":         strategy = _pipeStrategy; break;
                    case "Cable Trays":   strategy = _cableTrayStrategy; break;
                    case "Duct Accessories":
                        // DamperPlacementStrategy holds a Document reference — reuse if same doc
                        var doc = mepElement.Document;
                        if (_cachedDamperStrategy == null || !ReferenceEquals(_cachedDamperStrategyDoc, doc))
                        {
                            _cachedDamperStrategy = new Strategies.DamperPlacementStrategy(doc);
                            _cachedDamperStrategyDoc = doc;
                        }
                        strategy = _cachedDamperStrategy;
                        break;
                    default:              strategy = _ductStrategy; break; // Default fallback
                }

                if (diagMode) SafeLog("refresh.log", $"Using strategy: {strategy.GetType().Name}");

                // Get MEP element size with insulation information - Pass pre-captured parameters
                var mepElementSize = strategy.GetMepElementSize(mepElement, mepParamDict);
                
                if (diagMode) SafeLog("refresh.log", $"✅ Strategy analysis complete: Shape='{mepElementSize.Shape}', IsInsulated={mepElementSize.IsInsulated}, InsulationThickness={mepElementSize.InsulationThickness:F6}ft");

                
                return mepElementSize;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in strategy analysis: {ex.Message}\n");
                }
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClashZoneService] Error getting MEP element size with strategy: {ex.Message}");
                }
                return new MepElementSize(); // Return empty size on error
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
            return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
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
                return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceInMm);
            }
            catch (Exception ex)
            {
                _log($"Error calculating clearance: {ex.Message}, using default");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0); // Fallback to 50mm converted to feet
            }
        }
        
        private bool IsPointNear(XYZ point1, XYZ point2, double tolerance)
        {
            return point1.DistanceTo(point2) <= tolerance;
        }
        
        /// <summary>
        /// Get MEP orientation direction (X or Y) based on host type
        /// For floors: derive from MEP element orientation vector
        /// For walls/framing: use wall direction type
        /// 
        /// ⚠️ CRITICAL PROTECTION: DO NOT MODIFY THIS METHOD ⚠️
        /// This method is essential for correct floor sleeve rotation:
        /// - Y-oriented ducts rotate 90° (width runs along Y-axis)
        /// - X-oriented ducts rotate 0° (width runs along X-axis)
        /// 
        /// Date Fixed: 2025-10-27
        /// Issue: MepElementOrientationDirection was incorrectly calculated for floors,
        ///        causing all floor sleeves to rotate incorrectly (all showing 0° rotation)
        /// Fix: For floors, derive X/Y from mepOrientation vector (which comes from MepElementOrientationHelper)
        ///      instead of using GetWallOrientationFromType (which is for walls/framing only)
        /// Status: WORKING - VERIFIED IN placement_debug.log (2025-10-27 20:11:40)
        /// </summary>
        public string GetMepOrientationDirection(string structuralElementType, XYZ mepOrientation, string wallDirectionType)
        {
            try
            {
                // ✅ CRITICAL: For floors, derive X or Y from MEP element's orientation vector
                // DO NOT CHANGE THIS LOGIC - FLOOR SLEEVE ROTATION DEPENDS ON IT
                if (structuralElementType == "Floor" || structuralElementType == "Floors")
                {
                    // Determine if MEP orientation is primarily X or Y
                    if (mepOrientation != null && mepOrientation != XYZ.Zero)
                    {
                        double absX = Math.Abs(mepOrientation.X);

                        double absY = Math.Abs(mepOrientation.Y);
                        
                        if (absX > absY)
                        {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENT-FIX] Floor host: MEP orientation has absX={absX:F3} > absY={absY:F3} → returning X\n");
                        }
                            return "X"; // X-oriented → 0° rotation
                        }
                        else
                        {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENT-FIX] Floor host: MEP orientation has absY={absY:F3} >= absX={absX:F3} → returning Y\n");
                        }
                            return "Y"; // Y-oriented → 90° rotation
                        }
                    }
                    else
                    {
                        // Default to X if orientation is zero
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENT-FIX] Floor host: MEP orientation is zero/null → returning X\n");
                        }
                        return "X";
                    }
                }
                else
                {
                    // ✅ CRITICAL: For walls and framing, use wall direction type
                    // DO NOT CHANGE THIS LOGIC - WALL/FRAMING ROTATION DEPENDS ON IT
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ORIENT-FIX] {structuralElementType} host: Using GetWallOrientationFromType({wallDirectionType})\n");
                    }
                    return GetWallOrientationFromType(wallDirectionType);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("refresh.log", $"[GetMepOrientationDirection] Error: {ex.Message}");
                }
                return "X"; // Default fallback
            }
        }
        
        /// <summary>
        /// ✅ FLOOR ROTATION: Calculate MEP element rotation angle in radians for floor sleeves
        /// Calculated once during refresh, used many times during placement (no Revit calls)
        /// For floors: 
        ///   - Horizontal elements: Projects MEP orientation onto XY plane and calculates angle using atan2
        ///   - Vertical elements: Gets rotation from element's transform BasisX/BasisY vectors
        /// For walls/framing: Returns 0 (rotation handled differently)
        /// </summary>
        public double CalculateMepElementRotationAngle(string structuralElementType, XYZ mepOrientation, Element mepElement)
        {
            try
            {
                // ✅ FLOOR ROTATION: Only calculate for floors (walls use different rotation logic)
                if (structuralElementType == "Floor" || structuralElementType == "Floors")
                {
                    if (mepOrientation != null && mepOrientation != XYZ.Zero)
                    {
                        // ✅ CRITICAL FIX: For vertical MEP elements through floors, ALWAYS use CoordinateSystem.BasisX
                        // mepOrientation from GetMepElementOrientation returns BasisX/BasisY (cardinal only), which loses 45° angles
                        // CoordinateSystem.BasisX gives the actual element rotation in XY plane (captures arbitrary angles)
                        try
                        {
                            Transform transform = null;
                            
                            // Try to get CoordinateSystem from connectors (works for both Duct and CableTray)
                            if (mepElement is Duct verticalDuct)
                            {
                                var connectors = verticalDuct.ConnectorManager?.Connectors;
                                if (connectors != null)
                                {
                                    foreach (Connector connector in connectors)
                                    {
                                        if (connector?.CoordinateSystem != null)
                                        {
                                            transform = connector.CoordinateSystem;
                                            break;
                                        }
                                    }
                                }
                            }
                            else if (mepElement is Autodesk.Revit.DB.Electrical.CableTray verticalCableTray)
                            {
                                // ✅ CABLE TRAY FIX: Get CoordinateSystem from cable tray connectors
                                var connectors = verticalCableTray.ConnectorManager?.Connectors;
                                if (connectors != null)
                                {
                                    foreach (Connector connector in connectors)
                                    {
                                        if (connector?.CoordinateSystem != null)
                                        {
                                            transform = connector.CoordinateSystem;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (transform != null)
                            {
                                XYZ basisX = transform.BasisX;
                                // Project BasisX onto XY plane (floors are horizontal) to get rotation angle
                                XYZ basisXProjected = new XYZ(basisX.X, basisX.Y, 0.0);
                                double basisXLength = Math.Sqrt(basisXProjected.X * basisXProjected.X + basisXProjected.Y * basisXProjected.Y);
                                
                                if (basisXLength > 1e-6)
                                {
                                    // Normalize and calculate angle - this captures arbitrary angles (45°, etc.)
                                    basisXProjected = new XYZ(basisXProjected.X / basisXLength, basisXProjected.Y / basisXLength, 0.0);
                                    double rotationAngle = Math.Atan2(basisXProjected.Y, basisXProjected.X);
                                    
                                    string elementType = mepElement is Duct ? "DUCT" : "CABLETRAY";
                                    SafeLog("refresh.log", $"[ROTATION-ANGLE] Floor host: VERTICAL {elementType} {mepElement.Id} - Using CoordinateSystem.BasisX ({basisX.X:F3}, {basisX.Y:F3}, {basisX.Z:F3}) projected to ({basisXProjected.X:F3}, {basisXProjected.Y:F3}) → rotation angle {rotationAngle * 180 / Math.PI:F1}°");
                                    
                                    return rotationAngle;
                                }
                            }
                        }
                        catch (Exception transformEx)
                        {
                                // ✅ PERFORMANCE OPT: Use SafeLog
                                SafeLog("refresh.log", $"[ROTATION-ANGLE] Error getting CoordinateSystem for vertical element {mepElement.Id}: {transformEx.Message}, falling back to orientation vector (may only give 0°/90°)");
                        }
                        
                        // ✅ FALLBACK: If CoordinateSystem fails, use orientation vector (may only give 0° or 90°)
                        // This happens when mepOrientation is BasisX/BasisY from GetMepElementOrientation helper
                        // Project MEP orientation onto XY plane (floors are horizontal)
                        // Calculate rotation angle using atan2(Y, X) - gives angle in radians
                        double fallbackRotationAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                        
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        SafeLog("refresh.log", $"[ROTATION-ANGLE] Floor host: Using fallback orientation vector ({mepOrientation.X:F3}, {mepOrientation.Y:F3}, {mepOrientation.Z:F3}) → rotation angle {fallbackRotationAngle * 180 / Math.PI:F1}° (may be limited to 0°/90° - CoordinateSystem unavailable)");
                        
                        return fallbackRotationAngle;
                    }
                    else
                    {
                        // Default to 0° if orientation is zero/null
                        // ✅ PERFORMANCE OPT: Use SafeLog
                        SafeLog("refresh.log", $"[ROTATION-ANGLE] Floor host: MEP orientation is zero/null → returning 0°");
                        return 0.0;
                    }
                }
                else
                {
                    // ✅ WALLS/FRAMING: Rotation handled differently (not based on angle)
                    // Return 0 - wall rotation uses MepElementOrientationDirection logic
                    return 0.0;
                }
            }
            catch (Exception ex)
            {
                // ✅ PERFORMANCE OPT: Use SafeLog
                SafeLog("refresh.log", $"[CalculateMepElementRotationAngle] Error: {ex.Message}");
                return 0.0; // Default fallback
            }
        }

        /// <summary>
        /// Helper method to safely check if a parameter has a large value (> threshold mm)
        /// </summary>
        private bool IsParameterLargeValue(Parameter? param, double thresholdMm = 300)
        {
            if (param == null || !param.HasValue) 
                return false;
            
            try
            {
                var value = RevitUnitConversionService.Instance.FromInternalMillimeters(param.AsDouble());
                return value > thresholdMm;
            }
            catch
            {
                return false;
            }
        }
        
        #endregion

        // ─────────────────────────────────────────────────────────────────────
        // PERF OPT: Cache key generation for linked document support
        // ElementIds are only unique within a document. For linked files, we need
        // composite keys that include document path to avoid collisions.
        // ─────────────────────────────────────────────────────────────────────
        private static string GetElementCacheKey(Element element)
        {
            if (element == null) return string.Empty;
            // UniqueId is globally unique across all documents (includes document GUID internally).
            // PathName was unreliable for linked files — can be "" or vary between calls.
            return element.UniqueId ?? $"{element.Id.GetIntegerValue()}|fallback";
        }

        // ✅ DIAGNOSTIC: Track expensive fallback calls
        private static int _fallbackHostOrientationCount = 0;
        private string GetHostOrientationWithTracking(Element structuralElement)
        {
            _fallbackHostOrientationCount++;
            if (_fallbackHostOrientationCount <= 5 || _fallbackHostOrientationCount % 50 == 0)
            {
                SafeLog("refresh.log", $"[FALLBACK-BACK] GetHostOrientation called #{_fallbackHostOrientationCount} for Element {structuralElement?.Id}");
            }
            return WallDirectionService.GetHostOrientation(structuralElement);
        }

        // ─────────────────────────────────────────────────────────────────────
        // PERF OPT: Per-MEP-element pre-computed data cache
        // All fields are intrinsic to the element (not intersection-dependent).
        // Pre-computed once per unique MEP element ID before the zone loop.
        // NOTE: Uses string keys for linked document support (ElementId|DocPath)
        // ─────────────────────────────────────────────────────────────────────
        // PERF OPT: Per-host-element pre-computed data cache
        // Same host wall can intersect many MEP elements — compute once per unique host ID.
        // NOTE: Uses string keys for linked document support (ElementId|DocPath)
        private sealed class HostElementComputedData
        {
            public string StructuralElementType;       // "Wall", "Floor", "Structural Framing"
            public double StructuralElementThickness;  // from GetElementThickness
            public double WallThickness;               // from GetWallThickness
            public double FramingThickness;            // from GetFramingThickness
            public XYZ StructuralElementNormal;        // from WallDirectionService.GetStructuralElementNormal
            public string HostOrientation;             // from WallDirectionService.GetHostOrientation ("X" or "Y")
            public string GeometryHash;                // ✅ PATH 2 OPT: Pre-computed geometry hash to avoid string alloc per zone
        }

        private sealed class MepElementComputedData
        {
            public string Category;
            public double RawWidth;
            public double RawHeight;
            public MepElementSize StrategySize;   // from GetMepElementSizeWithStrategy
            public bool StrategyValid;             // true when strategy returned usable dimensions
            public double FinalWidth;              // resolved from strategy (wall-orient-independent path)
            public double FinalHeight;
            public double FinalDiameter;
            public string DuctShape;
            public bool IsInsulated;
            public double InsulationThickness;
            public string InsulationType;          // "Insulated" / "Normal"
            public string PipeOpeningType;
            public string LevelName;
            public double LevelElevation;
            public double PipeOuterDiameter;
            public double PipeNominalDiameter;
            public string FormattedSize;           // null when strategy invalid (bbox fallback needed)
            public string SizeParameterValue;
            public string SystemAbbreviation;
            public string SystemType;
            public string SystemName;
            public string ServiceType;
            public double ElevationFromLevel;
            public string TypeName;    // mepElement type name — cached to avoid GetTypeId+GetElement per zone
            public string FamilyName;  // mepElement family name
            public string UniqueId;    // ✅ PATH 2 OPT: MepElement.UniqueId cached to avoid COM call per zone
            public string GeometryHash; // ✅ PATH 2 OPT: Pre-computed geometry hash to avoid string alloc per zone
            // ✅ NEW: Cache expensive per-zone calculations (called 384 times, should be 247 times)
            public string MepOrientationDirection;  // From GetMepOrientationDirection
            public double? MepRotationAngle;         // From CalculateMepElementRotationAngle (Nullable for on-demand)
            public double MepAngleToXRad;
            public double MepAngleToXDeg;
            public double MepAngleToYRad;
            public double MepAngleToYDeg;
        }

        private MepElementComputedData ComputeMepElementData(Element mepElement, Dictionary<string, string>? mepParamDict)
        {
            var d = new MepElementComputedData();
            d.Category = GetElementCategoryName(mepElement, mepParamDict);
            (d.RawWidth, d.RawHeight) = GetMepElementDimensions(mepElement, mepParamDict);
            d.StrategySize = GetMepElementSizeWithStrategy(mepElement, d.Category, mepParamDict);

            // Resolve final dimensions from strategy (wall-orientation-independent path)
            double fw = d.RawWidth, fh = d.RawHeight, fd = 0.0;
            d.StrategyValid = d.StrategySize.Width > 0 || d.StrategySize.Height > 0 || d.StrategySize.Diameter > 0;
            if (d.StrategyValid)
            {
                fw = d.StrategySize.Width  > 0 ? d.StrategySize.Width  : fw;
                fh = d.StrategySize.Height > 0 ? d.StrategySize.Height : fh;
                fd = d.StrategySize.Diameter;
            }
            d.FinalWidth  = fw;
            d.FinalHeight = fh;
            d.FinalDiameter = fd;

            d.DuctShape = GetDuctShape(mepElement, mepParamDict);
            
            // ✅ INSULATION OPTIMIZATION: Check parameter cache first
            if (mepParamDict != null && mepParamDict.TryGetValue("Insulation Thickness", out var insStr) && 
                double.TryParse(insStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var insVal))
            {
                d.IsInsulated = insVal > 0.001; // ~0.3mm threshold
                d.InsulationThickness = insVal;
            }
            else
            {
                (d.IsInsulated, d.InsulationThickness) = _insulationDetector.GetInsulationInfo(mepElement, d.StrategySize);
            }
            
            d.InsulationType = d.IsInsulated ? "Insulated" : "Normal";
            d.PipeOpeningType = GetPipeOpeningType(mepElement);
            (d.LevelName, d.LevelElevation) = GetMepElementLevelInfo(mepElement, mepParamDict);

            // ✅ PIPE DIAMETER OPTIMIZATION: Check parameter cache first
            if (mepElement is Pipe pipe)
            {
                if (mepParamDict != null && mepParamDict.TryGetValue("Outside Diameter", out var odStr) && 
                    double.TryParse(odStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var odVal))
                {
                    d.PipeOuterDiameter = odVal;
                }
                else
                {
                    var odp = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    if (odp != null) d.PipeOuterDiameter = odp.AsDouble();
                }

                if (mepParamDict != null && mepParamDict.TryGetValue("Nominal Diameter", out var ndStr) && 
                    double.TryParse(ndStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var ndVal))
                {
                    d.PipeNominalDiameter = ndVal;
                }
                else
                {
                    var ndp = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                    if (ndp != null) d.PipeNominalDiameter = ndp.AsDouble();
                }
            }

            // Pre-compute formatted size only when strategy gave valid dims (bbox-fallback path is wall-orient-dependent)
            d.FormattedSize = d.StrategyValid
                ? GetMepElementSizeString(mepElement, fw, fh, d.DuctShape, d.PipeNominalDiameter, mepParamDict)
                : null;

            d.SizeParameterValue  = GetMepElementSizeParameterValue(mepElement, mepParamDict);
            d.SystemAbbreviation  = GetMepSystemAbbreviation(mepElement, mepParamDict);
            d.SystemType          = GetMepSystemType(mepElement, d.Category, mepParamDict);
            
            // ✅ SYSTEM NAME OPTIMIZATION: Check parameter cache first
            if (mepParamDict != null && mepParamDict.TryGetValue("System Name", out var sn))
            {
                d.SystemName = sn;
            }
            else
            {
                try
                {
                    var sysNameParam = mepElement.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
                    if (sysNameParam != null && sysNameParam.HasValue) d.SystemName = sysNameParam.AsString() ?? string.Empty;
                    else d.SystemName = string.Empty;
                }
                catch { d.SystemName = string.Empty; }
            }

            d.ServiceType       = GetMepServiceType(mepElement, mepParamDict);
            d.ElevationFromLevel = GetMepElementOffset(mepElement, mepParamDict);
            
            // ✅ PERSISTED ANGLES: Extract from parameters if available (whitelisted in ParameterSnapshotService)
            if (mepParamDict != null)
            {
                // Angle to X
                if (mepParamDict.TryGetValue("MepAngleToXRad", out var axr) && double.TryParse(axr, NumberStyles.Any, CultureInfo.InvariantCulture, out var axrv)) d.MepAngleToXRad = axrv;
                if (mepParamDict.TryGetValue("MepAngleToXDeg", out var axd) && double.TryParse(axd, NumberStyles.Any, CultureInfo.InvariantCulture, out var axdv)) d.MepAngleToXDeg = axdv;
                else if (mepParamDict.TryGetValue("Angle to X", out var ax) && double.TryParse(ax, NumberStyles.Any, CultureInfo.InvariantCulture, out var axv)) d.MepAngleToXDeg = axv;
                else if (mepParamDict.TryGetValue("Angle to X-axis", out var ax2) && double.TryParse(ax2, NumberStyles.Any, CultureInfo.InvariantCulture, out var axv2)) d.MepAngleToXDeg = axv2;

                // Angle to Y
                if (mepParamDict.TryGetValue("MepAngleToYRad", out var ayr) && double.TryParse(ayr, NumberStyles.Any, CultureInfo.InvariantCulture, out var ayrv)) d.MepAngleToYRad = ayrv;
                if (mepParamDict.TryGetValue("MepAngleToYDeg", out var ayd) && double.TryParse(ayd, NumberStyles.Any, CultureInfo.InvariantCulture, out var aydv)) d.MepAngleToYDeg = aydv;
                else if (mepParamDict.TryGetValue("Angle to Y", out var ay) && double.TryParse(ay, NumberStyles.Any, CultureInfo.InvariantCulture, out var ayv)) d.MepAngleToYDeg = ayv;
                else if (mepParamDict.TryGetValue("Angle to Y-axis", out var ay2) && double.TryParse(ay2, NumberStyles.Any, CultureInfo.InvariantCulture, out var ayv2)) d.MepAngleToYDeg = ayv2;
                
                // If we only have degrees, convert to radians for structured storage consistency
                if (d.MepAngleToXRad == 0 && d.MepAngleToXDeg != 0) d.MepAngleToXRad = d.MepAngleToXDeg * Math.PI / 180.0;
                if (d.MepAngleToYRad == 0 && d.MepAngleToYDeg != 0) d.MepAngleToYRad = d.MepAngleToYDeg * Math.PI / 180.0;
            }

            // ✅ TYPE INFO OPTIMIZATION: Check parameter cache first
            if (mepParamDict != null && mepParamDict.TryGetValue("Type Name", out var tn)) d.TypeName = tn;
            if (mepParamDict != null && mepParamDict.TryGetValue("Family Name", out var fn)) d.FamilyName = fn;

            // ✅ USER REQUESTED OPTIMIZATION: Type Name and Family Name are only needed for Accessories (Dampers/Valves).
            // For other categories (Ducts, Pipes, Trays), skipping this saves expensive doc.GetElement(typeId) calls.
            bool isAccessory = d.Category == "Duct Accessories" || d.Category == "Pipe Accessories";

            if (isAccessory && (string.IsNullOrEmpty(d.TypeName) || string.IsNullOrEmpty(d.FamilyName)))
            {
                try
                {
                    var typeId = mepElement.GetTypeId();
                    if (typeId != null && typeId != ElementId.InvalidElementId)
                    {
                        var typeElem = mepElement.Document?.GetElement(typeId);
                        if (typeElem != null)
                        {
                            d.TypeName = typeElem.Name;
                            if (typeElem is ElementType et) d.FamilyName = et.FamilyName;
                        }
                    }
                }
                catch { }
            }
            
            // ✅ PATH 2 OPTIMIZATION: Pre-compute expensive properties to avoid per-zone COM calls
            try
            {
                d.UniqueId = mepElement.UniqueId ?? string.Empty;  
                d.GeometryHash = $"{mepElement.Id.GetIntegerValue()}_{d.Category ?? "Unknown"}";  
            }
            catch { }

            // ✅ OPTIMIZATION: Rotation is calculated in batch only for Floor hosts (vertical elements).
            // Leave null here — the batch phase fills it in for Floor penetrations only.
            d.MepRotationAngle = null;

            return d;
        }
    }
}
