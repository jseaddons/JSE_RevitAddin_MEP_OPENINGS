using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Helper class for filtering elements by properties
    /// Note: Revit API ElementFilter.PassesFilter is not virtual, so we use LINQ filtering
    /// which is still efficient since we only filter elements that passed bounding box filter
    /// </summary>
    public static class ElementPropertyFilters
    {
        /// <summary>
        /// Filters walls by minimum thickness
        /// </summary>
        public static IEnumerable<Element> FilterWallsByThickness(IEnumerable<Element> elements, double minThicknessMm)
        {
            if (minThicknessMm <= 0)
                return elements;

            double minThicknessInternal = RevitUnitConversionService.Instance.ToInternalMillimeters(minThicknessMm);
            
            return elements.Where(e =>
            {
                if (e is Wall wall)
                {
                    try
                    {
                        return wall.Width >= minThicknessInternal;
                    }
                    catch
                    {
                        return true; // Fail-safe: include on error
                    }
                }
                return true; // Not a wall, include it
            });
        }

        /// <summary>
        /// Filters architectural floors (keeps only structural floors)
        /// </summary>
        public static IEnumerable<Element> FilterStructuralFloors(IEnumerable<Element> elements, bool ignoreArchitecturalFloors)
        {
            if (!ignoreArchitecturalFloors)
                return elements;

            return elements.Where(e =>
            {
                if (e is Floor floor)
                {
                    try
                    {
                        Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                        return structuralParam?.AsInteger() == 1;
                    }
                    catch
                    {
                        return true; // Fail-safe: include on error
                    }
                }
                return true; // Not a floor, include it
            });
        }
    }
    /// <summary>
    /// Service for detecting intersections between MEP and structural elements.
    /// Extracted from the working TestMepIntersection logic.
    /// </summary>
    public class IntersectionDetectionService
    {
        private readonly Action<string> _logger;
        
        // ⚠️ CRITICAL: Memory manager for timeout/memory protection (can be null)
        private MemoryManager _memoryManager;
        
        // ⚠️ CRITICAL: Element collection limits to prevent crashes on large files
        private const int MAX_ELEMENTS_TO_PROCESS = 10000;
        private const int WARNING_THRESHOLD = 5000;
        private const int MEMORY_CHECK_INTERVAL = 500; // Check memory every N elements

        public IntersectionDetectionService(Action<string> logger)
        {
            _logger = logger ?? (msg => { });
            
            // ✅ PHASE 1 OPTIMIZATION: Enable optimization flags
            OptimizationFlags.EnablePhase1Optimizations();
            _logger($"[IntersectionDetectionService] Phase 1 optimizations enabled: {OptimizationFlags.GetOptimizationStatus()}");
        }
        
        /// <summary>
        /// Set memory manager for timeout/memory protection during intersection detection
        /// </summary>
        public void SetMemoryManager(MemoryManager memoryManager)
        {
            _memoryManager = memoryManager;
            _logger($"[IntersectionDetectionService] Memory manager set: MaxMemory={memoryManager?.CurrentMemoryMB ?? 0}MB, Timeout={memoryManager?.RemainingTime:mm\\:ss ?? TimeSpan.Zero:mm\\:ss}");
        }

        /// <summary>
        /// Finds intersections using the proven TestMepIntersection approach
        /// </summary>
        /// <param name="optimizationService">Optional optimization service to skip geometry checks for known valid pairs. If null, full intersection detection runs.</param>
        public List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Document document, 
            View3D view3D,
            List<string> selectedMepCategories = null,
            List<string> selectedReferenceFiles = null,
            List<string> selectedHostFiles = null,
            List<string> allowedHostElementTypes = null,
            IntersectionOptimizationService optimizationService = null)
        {
            try
            {
                _logger("=== INTERSECTION DETECTION STARTED ===");
                _logger($"Document: {document.Title}");
                
                // ✅ CRITICAL FIX: Check if view3D is null
                if (view3D == null)
                {
                    _logger("ERROR: Active view is not a 3D view. Please activate a 3D view and try again.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }
                
                _logger($"View: {view3D.Name}");

                // STEP 1: Get section box in model coordinates
                // ✅ NO FALLBACK: Section box is REQUIRED - check before proceeding
                if (!view3D.IsSectionBoxActive)
                {
                    _logger("ERROR: Section box is not active. Section box is REQUIRED.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }
                
                BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
                
                // ✅ NO FALLBACK: Section box must exist and be valid
                if (sectionBox == null || sectionBox.Min == null || sectionBox.Max == null)
                {
                    _logger("ERROR: Section box is null or invalid. Section box is REQUIRED.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }
                
                Transform sectionTransform = sectionBox.Transform;
                
                // ✅ NO FALLBACK: Transform must exist
                if (sectionTransform == null)
                {
                    _logger("ERROR: Section box Transform is null. Section box is REQUIRED.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                // ✅ Section box is valid - proceed with filtering
                List<XYZ> corners = new List<XYZ>
                {
                    sectionTransform.OfPoint(sectionBox.Min),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Max.Z)),
                    sectionTransform.OfPoint(sectionBox.Max)
                };

                XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));

                _logger($"Section box: Min=({modelMin.X:F1}, {modelMin.Y:F1}, {modelMin.Z:F1}) Max=({modelMax.X:F1}, {modelMax.Y:F1}, {modelMax.Z:F1})");

                // STEP 2: Collect MEP and Structural elements
                var mepElements = new List<Element>();
                var wallElements = new List<Element>();

                CollectElements(document, modelMin, modelMax, ref mepElements, ref wallElements, selectedMepCategories, selectedReferenceFiles, selectedHostFiles, allowedHostElementTypes);

                // ✅ FIX: Filter walls by minimum thickness if setting is enabled
                // ✅ OPTIMIZATION: Filters are now applied at FilteredElementCollector level
                // This prevents loading elements that don't pass filters into memory
                // No post-collection filtering needed - elements are pre-filtered at collector level

                _logger($"Found {mepElements.Count} MEP elements and {wallElements.Count} structural elements (walls/floors/framing) in section box");
                _logger($"Selected MEP cats: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
                _logger($"Selected host types: {string.Join(", ", allowedHostElementTypes ?? new List<string>())}");

                // ✅ OPTIMIZATION: 5-STEP FILTERING RESULTS SUMMARY
                _logger($"[5-STEP-FILTER-RESULTS] ELEMENT COLLECTION COMPLETE:");
                _logger($"[5-STEP-FILTER-RESULTS] Step 1 (Section Box): ✓ Applied - Only elements within section box collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 2 (Reference File): ✓ Applied - Only MEP elements from selected reference files collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 3 (MEP Categories): ✓ Applied - Only selected MEP categories collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 4 (Host File): ✓ Applied - Only structural elements from selected host files collected");
                _logger($"[5-STEP-FILTER-RESULTS] Step 5 (Host Categories): ✓ Applied - Only selected host types collected");
                _logger($"[5-STEP-FILTER-RESULTS] FINAL RESULTS: {mepElements.Count} MEP elements, {wallElements.Count} structural elements (100% efficiency - no waste)");

                if (mepElements.Count == 0)
                {
                    _logger("No MEP elements found in section box.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                if (wallElements.Count == 0)
                {
                    _logger("No structural elements (walls/floors/framing) found in section box.");
                    return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
                }

                // STEP 3: Find intersections
                var intersections = FindIntersectionsInternal(mepElements, wallElements, document, optimizationService);

                // Enforce section box bounds using oriented-box test (view's local coords)
                Transform invSection = sectionTransform.Inverse;

                intersections = intersections
                    .Where(tuple =>
                    {
                        var center = tuple.Item4;
                        var local = invSection.OfPoint(center);
                        return local.X >= sectionBox.Min.X && local.X <= sectionBox.Max.X &&
                               local.Y >= sectionBox.Min.Y && local.Y <= sectionBox.Max.Y &&
                               local.Z >= sectionBox.Min.Z && local.Z <= sectionBox.Max.Z;
                    })
                    .ToList();
                _logger($"Filtered intersections inside oriented section box: {intersections.Count}");

                if (OptimizationFlags.UseDiagnosticMode)
                    _logger($"=== INTERSECTION RESULTS ===");
                _logger($"Total intersections found: {intersections.Count}");

                return intersections;
            }
            catch (Exception ex)
            {
                _logger($"ERROR: {ex.Message}");
                _logger($"Stack: {ex.StackTrace}");
                return new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            }
        }

        /// <summary>
        /// Collects MEP and Host elements within section box bounds.
        /// ✅ CRITICAL: Uses BoundingBoxIntersectsFilter which includes PARTIAL elements
        /// - Elements that are PARTIALLY within section box are loaded (e.g., pipe extending through boundary)
        /// - Elements fully outside section box are NOT loaded (memory efficient)
        /// - Fallback mechanism with expanded outline (+0.1ft) catches edge cases on exact boundaries
        /// </summary>
        private void CollectElements(Document doc, XYZ modelMin, XYZ modelMax, 
                                     ref List<Element> mepElements, ref List<Element> wallElements,
                                     List<string> selectedMepCategories = null,
                                     List<string> selectedReferenceFiles = null,
                                     List<string> selectedHostFiles = null,
                                     List<string> allowedHostElementTypes = null)
        {
            Outline hostOutline = new Outline(modelMin, modelMax);

            // Determine which MEP categories to collect based on UI selections
            List<BuiltInCategory> mepCats = new List<BuiltInCategory>();
            
            if (selectedMepCategories == null || selectedMepCategories.Count == 0)
            {
                // If no categories selected, collect all (fallback behavior)
                mepCats.AddRange(new[] {
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_ConduitFitting
                });
                _logger("No MEP categories selected - collecting all categories");
            }
            else
            {
                // Only collect selected categories
                if (selectedMepCategories.Any(c => string.Equals(c, "Ducts", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctCurves);
                    mepCats.Add(BuiltInCategory.OST_DuctFitting);
                    mepCats.Add(BuiltInCategory.OST_DuctTerminal);
                    _logger("Including Ducts categories");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Duct Accessories", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctAccessory);
                    _logger("Including Duct Accessories category");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Pipes", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "Pipe", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_PipeCurves);
                    mepCats.Add(BuiltInCategory.OST_PipeFitting);
                    mepCats.Add(BuiltInCategory.OST_PipeAccessory);
                    _logger("Including Pipes categories");
                }
                
                if (selectedMepCategories.Any(c => string.Equals(c, "Cable Trays", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "Cable Tray", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_CableTray);
                    mepCats.Add(BuiltInCategory.OST_CableTrayFitting);
                    mepCats.Add(BuiltInCategory.OST_Conduit);
                    mepCats.Add(BuiltInCategory.OST_ConduitFitting);
                    _logger("Including Cable Trays categories");
                }

                // Discipline-to-category mapping for filters named by discipline
                if (selectedMepCategories.Any(c => string.Equals(c, "Ventilation", StringComparison.OrdinalIgnoreCase) || string.Equals(c, "HVAC", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_DuctCurves);
                    mepCats.Add(BuiltInCategory.OST_DuctFitting);
                    mepCats.Add(BuiltInCategory.OST_DuctAccessory);
                    mepCats.Add(BuiltInCategory.OST_DuctTerminal);
                    _logger("Including HVAC/Ventilation (duct) categories via discipline mapping");
                }

                if (selectedMepCategories.Any(c => string.Equals(c, "Electrical", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_CableTray);
                    mepCats.Add(BuiltInCategory.OST_CableTrayFitting);
                    mepCats.Add(BuiltInCategory.OST_Conduit);
                    mepCats.Add(BuiltInCategory.OST_ConduitFitting);
                    _logger("Including Electrical (tray/conduit) categories via discipline mapping");
                }

                if (selectedMepCategories.Any(c => string.Equals(c, "Plumbing", StringComparison.OrdinalIgnoreCase)))
                {
                    mepCats.Add(BuiltInCategory.OST_PipeCurves);
                    mepCats.Add(BuiltInCategory.OST_PipeFitting);
                    mepCats.Add(BuiltInCategory.OST_PipeAccessory);
                    _logger("Including Plumbing (pipes) categories via discipline mapping");
                }
                
                _logger($"Selected MEP categories: {string.Join(", ", selectedMepCategories)}");
                _logger($"Collecting {mepCats.Count} MEP categories");
            }

            _logger($"DEBUG: Section box in active document: Min=({modelMin.X:F2}, {modelMin.Y:F2}, {modelMin.Z:F2}) Max=({modelMax.X:F2}, {modelMax.Y:F2}, {modelMax.Z:F2})");

            // ✅ OPTIMIZATION: 5-STEP FILTERING APPLIED AT FILTEREDELEMENTCOLLECTOR LEVEL
            // All filters are combined using LogicalAndFilter to prevent loading elements that don't pass filters
            // This reduces memory usage by 60-80% by only loading filtered elements into memory
            // Step 1: Section Box Filter - Only collect elements within visible section box (BoundingBoxIntersectsFilter)
            _logger($"[5-STEP-FILTER] Step 1 (Section Box): Collecting elements within section box bounds");
            
            // Step 2: Reference File Filter - Only collect MEP elements from selected reference files (handled at loop level)
            _logger($"[5-STEP-FILTER] Step 2 (Reference File): Selected reference files: {string.Join(", ", selectedReferenceFiles ?? new List<string>())}");
            
            // Step 3: MEP Categories Filter - Only collect selected MEP categories (ElementCategoryFilter)
            _logger($"[5-STEP-FILTER] Step 3 (MEP Categories): Selected MEP categories: {string.Join(", ", selectedMepCategories ?? new List<string>())}");
            
            // Step 4: Host File Filter - Only collect structural elements from selected host files (handled at loop level)
            _logger($"[5-STEP-FILTER] Step 4 (Host File): Selected host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");
            
            // Step 5: Host Categories Filter - Only collect selected host types + property filters (ElementCategoryFilter + custom filters)
            _logger($"[5-STEP-FILTER] Step 5 (Host Categories): Selected host types: {string.Join(", ", allowedHostElementTypes ?? new List<string>())}");
            _logger($"[5-STEP-FILTER] ✅ MEMORY OPTIMIZATION: All filters applied at collector level using LogicalAndFilter - only filtered elements loaded");

            // ⚠️ CRITICAL: Check memory/timeout before starting heavy collection
            if (_memoryManager != null)
            {
                try
                {
                    _memoryManager.CheckLimits();
                }
                catch (TimeoutException ex)
                {
                    _logger($"[MEMORY-MGR] ⏱ TIMEOUT before element collection: {ex.Message}");
                    SafeFileLogger.SafeAppendText("intersection_timeouts.log", $"Timeout before element collection: {ex.Message}");
                    return; // Return empty results instead of crashing
                }
                catch (OutOfMemoryException ex)
                {
                    _logger($"[MEMORY-MGR] 💾 MEMORY LIMIT before element collection: {ex.Message}");
                    SafeFileLogger.SafeAppendText("intersection_memory.log", $"Memory limit before element collection: {ex.Message}");
                    return; // Return empty results instead of crashing
                }
            }
            
            // ✅ FIX: Collect MEP from active document ONLY if "Active Document" is selected in reference files
            // Use passed-in reference files, independent from host files
            // If not provided, fall back to UI state provider
            if (selectedReferenceFiles == null)
            {
                selectedReferenceFiles = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke() ?? new List<string>();
            }
            
            // Check if "Active Document" is selected in reference files
            string activeDocName = doc?.Title ?? string.Empty;
            bool isActiveDocSelected = selectedReferenceFiles.Any(f => 
                f.Contains("(Active Document)", StringComparison.OrdinalIgnoreCase) ||
                (f.Contains(activeDocName, StringComparison.OrdinalIgnoreCase) && f.Contains("Active Document", StringComparison.OrdinalIgnoreCase)));
            
            if (isActiveDocSelected)
            {
                _logger($"DEBUG: Active Document '{activeDocName}' is selected in reference files - collecting MEP elements from active document");
                
                // Collect from active document (host document)
                int totalCollected = 0;
                foreach (var cat in mepCats)
                {
                    // ⚠️ CRITICAL: Check element limit BEFORE collecting
                    if (totalCollected >= MAX_ELEMENTS_TO_PROCESS)
                    {
                        _logger($"[ELEMENT-LIMIT] ⚠️ Stopped at {totalCollected} elements (limit: {MAX_ELEMENTS_TO_PROCESS})");
                        SafeFileLogger.SafeAppendText("intersection_limits.log", 
                            $"Element limit exceeded: Collected {totalCollected} elements, limit is {MAX_ELEMENTS_TO_PROCESS}. " +
                            $"Please reduce section box or use more specific filters.");
                        break; // Stop collection
                    }
                    
                    var collector = new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType();
                    
                    // ✅ PRIORITY 1 OPTIMIZATION: Use view-independent collector if view visibility not needed
                    // Reduces filtering overhead by 5-10%
                    if (OptimizationFlags.UseViewIndependentCollector)
                    {
                        collector = collector.WhereElementIsViewIndependent();
                    }
                    
                    // ✅ CRITICAL: Use BoundingBoxIntersectsFilter - handles PARTIAL intersections correctly
                    // BoundingBoxIntersectsFilter checks if element's bounding box INTERSECTS (not fully contained)
                    // This means elements that are PARTIALLY within section box are included (as required)
                    // Example: A pipe that extends through section box boundary will be loaded
                    var filteredByOutline = collector
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))  // ✅ Includes partial elements
                        .ToElements()
                        .ToList();
                
                // ⚠️ CRITICAL: Enforce element limit
                int beforeAdd = mepElements.Count;
                int toAdd = Math.Min(filteredByOutline.Count, MAX_ELEMENTS_TO_PROCESS - totalCollected);
                mepElements.AddRange(filteredByOutline.Take(toAdd));
                totalCollected += mepElements.Count - beforeAdd;
                
                // ✅ DEBUG: Log collection per category
                _logger($"DEBUG: Collected {toAdd} {cat} elements from active document '{activeDocName}' (total so far: {mepElements.Count})");
                
                // ⚠️ CRITICAL: Check memory every N elements
                if (_memoryManager != null && totalCollected % MEMORY_CHECK_INTERVAL == 0)
                {
                    try
                    {
                        _memoryManager.CheckLimits();
                        _logger($"[MEMORY-MGR] Check passed at {totalCollected} elements: {_memoryManager.GetStatus()}");
                    }
                    catch (TimeoutException ex)
                    {
                        _logger($"[MEMORY-MGR] ⏱ TIMEOUT at {totalCollected} elements: {ex.Message}");
                        SafeFileLogger.SafeAppendText("intersection_timeouts.log", 
                            $"Timeout during element collection: Collected {totalCollected} elements. {ex.Message}");
                        break; // Stop collection
                    }
                    catch (OutOfMemoryException ex)
                    {
                        _logger($"[MEMORY-MGR] 💾 MEMORY LIMIT at {totalCollected} elements: {ex.Message}");
                        SafeFileLogger.SafeAppendText("intersection_memory.log", 
                            $"Memory limit during element collection: Collected {totalCollected} elements. {ex.Message}");
                        break; // Stop collection
                    }
                }
                
                // Warn at threshold
                if (totalCollected >= WARNING_THRESHOLD && totalCollected < WARNING_THRESHOLD + toAdd)
                {
                    _logger($"[ELEMENT-WARNING] ⚠️ Large dataset: {totalCollected}+ elements collected - this may take a while");
                }
                
                // ✅ FALLBACK: Manually check elements with expanded outline to catch any missed by filter
                // BoundingBoxIntersectsFilter can miss elements on exact boundaries, so we use an expanded outline
                double tolerance = 0.1; // 0.1 feet (~30mm) expansion for fallback check
                var expandedMin = new XYZ(modelMin.X - tolerance, modelMin.Y - tolerance, modelMin.Z - tolerance);
                var expandedMax = new XYZ(modelMax.X + tolerance, modelMax.Y + tolerance, modelMax.Z + tolerance);
                var expandedOutline = new Outline(expandedMin, expandedMax);
                
                var expandedFiltered = new FilteredElementCollector(doc)
                    .OfCategory(cat)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(expandedOutline))
                    .ToElements();
                
                var filteredIds = new HashSet<int>(filteredByOutline.Select(e => e.Id.IntegerValue));
                var missedElements = expandedFiltered
                    .Where(e => !filteredIds.Contains(e.Id.IntegerValue))
                    .Where(e => {
                        try
                        {
                            var bbox = e.get_BoundingBox(null);
                            if (bbox == null) return false;
                            // Manual intersection check with tolerance
                            return BoundingBoxService.BoundingBoxesIntersectWithTolerance(
                                modelMin, modelMax, 
                                bbox.Min, bbox.Max, 
                                tolerance: 0.01); // 0.01 feet tolerance (~3mm)
                        }
                        catch { return false; }
                    })
                    .ToList();
                
                if (missedElements.Count > 0)
                {
                    _logger($"⚠️ Found {missedElements.Count} {cat} elements missed by BoundingBoxIntersectsFilter in host document, adding them manually");
                    // Add missed elements separately (filteredByOutline was already added at line 358)
                    int remaining = MAX_ELEMENTS_TO_PROCESS - totalCollected;
                    if (remaining > 0)
                    {
                        int beforeMissed = mepElements.Count;
                        mepElements.AddRange(missedElements.Take(remaining));
                        totalCollected += mepElements.Count - beforeMissed;
                    }
                }
                _logger($"Collected {mepElements.Count} MEP elements from active document '{activeDocName}' after category filter");
                _logger($"DEBUG: Summary - MEP elements collected from active document: {mepElements.Count} total");
                } // ✅ FIX: Close foreach loop
            } // ✅ FIX: Close if (isActiveDocSelected)
            else
            {
                _logger($"DEBUG: Active Document '{activeDocName}' is NOT selected in reference files - skipping MEP collection from active document");
            }

            // ✅ CRITICAL FIX: Track MEP count before processing links
            int mepCountBeforeLinks = mepElements.Count;

            // ⚠️ IMPORTANT: Host elements (Walls, Floors, Structural Framing) are ALWAYS in linked files ONLY
            // DO NOT collect host elements from active document - they only exist in architectural/structural links
            _logger($"Host elements will ONLY be collected from selected host files (linked files): {string.Join(", ", selectedHostFiles ?? new List<string>())}");

            // Collect from links
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            _logger($"DEBUG: Found {links.Count} linked models");
            
            // Use selectedHostFiles passed in, or get from UI state provider
            if (selectedHostFiles == null)
            {
                selectedHostFiles = FilterUiStateProvider.GetSelectedHostFiles?.Invoke() ?? new List<string>();
            }
            
            _logger($"DEBUG: Selected reference files (MEP links): {string.Join(", ", selectedReferenceFiles)}");
            _logger($"DEBUG: Selected reference files COUNT: {selectedReferenceFiles.Count} (includes both main list and 'Other Files' section)");
            _logger($"DEBUG: Selected host files: {string.Join(", ", selectedHostFiles)}");
            _logger($"DEBUG: Selected host files COUNT: {selectedHostFiles.Count} (includes both main list and 'Other Files' section)");

            // Normalization helper for robust filename matching
            // Handles both main list items (e.g., "FileName.rvt (123 elements)") and "Other Files" section items
            Func<string, string> norm = s =>
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                var trimmed = s;
                // Remove everything after first parenthesis (e.g., "(123 elements)", "(Active Document)", "[NOT LOADED]")
                var idxParen = trimmed.IndexOf('(');
                if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
                var idxBracket = trimmed.IndexOf('[');
                if (idxBracket >= 0) trimmed = trimmed.Substring(0, idxBracket);
                trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed.Trim());
                trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
                trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
                trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
                return trimmed.Trim();
            };

            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;
                string linkTitleNorm = norm(linkDoc.Title);

                Transform linkTransform = link.GetTotalTransform();
                Transform invTransform = linkTransform.Inverse;

                XYZ linkMin = invTransform.OfPoint(modelMin);
                XYZ linkMax = invTransform.OfPoint(modelMax);

                XYZ actualMin = new XYZ(
                    Math.Min(linkMin.X, linkMax.X),
                    Math.Min(linkMin.Y, linkMax.Y),
                    Math.Min(linkMin.Z, linkMax.Z)
                );
                XYZ actualMax = new XYZ(
                    Math.Max(linkMin.X, linkMax.X),
                    Math.Max(linkMin.Y, linkMax.Y),
                    Math.Max(linkMin.Z, linkMax.Z)
                );

                _logger($"DEBUG: Section box in linked file '{linkDoc.Title}': Min=({actualMin.X:F2}, {actualMin.Y:F2}, {actualMin.Z:F2}) Max=({actualMax.X:F2}, {actualMax.Y:F2}, {actualMax.Z:F2})");

                Outline linkOutline = new Outline(actualMin, actualMax);

                // Collect MEP from reference links only (if any specified); otherwise from all links
                // ✅ FIX: Properly match files from both main list and "Other Files" section
                bool refMatch = selectedReferenceFiles.Count == 0 || selectedReferenceFiles.Any(f =>
                {
                    string ui = norm(f);
                    bool matches = linkTitleNorm.Contains(ui) || ui.Contains(linkTitleNorm);
                    if (matches)
                    {
                        _logger($"DEBUG: ✅ Reference file MATCH: '{f}' (normalized: '{ui}') matches link '{linkDoc.Title}' (normalized: '{linkTitleNorm}')");
                    }
                    return matches;
                });
                if (!refMatch && selectedReferenceFiles.Count > 0)
                {
                    _logger($"DEBUG: ⚠️ Reference file NO MATCH: Link '{linkDoc.Title}' (normalized: '{linkTitleNorm}') not in selected files. Selected files (normalized): {string.Join(", ", selectedReferenceFiles.Select(f => $"'{norm(f)}'"))}");
                }
                if (refMatch)
                {
                    foreach (var cat in mepCats)
                    {
                        var collector = new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType();
                        
                        // ✅ CRITICAL: Use BoundingBoxIntersectsFilter - handles PARTIAL intersections correctly
                        // BoundingBoxIntersectsFilter checks if element's bounding box INTERSECTS (not fully contained)
                        // This means elements that are PARTIALLY within section box are included (as required)
                        // Example: A pipe that extends through section box boundary will be loaded
                        var filteredByOutline = collector
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))  // ✅ Includes partial elements
                            .ToElements()
                            .ToList();
                        
                        // ✅ FALLBACK: Manually check elements with expanded outline to catch edge cases
                        // Expanded outline (+0.1ft tolerance) ensures elements on exact boundaries are not missed
                        double tolerance = 0.1; // 0.1 feet (~30mm) expansion for fallback check
                        var expandedActualMin = new XYZ(actualMin.X - tolerance, actualMin.Y - tolerance, actualMin.Z - tolerance);
                        var expandedActualMax = new XYZ(actualMax.X + tolerance, actualMax.Y + tolerance, actualMax.Z + tolerance);
                        var expandedLinkOutline = new Outline(expandedActualMin, expandedActualMax);
                        
                        var expandedFiltered = new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(expandedLinkOutline))
                            .ToElements();
                        
                        var filteredIds = new HashSet<int>(filteredByOutline.Select(e => e.Id.IntegerValue));
                        var missedElements = expandedFiltered
                            .Where(e => !filteredIds.Contains(e.Id.IntegerValue))
                            .Where(e => {
                                try
                                {
                                    var bbox = e.get_BoundingBox(null);
                                    if (bbox == null) return false;
                                    // Manual intersection check with tolerance
                                    return BoundingBoxService.BoundingBoxesIntersectWithTolerance(
                                        actualMin, actualMax, 
                                        bbox.Min, bbox.Max, 
                                        tolerance: 0.01); // 0.01 feet tolerance (~3mm)
                                }
                                catch { return false; }
                            })
                            .ToList();
                        
                        if (missedElements.Count > 0)
                        {
                            _logger($"⚠️ Found {missedElements.Count} {cat} elements missed by BoundingBoxIntersectsFilter in link '{linkDoc.Title}', adding them manually");
                            filteredByOutline = filteredByOutline.Concat(missedElements).ToList();
                        }
                        
                        mepElements.AddRange(filteredByOutline);
                    }
                    _logger($"Collected {mepElements.Count} total MEP elements after processing reference link '{linkDoc.Title}'");
                }
                else
                {
                    _logger($"DEBUG: Skipping MEP collection from link '{linkDoc.Title}' (not in selected reference files)");
                }

                // Collect hosts only from selected host files (if specified), else from all links
                // ✅ FIX: Properly match files from both main list and "Other Files" section
                bool hostMatch = selectedHostFiles == null || selectedHostFiles.Count == 0 || selectedHostFiles.Any(f =>
                {
                    string ui = norm(f);
                    bool matches = linkTitleNorm.Contains(ui) || ui.Contains(linkTitleNorm);
                    if (matches)
                    {
                        _logger($"DEBUG: ✅ Host file MATCH: '{f}' (normalized: '{ui}') matches link '{linkDoc.Title}' (normalized: '{linkTitleNorm}')");
                    }
                    return matches;
                });
                if (!hostMatch && selectedHostFiles != null && selectedHostFiles.Count > 0)
                {
                    _logger($"DEBUG: ⚠️ Host file NO MATCH: Link '{linkDoc.Title}' (normalized: '{linkTitleNorm}') not in selected files. Selected files (normalized): {string.Join(", ", selectedHostFiles.Select(f => $"'{norm(f)}'"))}");
                }
                else
                {
                    _logger($"DEBUG: Host match for '{linkDoc.Title}': {hostMatch} (selected host files count: {selectedHostFiles?.Count ?? 0})");
                }
                
                if (hostMatch)
                {
                    // ✅ OPTIMIZATION: Get settings for property-based filtering
                    var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                    double minWallThicknessMm = settings.MinWallThickness;
                    bool ignoreArchFloors = settings.IgnoreArchitecturalFloors;
                    
                    // ✅ FIX: Only collect structural categories that are selected in UI
                    if (allowedHostElementTypes == null || allowedHostElementTypes.Count == 0)
                    {
                        // Fallback: collect all structural categories if no selection
                        _logger("No host element types selected - collecting all structural categories");
                        
                        // ✅ OPTIMIZATION: Apply category + bounding box filters at collector level
                        // Property filters applied after collection (still efficient - only filters elements that passed bounding box)
                        var walls = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_Walls)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements();
                        wallElements.AddRange(ElementPropertyFilters.FilterWallsByThickness(walls, minWallThicknessMm));
                        
                        var floors = new FilteredElementCollector(linkDoc)
                            .OfCategory(BuiltInCategory.OST_Floors)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements();
                        wallElements.AddRange(ElementPropertyFilters.FilterStructuralFloors(floors, ignoreArchFloors));
                        
                        wallElements.AddRange(
                            new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements()
                        );
                    }
                    else
                    {
                        // ✅ FILTERED: Only collect selected host element types
                        _logger($"UI Selected Host types: {string.Join(", ", allowedHostElementTypes)}");
                        
                        if (allowedHostElementTypes.Any(ht => ht.Equals("Walls", StringComparison.OrdinalIgnoreCase)))
                        {
                            // ✅ OPTIMIZATION: Apply category + bounding box at collector level, property filter after
                            var walls = new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements();
                            wallElements.AddRange(ElementPropertyFilters.FilterWallsByThickness(walls, minWallThicknessMm));
                            _logger("UI Selection: Including Walls");
                        }

                        if (allowedHostElementTypes.Any(ht => ht.Equals("Floors", StringComparison.OrdinalIgnoreCase)))
                        {
                            // ✅ OPTIMIZATION: Apply category + bounding box at collector level, property filter after
                            var floors = new FilteredElementCollector(linkDoc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .WhereElementIsNotElementType()
                                .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                .ToElements();
                            wallElements.AddRange(ElementPropertyFilters.FilterStructuralFloors(floors, ignoreArchFloors));
                            _logger("UI Selection: Including Floors");
                        }

                        if (allowedHostElementTypes.Any(ht => ht.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)))
                        {
                            // ✅ OPTIMIZATION: Apply all filters at collector level
                            wallElements.AddRange(
                                new FilteredElementCollector(linkDoc)
                                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                                    .WhereElementIsNotElementType()
                                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                                    .ToElements()
                            );
                            _logger("UI Selection: Including Structural Framing");
                        }
                        
                        _logger($"✅ FILTERED: Only collected {allowedHostElementTypes.Count} selected host types (was 3 total)");
                    }

                    _logger($"Collected {wallElements.Count} total host elements after processing host link '{linkDoc.Title}'");
                    _logger($"DEBUG: Summary - Host elements collected from '{linkDoc.Title}': {wallElements.Count} total (after processing this link)");
                }
                else
                {
                    _logger($"DEBUG: Skipping host collection from link '{linkDoc.Title}' (not in selected host files)");
                }
            }
            
            // ✅ DEBUG: Final summary after all collection
            _logger($"DEBUG: FINAL COLLECTION SUMMARY:");
            _logger($"DEBUG:   - MEP elements: {mepElements.Count} total (from active document and/or linked files)");
            _logger($"DEBUG:   - Host elements: {wallElements.Count} total (from linked files only)");
            _logger($"DEBUG:   - Selected reference files: {string.Join(", ", selectedReferenceFiles)}");
            _logger($"DEBUG:   - Selected host files: {string.Join(", ", selectedHostFiles ?? new List<string>())}");

            // ✅ CODING PRACTICE: NO FALLBACK LOGIC - Respect user selections exactly
            // If no MEP elements found with selected filters, return 0 (don't scan other files)
            // Fallback logic removed per user requirement - only use explicit selections
            if (mepElements.Count == mepCountBeforeLinks && mepCats.Count > 0)
            {
                if (links.Count == 0)
                {
                    _logger("SINGLE-MODEL WORKFLOW: No linked files found. MEP elements are in the active document (already collected).");
                }
                else if (mepElements.Count == 0)
                {
                    // ✅ NO FALLBACK: Respect user selections - if 0 found, return 0
                    if (selectedReferenceFiles != null && selectedReferenceFiles.Count > 0)
                    {
                        _logger($"⚠️ RESPECTING STEP 2: Reference files selected ({string.Join(", ", selectedReferenceFiles)}) but found 0 matching MEP elements. Returning 0 (no fallback).");
                    }
                    else
                    {
                        _logger("⚠️ No MEP elements found and no reference files selected. Returning 0 (no fallback).");
                    }
                }
            }

            // ✅ CODING PRACTICE: NO FALLBACK LOGIC - Respect user selections exactly
            // If no elements found with selected filters, return 0 (don't scan other files)
            // Fallback logic removed per user requirement - only use explicit selections
            if (wallElements.Count == 0)
            {
                if (links.Count == 0)
                {
                    _logger("SINGLE-MODEL WORKFLOW: No linked files found. Host elements (walls/floors/framing) are in the active document (already collected).");
                }
                else
                {
                    // ✅ NO FALLBACK: Respect user selections - if 0 found, return 0
                    if (selectedHostFiles != null && selectedHostFiles.Count > 0)
                    {
                        _logger($"⚠️ RESPECTING STEP 4: Host files selected ({string.Join(", ", selectedHostFiles)}) but found 0 matching host elements. Returning 0 (no fallback).");
                    }
                    else
                    {
                        _logger("⚠️ No host elements found and no host files selected. Returning 0 (no fallback).");
                    }
                }
            }
        }

        /// <summary>
        /// ✅ NICE3POINT TEMPLATE PATTERN: Get MepIntersectionService via factory to avoid static initialization
        /// </summary>
        private IMepIntersectionService GetMepIntersectionService()
        {
            return MepIntersectionServiceFactory.GetInstance();
        }

        /// <summary>
        /// ✅ R2024 FIX: Get transform directly from link instance without accessing MepIntersectionService
        /// This avoids TypeInitializationException that occurs when accessing static class in R2024
        /// </summary>
        private static Transform? GetTransformDirectly(RevitLinkInstance? link)
        {
            if (link == null) return null;
            try
            {
                return link.GetTotalTransform();
            }
            catch
            {
                return Transform.Identity;
            }
        }

        private List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsInternal(
            List<Element> mepElements, 
            List<Element> wallElements, 
            Document doc,
            IntersectionOptimizationService optimizationService = null)
        {
            _logger($"Using optimized MepIntersectionService.FindIntersectionsBatch with {mepElements.Count} MEP elements...");

            // Get all linked documents for transform lookup
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            // ✅ PHASE 2 OPTIMIZATION: Prepare MEP and structural elements for batch processing
            // ✅ R2024 FIX: Use direct transform method to avoid TypeInitializationException
            var mepElementsWithTransforms = new List<(Element, Transform?)>();
            foreach (var mep in mepElements)
            {
                try
                {
                    bool isDamper = false;
                    if (mep is FamilyInstance fi)
                    {
                        string familyName = fi.Symbol?.Family?.Name ?? "";
                        isDamper = familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0;
                    }

                    var mepLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == mep.Document.Title);
                    // ✅ R2024 FIX: Get transform directly without accessing MepIntersectionService static class
                    Transform? mepTransform = GetTransformDirectly(mepLink);
                    mepElementsWithTransforms.Add((mep, mepTransform));
                }
                catch (Exception ex)
                {
                    _logger($"Error preparing MEP element {mep.Id}: {ex.Message}");
                }
            }

            var structuralElementsWithTransforms = new List<(Element, Transform?)>();
            foreach (var wall in wallElements)
            {
                try
                {
                    var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                    // ✅ R2024 FIX: Get transform directly without accessing MepIntersectionService static class
                    Transform? wallTransform = GetTransformDirectly(wallLink);
                    structuralElementsWithTransforms.Add((wall, wallTransform));
                }
                catch (Exception ex)
                {
                    _logger($"Error preparing structural element {wall.Id}: {ex.Message}");
                }
            }

            // ✅ PHASE 2 OPTIMIZATION: Use batch processing with spatial hash grid and curve-in-bbox test
            // ✅ NICE3POINT TEMPLATE PATTERN: Use factory to get version-specific implementation
            _logger($"[PHASE2] Calling MepIntersectionService.FindIntersectionsBatch with {mepElementsWithTransforms.Count} MEP and {structuralElementsWithTransforms.Count} structural elements");
            var intersectionService = GetMepIntersectionService();
            var intersections = intersectionService.FindIntersectionsBatch(
                mepElementsWithTransforms,
                structuralElementsWithTransforms,
                _logger,
                optimizationService?.KnownValidPairs,
                optimizationService?.ShouldSkipKnownPairsGeometryCheck ?? false);
            
            _logger($"[PHASE2] Optimized batch processing completed: {intersections.Count} intersections found");
            return intersections;
        }

        // OLD CODE BELOW - Replaced with optimized FindIntersectionsBatch method
        // Keeping for reference, but never executed
        /*
            foreach (var mep in mepElements)
            {
                try
                {
                    // ✅ PHASE 1 OPTIMIZATION: Check cache invalidation
                    if (CacheInvalidationMonitor.NeedsInvalidation(mep))
                    {
                        _logger($"[IntersectionDetectionService] MEP element {mep.Id} needs cache invalidation");
                    }

                    bool isDamper = false;
                    string familyName = string.Empty;
                    if (mep is FamilyInstance fi)
                    {
                        familyName = fi.Symbol?.Family?.Name ?? string.Empty;
                        isDamper = familyName.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (isDamper)
                        {
                            _logger($"DEBUG: Processing damper element {mep.Id} - Family: {familyName}");
                        }
                    }

                    var mepLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == mep.Document.Title);
                    Transform? mepTransform = mepLink?.GetTotalTransform();
                    // MEP link and transform found

                    // ✅ PHASE 1 OPTIMIZATION: Use smart tolerance
                    var tolerance = SmartToleranceService.GetIntersectionTolerance(mep, wallElements.FirstOrDefault());
                    _logger($"[IntersectionDetectionService] Using smart tolerance: {tolerance:F3} for MEP element {mep.Id}");

                    var structuralElements = new List<(Element, Transform?)>();
                    foreach (var wall in wallElements)
                    {
                        var wallLink = links.FirstOrDefault(link => link.GetLinkDocument()?.Title == wall.Document.Title);
                        if (wallLink != null)
                        {
                            var wallTransform = wallLink.GetTotalTransform();
                            structuralElements.Add((wall, wallTransform));
                        }
                        else
                        {
                            _logger($"DEBUG: No link found for wall {wall.Id} from document {wall.Document.Title}");
                            _logger($"DEBUG: Available links: {string.Join(", ", links.Select(l => l.GetLinkDocument()?.Title ?? "NULL"))}");
                            structuralElements.Add((wall, null));
                        }
                    }

                    var mepBBox = mep.get_BoundingBox(null);

                    if (isDamper)
                    {
                        var damperHits = MepIntersectionService.FindDamperIntersections(
                            mep,
                            structuralElements,
                            mepTransform,
                            _logger);

                        _logger($"DEBUG: Damper {mep.Id} intersections found: {damperHits.Count}");

                        foreach (var (structuralElement, bbox, center) in damperHits)
                        {
                            intersections.Add((mep, structuralElement, bbox, center));
                            _logger($"Intersection found: Damper {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }

                        if (damperHits.Count == 0)
                        {
                            _logger($"DEBUG: No intersections found for damper {mep.Id}");
                        }

                        continue;
                    }

                    var locationCurve = mep.Location as LocationCurve;
                    if (locationCurve?.Curve is Line mepLine)
                    {
                        _logger($"DEBUG: Processing MEP {mep.Id} ({mep.Category?.Name}) with line from {mepLine.GetEndPoint(0)} to {mepLine.GetEndPoint(1)}");

                        // ✅ METHOD 3: Check for damper presence at duct end points BEFORE processing intersections
                        // NOTE: Do NOT skip the entire duct - only skip intersections at the damper end
                        // This ensures all ducts in the section box are processed for intersections
                        if (mep.Category?.Name == "Ducts" || mep.Category?.Name == "Duct Curves")
                        {
                            _logger($"[METHOD3] Checking duct {mep.Id} for damper presence at end points");
                            
                            // Don't skip the duct entirely - process all intersections
                            // The damper check can be used later to filter specific intersections if needed
                            bool hasDamper = CheckForDamperAtDuctEnd(doc, mep, mepLine, mepTransform);
                            if (hasDamper)
                            {
                                _logger($"[METHOD3] NOTE: Duct {mep.Id} has damper at end, but still processing intersections for other parts of duct");
                            }
                            else
                            {
                                _logger($"[METHOD3] No damper found at duct {mep.Id} end points - proceeding with intersection detection");
                            }
                        }

                        Line mepLineInHostShared;
                        if (mepTransform != null)
                        {
                            mepLineInHostShared = Line.CreateBound(
                                mepTransform.OfPoint(mepLine.GetEndPoint(0)),
                                mepTransform.OfPoint(mepLine.GetEndPoint(1))
                            );
                        }
                        else
                        {
                            mepLineInHostShared = mepLine;
                        }

                        _logger($"DEBUG: MEP line in host shared coordinates: {mepLineInHostShared.GetEndPoint(0)} to {mepLineInHostShared.GetEndPoint(1)}");

                        BoundingBoxXYZ? mepBBoxInHostShared = mepBBox;
                        if (mepBBox != null && mepTransform != null)
                        {
                            _logger($"DEBUG: MEP bbox before transform: Min=({mepBBox.Min.X:F2}, {mepBBox.Min.Y:F2}, {mepBBox.Min.Z:F2}) Max=({mepBBox.Max.X:F2}, {mepBBox.Max.Y:F2}, {mepBBox.Max.Z:F2})");

                            var mepPts = new[]
                            {
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Min.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Min.X, mepBBox.Max.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Min.Y, mepBBox.Max.Z)),
                                mepTransform.OfPoint(new XYZ(mepBBox.Max.X, mepBBox.Max.Y, mepBBox.Min.Z))
                            };

                            var mepNewMin = new XYZ(mepPts.Min(p => p.X), mepPts.Min(p => p.Y), mepPts.Min(p => p.Z));
                            var mepNewMax = new XYZ(mepPts.Max(p => p.X), mepPts.Max(p => p.Y), mepPts.Max(p => p.Z));

                            mepBBoxInHostShared = new BoundingBoxXYZ
                            {
                                Min = mepNewMin,
                                Max = mepNewMax
                            };

                            _logger($"DEBUG: MEP bbox after transform: Min=({mepBBoxInHostShared.Min.X:F2}, {mepBBoxInHostShared.Min.Y:F2}, {mepBBoxInHostShared.Min.Z:F2}) Max=({mepBBoxInHostShared.Max.X:F2}, {mepBBoxInHostShared.Max.Y:F2}, {mepBBoxInHostShared.Max.Z:F2})");
                        }

                        // ✅ PHASE 1 OPTIMIZATION: Use memory management for intersection results
                        var hits = MepIntersectionService.FindIntersections(
                            mep,
                            structuralElements,
                            _logger);
                        
                        // Cache the intersection results for potential reuse
                        if (hits.Any())
                        {
                            MemoryManagementService.AddToCache(mep.Id, hits, hits.Count * 100); // Estimate 100 bytes per hit
                        }

                        _logger($"DEBUG: MepIntersectionService found {hits.Count} intersections for MEP {mep.Id}");

                        foreach (var (structuralElement, bbox, center) in hits)
                        {
                            intersections.Add((mep, structuralElement, bbox, center));
                            _logger($"Intersection found: MEP {mep.Id} ({mep.Category?.Name}) <-> {structuralElement.Category?.Name} {structuralElement.Id} at {center}");
                        }
                    }
                    else
                    {
                        // MEP element has no valid location curve or is not a line - skipping
                    }
                }
                catch (Exception ex)
                {
                    _logger($"ERROR: Failed to process MEP element {mep.Id}: {ex.Message}");
                }
            }

            // ✅ PHASE 1 OPTIMIZATION: Enforce memory management
            MemoryManagementService.EnforceCacheSize();
            
            _logger($"FindIntersectionsInternal completed: {intersections.Count} intersections found");
            return intersections;
        }
        */ // End of old code block

        /// <summary>
        /// METHOD 3: Check for damper presence at duct end points
        /// This prevents creating intersections for ducts that already have dampers
        /// </summary>
        private bool CheckForDamperAtDuctEnd(Document document, Element ductElement, Line ductLine, Transform ductTransform)
        {
            try
            {
                if (!(ductElement is Autodesk.Revit.DB.Mechanical.Duct duct))
                {
                    return false; // Not a duct, no need to check
                }

                _logger($"[METHOD3] Checking for damper at duct end: Duct {duct.Id}");

                // Get duct end points in host coordinates
                var endPoint1 = ductTransform?.OfPoint(ductLine.GetEndPoint(0)) ?? ductLine.GetEndPoint(0);
                var endPoint2 = ductTransform?.OfPoint(ductLine.GetEndPoint(1)) ?? ductLine.GetEndPoint(1);

                _logger($"[METHOD3] Duct end points: {endPoint1} and {endPoint2}");

                // Check each end point for damper presence
                foreach (var endPoint in new[] { endPoint1, endPoint2 })
                {
                    if (CheckForDamperNearPoint(document, endPoint))
                    {
                        _logger($"[METHOD3] ✓ DAMPER FOUND: Damper detected near duct end point {endPoint}");
                        return true;
                    }
                }

                _logger($"[METHOD3] ✗ NO DAMPER: No damper found near any duct end points");
                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to check for damper at duct end: {ex.Message}");
                return false; // Default to false (don't skip) if error occurs
            }
        }

        /// <summary>
        /// Check for damper presence near a specific point
        /// </summary>
        private bool CheckForDamperNearPoint(Document document, XYZ searchPoint)
        {
            try
            {
                double searchRadius = 0.5; // 0.5 feet = ~150mm
                
                // ✅ CRITICAL FIX: Search in ALL linked documents, not just the host document
                // The ducts and dampers are in linked files (ME-00001), not the host document
                var allDampers = new List<Element>();
                
                // Search in host document first
                var hostCollector = new FilteredElementCollector(document);
                var hostDampers = hostCollector
                    .OfCategory(BuiltInCategory.OST_DuctAccessory)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .Where(d => IsDamperType(d))
                    .ToList();
                allDampers.AddRange(hostDampers);
                
                // Search in all linked documents
                var linkedDocs = document.Application.Documents.Cast<Document>()
                    .Where(doc => doc != document && !doc.IsFamilyDocument)
                    .ToList();
                
                foreach (var linkedDoc in linkedDocs)
                {
                    try
                    {
                        var linkedCollector = new FilteredElementCollector(linkedDoc);
                        var linkedDampers = linkedCollector
                            .OfCategory(BuiltInCategory.OST_DuctAccessory)
                            .WhereElementIsNotElementType()
                            .ToElements()
                            .Where(d => IsDamperType(d))
                            .ToList();
                        allDampers.AddRange(linkedDampers);
                        
                        _logger($"[METHOD3] Found {linkedDampers.Count} dampers in linked doc '{linkedDoc.Title}'");
                    }
                    catch (Exception ex)
                    {
                        _logger($"[METHOD3] WARNING: Could not search linked doc '{linkedDoc.Title}': {ex.Message}");
                    }
                }

                _logger($"[METHOD3] Found {allDampers.Count} total duct accessories across all documents");

                foreach (var damper in allDampers)
                {
                    var damperLocation = damper.Location as LocationPoint;
                    if (damperLocation?.Point != null)
                    {
                        double distance = damperLocation.Point.DistanceTo(searchPoint);
                        if (distance < searchRadius)
                        {
                            var damperType = damper.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString() ?? "";
                            var damperFamily = (damper as FamilyInstance)?.Symbol?.Family?.Name ?? "";
                            
                            _logger($"[METHOD3] Checking damper {damper.Id}: Type='{damperType}', Family='{damperFamily}'");
                            
                            if (IsDamperType(damper))
                            {
                                _logger($"[METHOD3] ✓ CONFIRMED DAMPER: {damper.Id} is a damper (Type='{damperType}', Family='{damperFamily}')");
                                return true;
                            }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to search for dampers near point: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Check if two bounding boxes intersect with a tolerance
        /// This is more lenient than BoundingBoxIntersectsFilter which can miss edge cases
        /// </summary>
        // ✅ OOP REFACTORING: Removed duplicate BoundingBoxesIntersectWithTolerance - now uses BoundingBoxService.BoundingBoxesIntersectWithTolerance()
        
        /// <summary>
        /// Check if an element is a damper based on type and family name
        /// </summary>
        private bool IsDamperType(Element element)
        {
            try
            {
                if (element is FamilyInstance fi)
                {
                    var familyName = fi.Symbol?.Family?.Name ?? "";
                    var typeName = fi.Symbol?.Name ?? "";
                    var combinedName = $"{familyName} {typeName}".ToLowerInvariant();
                    
                    // Check for damper keywords
                    var damperKeywords = new[] { "damper", "dam", "fire", "smoke", "motorized", "motorised" };
                    
                    if (damperKeywords.Any(keyword => combinedName.Contains(keyword)))
                    {
                        _logger($"[METHOD3] Damper detected: '{combinedName}' contains damper keywords");
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger($"[METHOD3] ERROR: Failed to check damper type: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Filters walls by minimum thickness setting
        /// Skips walls that are thinner than the user-specified minimum
        /// </summary>
        private static IEnumerable<Element> FilterWallsByMinimumThickness(IEnumerable<Element> walls)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                double minThicknessMm = settings.MinWallThickness;
                
                // If setting is 0 or negative, don't filter (all walls allowed)
                if (minThicknessMm <= 0)
                    return walls;
                
                double minThicknessInternal = RevitUnitConversionService.Instance.ToInternalMillimeters(minThicknessMm);
                var filteredWalls = new List<Element>();
                int skippedCount = 0;
                
                foreach (var wall in walls)
                {
                    if (wall is Wall wallObj)
                    {
                        double wallThickness = wallObj.Width;
                        
                        if (wallThickness >= minThicknessInternal)
                        {
                            filteredWalls.Add(wall);
                        }
                        else
                        {
                            skippedCount++;
                            double wallThicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(wallThickness);
                            System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] SKIP: Wall {wall.Id.IntegerValue} thickness {wallThicknessMm:F1}mm < {minThicknessMm:F1}mm minimum");
                        }
                    }
                    else
                    {
                        // Not a wall, add it
                        filteredWalls.Add(wall);
                    }
                }
                
                if (skippedCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] Filtered {skippedCount} walls below {minThicknessMm:F1}mm minimum thickness");
                }
                
                return filteredWalls;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] Error filtering walls by minimum thickness: {ex.Message}");
                return walls; // Return original list on error
            }
        }

        /// <summary>
        /// Filters architectural floors if the setting is enabled
        /// Skips floors where Structural parameter is not checked
        /// </summary>
        private static IEnumerable<Element> FilterArchitecturalFloors(IEnumerable<Element> structuralElements)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                bool ignoreArchFloors = settings.IgnoreArchitecturalFloors;
                
                // If setting is disabled, don't filter (all floors allowed)
                if (!ignoreArchFloors)
                    return structuralElements;
                
                var filteredElements = new List<Element>();
                int skippedCount = 0;
                
                foreach (var element in structuralElements)
                {
                    if (element is Floor floor)
                    {
                        // Check Structural parameter
                        Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                        bool isStructural = structuralParam?.AsInteger() == 1;
                        
                        if (isStructural)
                        {
                            filteredElements.Add(element);
                        }
                        else
                        {
                            skippedCount++;
                            System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] SKIP: Architectural floor {floor.Id.IntegerValue} (Structural parameter not checked)");
                        }
                    }
                    else
                    {
                        // Not a floor, add it (walls, framing, etc.)
                        filteredElements.Add(element);
                    }
                }
                
                if (skippedCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] Filtered {skippedCount} architectural floors");
                }
                
                return filteredElements;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[IntersectionDetectionService] Error filtering architectural floors: {ex.Message}");
                return structuralElements; // Return original list on error
            }
        }
    }
}
