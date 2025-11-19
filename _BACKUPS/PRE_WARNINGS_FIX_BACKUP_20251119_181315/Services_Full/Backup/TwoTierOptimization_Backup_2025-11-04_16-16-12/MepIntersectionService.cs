#nullable enable
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class MepIntersectionService
    {
        // ✅ MEMORY OPTIMIZATION: LRU cache with max size to prevent unbounded growth
        // Large linked files can have thousands of structural elements - cache can grow to 100MB+ without limits
        private const int MAX_GEOMETRY_CACHE_SIZE = 5000; // Limit to 5000 entries (~10-20MB typical)
        private static readonly Dictionary<string, Solid?> _geometryCache = new Dictionary<string, Solid?>();
        private static readonly LinkedList<string> _geometryCacheOrder = new LinkedList<string>(); // LRU tracking
        
        // PHASE 2 OPTIMIZATION 1: Spatial partitioning service
        private static readonly SpatialPartitioningService _spatialService = new SpatialPartitioningService(1.0); // 1ft grid
        
        // PHASE 2 OPTIMIZATION 3: Transform cache (1.5x speedup)
        private static readonly Dictionary<Document, Transform> _transformCache = new Dictionary<Document, Transform>();
        
        // PHASE 1 OPTIMIZATION 2: Category Whitelist (2x speedup)
        private static readonly BuiltInCategory[] MEP_CATEGORY_WHITELIST = {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctAccessory,  // Includes dampers
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ConduitFitting
        };
        
        private static readonly BuiltInCategory[] STRUCTURAL_CATEGORY_WHITELIST = {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFoundation
        };
        
        // ✅ MEMORY OPTIMIZATION: LRU cache add with eviction
        private static void AddToGeometryCache(string key, Solid? solid)
        {
            // Remove if exists (move to end = most recently used)
            if (_geometryCache.ContainsKey(key))
            {
                _geometryCacheOrder.Remove(key);
                _geometryCache.Remove(key);
            }
            
            // Add to cache
            _geometryCache[key] = solid;
            _geometryCacheOrder.AddLast(key);
            
            // ✅ MEMORY OPTIMIZATION: Evict oldest entries if over limit
            while (_geometryCache.Count > MAX_GEOMETRY_CACHE_SIZE && _geometryCacheOrder.First != null)
            {
                var oldestKey = _geometryCacheOrder.First.Value;
                _geometryCache.Remove(oldestKey);
                _geometryCacheOrder.RemoveFirst();
            }
        }
        
        // ✅ MEMORY OPTIMIZATION: Get from cache and update LRU order
        private static bool TryGetFromGeometryCache(string key, out Solid? solid)
        {
            if (_geometryCache.TryGetValue(key, out solid))
            {
                // Move to end (most recently used)
                _geometryCacheOrder.Remove(key);
                _geometryCacheOrder.AddLast(key);
                return true;
            }
            return false;
        }
        
        // Clear cache method for memory management
        public static void ClearGeometryCache()
        {
            _geometryCache.Clear();
            _geometryCacheOrder.Clear();
        }
        
        // ✅ MEMORY OPTIMIZATION: Get cache statistics for monitoring
        public static (int count, int maxSize, double memoryEstimateMB) GetGeometryCacheStats()
        {
            // Rough estimate: 2-5 MB per 1000 entries (depends on solid complexity)
            double memoryEstimateMB = _geometryCache.Count * 0.003; // 3KB per entry average
            return (_geometryCache.Count, MAX_GEOMETRY_CACHE_SIZE, memoryEstimateMB);
        }
        
        /// <summary>
        /// PHASE 2 OPTIMIZATION 3: Get cached transform for a document
        /// </summary>
        public static Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null)
        {
            if (!_transformCache.ContainsKey(doc))
            {
                var link = links.FirstOrDefault(l => l.GetLinkDocument()?.Title == doc.Title);
                if (link != null)
                {
                    var transform = link.GetTotalTransform();
                    _transformCache[doc] = transform;
                    log?.Invoke($"[TransformCache] Cached transform for document: {doc.Title}");
                }
                else
                {
                    _transformCache[doc] = Transform.Identity;
                    log?.Invoke($"[TransformCache] No link found for document: {doc.Title}, using Identity transform");
                }
            }
            return _transformCache[doc];
        }
        
        /// <summary>
        /// PHASE 2 OPTIMIZATION 3: Clear transform cache
        /// </summary>
        public static void ClearTransformCache()
        {
            _transformCache.Clear();
        }
        
        // PHASE 1 OPTIMIZATION 2: Category whitelist filtering methods
        public static bool IsMepCategoryWhitelisted(Element element)
        {
            if (element.Category?.Id?.IntegerValue == null) return false;
            
            var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
            return MEP_CATEGORY_WHITELIST.Contains(categoryId);
        }
        
        public static bool IsStructuralCategoryWhitelisted(Element element)
        {
            if (element.Category?.Id?.IntegerValue == null) return false;
            
            var categoryId = (BuiltInCategory)element.Category.Id.IntegerValue;
            return STRUCTURAL_CATEGORY_WHITELIST.Contains(categoryId);
        }
        
        public static bool IsDamperElement(Element element)
        {
            if (element is FamilyInstance fi)
            {
                var familyName = fi.Symbol?.Family?.Name?.ToLower() ?? "";
                return familyName.Contains("damper");
            }
            return false;
        }
        
        // BATCH PROCESSING: Find intersections for multiple MEP elements efficiently
        /// <param name="knownValidPairs">Set of (MEP Element ID, Structural Element ID) pairs that have valid clash zones. When skipKnownPairsGeometryCheck is true, expensive geometry intersection is skipped for these pairs.</param>
        /// <param name="skipKnownPairsGeometryCheck">If true, skip expensive geometry intersection calculation for known valid pairs (just verify element existence with fast bounding box check).</param>
        // ✅ MEMORY OPTIMIZATION: Chunk size for processing large batches
        private const int MEP_CHUNK_SIZE = 500; // Process 500 MEP elements at a time to prevent memory buildup
        
        public static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            HashSet<(int mepId, int structuralId)> knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false)
        {
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            if (OptimizationFlags.UseDiagnosticMode)
                log($"[BatchIntersection] Processing {mepElements.Count} MEP elements against {structuralElements.Count} structural elements");
            
            // ✅ MEMORY OPTIMIZATION: Process in chunks for large batches
            // This prevents memory buildup when processing thousands of MEP elements
            if (mepElements.Count <= MEP_CHUNK_SIZE)
            {
                // Small batch - process all at once
                return FindIntersectionsBatchInternal(mepElements, structuralElements, log, knownValidPairs, skipKnownPairsGeometryCheck);
            }
            
            // Large batch - process in chunks
            log($"[Memory] Processing {mepElements.Count} MEP elements in chunks of {MEP_CHUNK_SIZE} to reduce memory usage");
            int totalChunks = (mepElements.Count + MEP_CHUNK_SIZE - 1) / MEP_CHUNK_SIZE;
            
            for (int chunkIndex = 0; chunkIndex < totalChunks; chunkIndex++)
            {
                int startIndex = chunkIndex * MEP_CHUNK_SIZE;
                int endIndex = Math.Min(startIndex + MEP_CHUNK_SIZE, mepElements.Count);
                var chunk = mepElements.GetRange(startIndex, endIndex - startIndex);
                
                log($"[Memory] Processing chunk {chunkIndex + 1}/{totalChunks} ({chunk.Count} MEP elements)");
                
                var chunkResults = FindIntersectionsBatchInternal(chunk, structuralElements, log, knownValidPairs, skipKnownPairsGeometryCheck);
                results.AddRange(chunkResults);
                
                // ✅ MEMORY OPTIMIZATION: Clear cache periodically if getting large
                if ((chunkIndex + 1) % 5 == 0) // Every 5 chunks
                {
                    var (cacheCount, _, cacheMB) = GetGeometryCacheStats();
                    if (cacheMB > 15.0) // If cache > 15MB, clear half
                    {
                        log($"[Memory] Clearing half of geometry cache ({cacheCount} entries, ~{cacheMB:F1} MB) to free memory");
                        // Clear oldest 50% of cache
                        int entriesToRemove = cacheCount / 2;
                        for (int i = 0; i < entriesToRemove && _geometryCacheOrder.First != null; i++)
                        {
                            var oldestKey = _geometryCacheOrder.First.Value;
                            _geometryCache.Remove(oldestKey);
                            _geometryCacheOrder.RemoveFirst();
                        }
                        GC.Collect(2, GCCollectionMode.Optimized); // Force GC after cache cleanup
                    }
                }
            }
            
            return results;
        }
        
        // Internal method for actual batch processing
        private static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatchInternal(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            HashSet<(int mepId, int structuralId)> knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false)
        {

            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            // ✅ MEMORY OPTIMIZATION: Only pre-compute bounding boxes, NOT solids (lazy loading)
            // Solids are expensive (2-5KB each) and many won't be needed after spatial filtering
            // Delay solid creation until after bounding box/curve checks pass
            var structuralData = new List<(Element element, Transform? transform, BoundingBoxXYZ bbox, string cacheKey)>();
            
            foreach (var (structElement, structTransform) in structuralElements)
            {
                var structBBox = structElement.get_BoundingBox(null);
                if (structBBox == null) continue;

                // Transform structural bbox to host shared coordinates
                if (structTransform != null)
                {
                    var transformedMin = structTransform.OfPoint(structBBox.Min);
                    var transformedMax = structTransform.OfPoint(structBBox.Max);
                    structBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }

                // ✅ MEMORY OPTIMIZATION: Store cache key only, compute solid lazily when needed
                string cacheKey = $"{structElement.Id.IntegerValue}_{structTransform?.GetHashCode() ?? 0}";
                structuralData.Add((structElement, structTransform, structBBox, cacheKey));
            }
            
            // ✅ MEMORY OPTIMIZATION: Log cache stats before processing
            var (cacheCount, cacheMax, cacheMB) = GetGeometryCacheStats();
            if (OptimizationFlags.UseDiagnosticMode)
            {
                log($"[Memory] Geometry cache: {cacheCount}/{cacheMax} entries (~{cacheMB:F1} MB)");
                log($"[BatchIntersection] Pre-computed {structuralData.Count} structural element bounding boxes (solids loaded lazily)");
            }
            
            // ✅ MEMORY OPTIMIZATION: Warn if cache is getting large
            if (cacheMB > 10.0)
            {
                log($"[Memory] ⚠️ WARNING: Geometry cache is large ({cacheMB:F1} MB). Consider clearing cache.");
            }

            // PHASE 2 OPTIMIZATION 1: Build spatial hash grid
            // Convert structuralData to format expected by SpatialPartitioningService (pass null for solids)
            var structuralDataForSpatial = structuralData.Select(sd => (sd.element, sd.transform, sd.bbox, (Solid?)null)).ToList();
            _spatialService.BuildGrid(structuralDataForSpatial);
            var (totalCells, usedCells, avgElements) = _spatialService.GetStatistics();
            if (OptimizationFlags.UseDiagnosticMode)
                log($"[SpatialHash] Built spatial grid: {usedCells} used cells, avg {avgElements:F1} elements per cell");

            // Process each MEP element against pre-computed structural data
            foreach (var (mepElement, mepTransform) in mepElements)
            {
                var mepBBox = mepElement.get_BoundingBox(null);
                if (mepBBox == null) continue;

                // Transform MEP bbox to host shared coordinates
                if (mepTransform != null)
                {
                    var transformedMin = mepTransform.OfPoint(mepBBox.Min);
                    var transformedMax = mepTransform.OfPoint(mepBBox.Max);
                    mepBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }

                var line = GetElementLine(mepElement, mepBBox, log);
                if (line == null)
                {
                    // Handle damper-style elements
                    var damperResults = FindDamperIntersectionsInternal(mepElement, mepBBox, 
                        structuralData.Select(sd => (sd.element, sd.transform)).ToList(), null, log);
                    results.AddRange(damperResults.Select(i => (mepElement, i.Item1, i.Item2, i.Item3)));
                    continue;
                }

                // Quick spatial pre-filtering with tolerance
                // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
                const double tolerance = 1.0; // 1.0ft tolerance - original working value
                var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
                var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
                var expandedBBox = new BoundingBoxXYZ { Min = expandedMin, Max = expandedMax };

                // ✅ TWO-TIER SPATIAL INDEX: TIER 1 - Spatial hash grid (fast rejection)
                var nearbyElements = _spatialService.GetNearbyElements(expandedBBox);
                if (OptimizationFlags.UseDiagnosticMode)
                    log($"[TwoTier] TIER 1 (SpatialGrid): MEP {mepElement.Id}: {nearbyElements.Count}/{structuralData.Count} nearby elements ({100.0 * nearbyElements.Count / structuralData.Count:F1}%)");

                // ✅ TWO-TIER SPATIAL INDEX: TIER 2 - R-tree precise filtering (if enabled)
                List<(Element element, Transform? transform, BoundingBoxXYZ bbox)> preciseCandidates;
                int rtreeFiltered = 0;
                
                if (OptimizationFlags.UseRTreeFilter && nearbyElements.Count > 0)
                {
                    // Create Outline for R-tree filter
                    var mepOutline = new Outline(expandedMin, expandedMax);
                    
                    // Group nearby elements by document for efficient R-tree filtering
                    var elementsByDocument = nearbyElements
                        .GroupBy(ne => ne.element.Document)
                        .ToList();
                    
                    var rtreeFilteredElements = new List<(Element element, Transform? transform, BoundingBoxXYZ bbox)>();
                    
                    foreach (var docGroup in elementsByDocument)
                    {
                        var doc = docGroup.Key;
                        var docElements = docGroup.ToList();
                        
                        // Use Revit's built-in R-tree filter (BoundingBoxIntersectsFilter)
                        try
                        {
                            var rtreeFilter = new BoundingBoxIntersectsFilter(mepOutline);
                            var filteredIds = new FilteredElementCollector(doc)
                                .WherePasses(rtreeFilter)
                                .WhereElementIsNotElementType()
                                .ToElementIds()
                                .ToHashSet();
                            
                            // Only keep elements that passed R-tree filter
                            foreach (var (element, transform, bbox) in docElements)
                            {
                                if (filteredIds.Contains(element.Id))
                                {
                                    rtreeFilteredElements.Add((element, transform, bbox));
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Fallback: If R-tree filter fails, use all elements from spatial grid
                            if (OptimizationFlags.UseDiagnosticMode)
                                log($"[TwoTier] R-tree filter failed for document {doc.Title}: {ex.Message}, falling back to spatial grid results");
                            rtreeFilteredElements.AddRange(docElements);
                        }
                    }
                    
                    preciseCandidates = rtreeFilteredElements;
                    rtreeFiltered = nearbyElements.Count - preciseCandidates.Count;
                    
                    if (OptimizationFlags.UseDiagnosticMode)
                        log($"[TwoTier] TIER 2 (R-tree): MEP {mepElement.Id}: {preciseCandidates.Count}/{nearbyElements.Count} precise candidates after R-tree filtering (rejected {rtreeFiltered})");
                }
                else
                {
                    // R-tree filtering disabled or no nearby elements, use spatial grid results directly
                    preciseCandidates = nearbyElements;
                    if (OptimizationFlags.UseDiagnosticMode && !OptimizationFlags.UseRTreeFilter)
                        log($"[TwoTier] R-tree filtering disabled, using spatial grid results directly");
                }

                int spatiallyFiltered = 0;
                int geometrySkippedForKnownPairs = 0;
                foreach (var (structElement, structTransform, structBBox) in preciseCandidates)
                {
                    // ✅ OOP OPTIMIZATION: Skip expensive geometry intersection for known valid pairs
                    // When 3-point validation is disabled, user trusts model unchanged
                    // Just verify bounding boxes intersect (fast check) instead of full geometry intersection
                    bool isKnownValidPair = skipKnownPairsGeometryCheck && 
                                           knownValidPairs != null && 
                                           knownValidPairs.Contains((mepElement.Id.IntegerValue, structElement.Id.IntegerValue));

                    if (isKnownValidPair)
                    {
                        // ✅ FAST PATH: Known valid pair - just verify bounding boxes intersect
                        // Skip expensive solid geometry intersection calculation
                        
                        // ✅ CRITICAL FIX: Validate structBBox is not null before using it
                        if (structBBox == null)
                        {
                            log?.Invoke($"[⚠️ SKIP-NULL-BBOX-KNOWN] Skipping known pair with null bounding box: MEP={mepElement.Id}, Structural={structElement.Id}");
                            spatiallyFiltered++;
                            continue;
                        }
                        
                        if (BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            // Use structural element's bounding box center as intersection point (approximation for known pairs)
                            var center = BoundingBoxService.GetBoundingBoxCenter(structBBox);
                            
                            // ✅ CRITICAL FIX: Validate intersection point is NOT zero before adding
                            if (center != null && Math.Abs(center.X) > 1e-9 && Math.Abs(center.Y) > 1e-9 && Math.Abs(center.Z) > 1e-9)
                            {
                                results.Add((mepElement, structElement, structBBox, center));
                                geometrySkippedForKnownPairs++;
                            }
                            else
                            {
                                log?.Invoke($"[⚠️ SKIP-ZERO-KNOWN] Skipping known pair with zero center: MEP={mepElement.Id}, Structural={structElement.Id}, Center={center}, BBox=Min({structBBox.Min.X},{structBBox.Min.Y},{structBBox.Min.Z}) Max({structBBox.Max.X},{structBBox.Max.Y},{structBBox.Max.Z})");
                            }
                        }
                        else
                        {
                            // Bounding boxes don't intersect - element may have moved (shouldn't happen if user trusts model)
                            // Still skip geometry check but log warning
                            if (OptimizationFlags.UseDiagnosticMode)
                                log($"[OPTIMIZATION] Known pair (MEP {mepElement.Id}, Structural {structElement.Id}) bounding boxes don't intersect - skipping");
                            spatiallyFiltered++;
                        }
                        continue;
                    }

                    // ✅ CRITICAL FIX: Validate structBBox is not null before using it in normal path
                    if (structBBox == null)
                    {
                        log?.Invoke($"[⚠️ SKIP-NULL-BBOX] Skipping intersection with null bounding box: MEP={mepElement.Id}, Structural={structElement.Id}");
                        spatiallyFiltered++;
                        continue;
                    }

                    // ✅ UNKNOWN/NEW PAIR: Run full geometry intersection (normal path)
                    // Quick bounding box intersection test
                    if (!BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                    {
                        spatiallyFiltered++;
                        continue;
                    }

                    // PHASE 2 OPTIMIZATION 2: Fast curve-in-bbox test
                    if (!TestCurveInBoundingBox(line, structBBox, tolerance))
                    {
                        spatiallyFiltered++;
                        continue; // Skip expensive solid intersection
                    }

                    // ✅ MEMORY OPTIMIZATION: Lazy solid loading - only compute after all cheap checks pass
                    // Get cache key from structural data
                    var structData = structuralData.First(sd => sd.element.Id == structElement.Id);
                    string cacheKey = structData.cacheKey;
                    
                    // Try to get from cache first
                    Solid? solid = null;
                    if (!TryGetFromGeometryCache(cacheKey, out solid))
                    {
                        // Cache miss - compute solid now (only for elements that passed all filters)
                        var options = new Options();
                        var geometry = structElement.get_Geometry(options);
                        if (geometry != null)
                        {
                            solid = GetSolidFromGeometry(geometry);
                            if (solid != null && structTransform != null)
                            {
                                solid = SolidUtils.CreateTransformed(solid, structTransform);
                            }
                        }
                        // Add to LRU cache (will evict oldest if over limit)
                        AddToGeometryCache(cacheKey, solid);
                    }
                    
                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, line, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        
                        // ✅ CRITICAL FIX: Validate bounding box is not null before using it
                        if (bbox == null)
                        {
                            log?.Invoke($"[⚠️ SKIP-NULL-BBOX] Skipping intersection with null bounding box: MEP={mepElement.Id}, Structural={structElement.Id}, IntersectionPoints={intersectionPoints.Count}");
                            continue;
                        }
                        
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
                        var center = BoundingBoxService.GetBoundingBoxCenter(bbox);
                        
                        // ✅ CRITICAL FIX: Validate intersection point is NOT zero before adding
                        // This prevents creating clash zones with (0,0,0) intersection points
                        if (center != null && Math.Abs(center.X) > 1e-9 && Math.Abs(center.Y) > 1e-9 && Math.Abs(center.Z) > 1e-9)
                        {
                            results.Add((mepElement, structElement, bbox, center));
                        }
                        else
                        {
                            // Log warning but don't add invalid intersection
                            log?.Invoke($"[⚠️ SKIP-ZERO] Skipping intersection with zero center point: MEP={mepElement.Id}, Structural={structElement.Id}, Center={center}");
                        }
                    }
                }
                
                if (geometrySkippedForKnownPairs > 0 && OptimizationFlags.UseDiagnosticMode)
                {
                    log($"[OPTIMIZATION] MEP {mepElement.Id}: Skipped geometry checks for {geometrySkippedForKnownPairs} known valid pairs");
                }

                if (spatiallyFiltered > 0)
                {
                    if (OptimizationFlags.UseDiagnosticMode)
                        log($"[BatchIntersection] MEP {mepElement.Id}: spatially filtered {spatiallyFiltered}/{preciseCandidates.Count} structural elements after two-tier filtering");
                }
                
                // ✅ TWO-TIER SUMMARY: Log filtering effectiveness
                if (OptimizationFlags.UseDiagnosticMode && OptimizationFlags.UseRTreeFilter && rtreeFiltered > 0)
                {
                    var totalFiltered = structuralData.Count - preciseCandidates.Count;
                    var tier1Rejected = structuralData.Count - nearbyElements.Count;
                    var tier2Rejected = rtreeFiltered;
                    log($"[TwoTier] Summary MEP {mepElement.Id}: Tier1 rejected {tier1Rejected}, Tier2 rejected {tier2Rejected}, Total rejected {totalFiltered}/{structuralData.Count} ({100.0 * totalFiltered / structuralData.Count:F1}%)");
                }
            }

            if (OptimizationFlags.UseDiagnosticMode)
                log($"[BatchIntersection] Found {results.Count} total intersections");
            return results;
        }

        // Individual method for backwards compatibility - now delegates to batch processing
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            // Delegate to batch processing with single element
            var batchResults = FindIntersectionsBatch(
                new List<(Element, Transform?)> { (mepElement, null) },
                structuralElements,
                log);
            
            // Convert batch results to individual format
            return batchResults.Select(r => (r.Item2, r.Item3, r.Item4)).ToList();
        }

        // Legacy method - kept for compatibility
        private static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersectionsLegacy(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            var mepBBox = mepElement.get_BoundingBox(null);
            if (mepBBox == null)
            {
                log($"WARNING: Could not get bounding box for MEP element {mepElement.Id}.");
                return results;
            }

            var line = GetElementLine(mepElement, mepBBox, log);
            if (line == null)
            {
                log($"INFO: Falling back to damper-style processing for element {mepElement.Id}.");
                results.AddRange(FindDamperIntersectionsInternal(mepElement, mepBBox, structuralElements, null, log));
                return results;
            }

            // REVERTED: Back to original 1.0 foot tolerance since coordinate transform issue is fixed
            // Both MEP and wall bounding boxes are now in host shared coordinates
            // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
            const double tolerance = 1.0; // 1.0ft tolerance - original working value
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

            int processedCount = 0;
            int spatiallyFilteredCount = 0;
            int lastLoggedSkipCount = 0;

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;
                processedCount++;
                
                try
                {
                    // SPATIAL PRE-FILTERING: Check bounding box intersection first
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox != null)
                    {
                        // CRITICAL FIX: Transform structural bbox to host shared coordinates for distance check
                        if (linkTransform != null)
                        {
                            var transformedMin = linkTransform.OfPoint(structBBox.Min);
                            var transformedMax = linkTransform.OfPoint(structBBox.Max);
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                                Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                            };
                        }
                        
                        // Quick bounding box intersection test (both in host shared coordinates now)
                        if (!BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            if (spatiallyFilteredCount - lastLoggedSkipCount >= 100)
                            {
                                lastLoggedSkipCount = spatiallyFilteredCount;
                                var wallType = structuralElement.GetType().Name;
                                var wallId = structuralElement.Id.IntegerValue;
                                var distance = GetDistanceToMepElement(mepBBox, structBBox, null);
                                log($"[MepIntersectionService] SPATIAL FILTER: Skipped {spatiallyFilteredCount} elements so far. Latest: {wallType} ID:{wallId} - Distance: {distance:F2}ft");
                            }
                            continue; // Skip expensive geometry processing
                        }
                    }

                    // ✅ MEMORY OPTIMIZATION: Get geometry from LRU cache or compute it
                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    Solid? solid;
                    
                    if (!TryGetFromGeometryCache(cacheKey, out solid))
                    {
                        var options = new Options();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;

                        solid = GetSolidFromGeometry(geometry);
                        if (solid != null && linkTransform != null)
                        {
                            solid = SolidUtils.CreateTransformed(solid, linkTransform);
                        }
                        
                        // ✅ MEMORY OPTIMIZATION: Add to LRU cache (will evict oldest if over limit)
                        AddToGeometryCache(cacheKey, solid);
                    }
                    
                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, line, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        
                        // ✅ CRITICAL FIX: Validate bounding box is not null before using it
                        if (bbox == null)
                        {
                            log?.Invoke($"[⚠️ SKIP-NULL-BBOX] Skipping intersection with null bounding box: Structural={structuralElement.Id}, IntersectionPoints={intersectionPoints.Count}");
                            continue;
                        }
                        
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
                        var center = BoundingBoxService.GetBoundingBoxCenter(bbox);
                        
                        // ✅ CRITICAL FIX: Validate intersection point is NOT zero before adding
                        if (center != null && Math.Abs(center.X) > 1e-9 && Math.Abs(center.Y) > 1e-9 && Math.Abs(center.Z) > 1e-9)
                        {
                            results.Add((structuralElement, bbox, center));
                        }
                        else
                        {
                            log?.Invoke($"[⚠️ SKIP-ZERO] Skipping intersection with zero center point: Structural={structuralElement.Id}, Center={center}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }
            
            log($"Spatial filtering: processed {processedCount}, skipped {spatiallyFilteredCount} elements via bounding box check");
            return results;
        }

        // Overload: accept a host-space Line (e.g. when the MEP element comes from a linked doc
        // and the caller has already transformed its curve into the active document coords).
        public static List<(Element, BoundingBoxXYZ, XYZ)> FindIntersections(
            Line hostLine,
            BoundingBoxXYZ? mepBoundingBox,
            List<(Element, Transform?)> structuralElements,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            if (hostLine == null)
            {
                log("ERROR: hostLine is null in FindIntersections overload.");
                return results;
            }

            // Derive a MEP bounding box from provided bbox or from the line
            BoundingBoxXYZ mepBBox = mepBoundingBox ?? new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(hostLine.GetEndPoint(0).X, hostLine.GetEndPoint(1).X), Math.Min(hostLine.GetEndPoint(0).Y, hostLine.GetEndPoint(1).Y), Math.Min(hostLine.GetEndPoint(0).Z, hostLine.GetEndPoint(1).Z)),
                Max = new XYZ(Math.Max(hostLine.GetEndPoint(0).X, hostLine.GetEndPoint(1).X), Math.Max(hostLine.GetEndPoint(0).Y, hostLine.GetEndPoint(1).Y), Math.Max(hostLine.GetEndPoint(0).Z, hostLine.GetEndPoint(1).Z))
            };

            // REVERTED: Back to original 1.0 foot tolerance since coordinate transform issue is fixed
            // Both MEP and wall bounding boxes are now in host shared coordinates
            // TEMPORARILY REVERTED: Using original 1.0ft tolerance for intersection detection
            const double tolerance = 1.0; // 1.0ft tolerance - original working value
            var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
            
            log($"[MepIntersectionService] REVERTED: Using {tolerance} foot tolerance for spatial pre-filtering (overload method)");
            log($"[MepIntersectionService] The real issue is coordinate system/transform problems, not spatial filtering");

            int processedCount = 0;
            int spatiallyFilteredCount = 0;
            int lastLoggedSkipCount = 0;

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;
                processedCount++;
                try
                {
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox != null)
                    {
                        // CRITICAL FIX: Transform structural bbox to host shared coordinates for distance check
                        if (linkTransform != null)
                        {
                            // Apply 8-corner transform to wall bounding box
                            
                            // Transform all 8 corners of the bounding box to shared coordinates
                            var pts = new[]
                            {
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Min.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Min.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Max.Y, structBBox.Min.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Min.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Max.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Min.X, structBBox.Max.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Min.Y, structBBox.Max.Z)),
                                linkTransform.OfPoint(new XYZ(structBBox.Max.X, structBBox.Max.Y, structBBox.Min.Z))
                            };
                            
                            structBBox = new BoundingBoxXYZ
                            {
                                Min = new XYZ(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Min(p => p.Z)),
                                Max = new XYZ(pts.Max(p => p.X), pts.Max(p => p.Y), pts.Max(p => p.Z))
                            };
                        }
                        
                        if (!BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        {
                            spatiallyFilteredCount++;
                            if (spatiallyFilteredCount - lastLoggedSkipCount >= 100)
                            {
                                lastLoggedSkipCount = spatiallyFilteredCount;
                                var wallType = structuralElement.GetType().Name;
                                var wallId = structuralElement.Id.IntegerValue;
                                var distance = GetDistanceToMepElement(mepBBox, structBBox, null);
                                log($"[MepIntersectionService] SPATIAL FILTER: Skipped {spatiallyFilteredCount} elements so far. Latest: {wallType} ID:{wallId} - Distance: {distance:F2}ft");
                            }
                            continue;
                        }
                    }

                    // ✅ MEMORY OPTIMIZATION: Get from LRU cache or compute
                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    Solid? solid;
                    if (!TryGetFromGeometryCache(cacheKey, out solid))
                    {
                        var options = new Options();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;
                        solid = GetSolidFromGeometry(geometry);
                        if (solid != null && linkTransform != null)
                            solid = SolidUtils.CreateTransformed(solid, linkTransform);
                        // ✅ MEMORY OPTIMIZATION: Add to LRU cache (will evict oldest if over limit)
                        AddToGeometryCache(cacheKey, solid);
                    }
                    if (solid == null) continue;

                    var intersectionPoints = GetIntersectionPoints(solid, hostLine, log);
                    if (intersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(intersectionPoints);
                        
                        // ✅ CRITICAL FIX: Validate bounding box is not null before using it
                        if (bbox == null)
                        {
                            log?.Invoke($"[⚠️ SKIP-NULL-BBOX] Skipping intersection with null bounding box: Structural={structuralElement.Id}, IntersectionPoints={intersectionPoints.Count}");
                            continue;
                        }
                        
                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
                        var center = BoundingBoxService.GetBoundingBoxCenter(bbox);
                        
                        // ✅ CRITICAL FIX: Validate intersection point is NOT zero before adding
                        if (center != null && Math.Abs(center.X) > 1e-9 && Math.Abs(center.Y) > 1e-9 && Math.Abs(center.Z) > 1e-9)
                        {
                            results.Add((structuralElement, bbox, center));
                        }
                        else
                        {
                            log?.Invoke($"[⚠️ SKIP-ZERO] Skipping intersection with zero center point: Structural={structuralElement.Id}, Center={center}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }

            log($"Spatial filtering: processed {processedCount}, skipped {spatiallyFilteredCount} elements via bounding box check");
            return results;
        }
        
        // ✅ OOP REFACTORING: Removed duplicate BoundingBoxesIntersect - now uses BoundingBoxService
        
        /// <summary>
        /// PHASE 2 OPTIMIZATION 2: Test if curve intersects bounding box
        /// Fast pre-check before expensive solid intersection
        /// 
        /// ✅ FIXED: More lenient tolerance and better line-box intersection test
        /// This prevents false negatives that were causing (0,0,0) intersection points
        /// </summary>
        private static bool TestCurveInBoundingBox(Line curve, BoundingBoxXYZ bbox, double tolerance)
        {
            // ✅ FIX: Use LARGER tolerance for this pre-filter to avoid false negatives
            // The actual geometry intersection is more precise, so be generous here
            double safetyMargin = tolerance * 2.0; // Double the tolerance for safety
            
            // Create expanded bounding box outline
            var expandedMin = new XYZ(
                bbox.Min.X - safetyMargin,
                bbox.Min.Y - safetyMargin,
                bbox.Min.Z - safetyMargin
            );
            var expandedMax = new XYZ(
                bbox.Max.X + safetyMargin,
                bbox.Max.Y + safetyMargin,
                bbox.Max.Z + safetyMargin
            );
            
            var structOutline = new Outline(expandedMin, expandedMax);
            
            // ✅ FIX: Test both endpoints AND curve midpoint
            // A curve can pass through a box even if both endpoints are outside
            var p1 = curve.GetEndPoint(0);
            var p2 = curve.GetEndPoint(1);
            var midPoint = curve.Evaluate(0.5, true);
            
            // Check if any point is inside the expanded box
            if (structOutline.ContainsPoint(p1, safetyMargin) ||
                structOutline.ContainsPoint(p2, safetyMargin) ||
                structOutline.ContainsPoint(midPoint, safetyMargin))
            {
                return true;
            }
            
            // ✅ FIX: Test if line actually intersects the box (not just endpoints)
            // Project line direction onto box axes to detect intersections
            XYZ lineDir = (p2 - p1).Normalize();
            
            // Test intersection with each face of the bounding box
            // X-axis faces
            if (Math.Abs(lineDir.X) > 1e-6)
            {
                double t1 = (expandedMin.X - p1.X) / lineDir.X;
                double t2 = (expandedMax.X - p1.X) / lineDir.X;
                
                if (IsPointOnLineSegmentInBox(p1, lineDir, t1, expandedMin, expandedMax) ||
                    IsPointOnLineSegmentInBox(p1, lineDir, t2, expandedMin, expandedMax))
                {
                    return true;
                }
            }
            
            // Y-axis faces
            if (Math.Abs(lineDir.Y) > 1e-6)
            {
                double t1 = (expandedMin.Y - p1.Y) / lineDir.Y;
                double t2 = (expandedMax.Y - p1.Y) / lineDir.Y;
                
                if (IsPointOnLineSegmentInBox(p1, lineDir, t1, expandedMin, expandedMax) ||
                    IsPointOnLineSegmentInBox(p1, lineDir, t2, expandedMin, expandedMax))
                {
                    return true;
                }
            }
            
            // Z-axis faces
            if (Math.Abs(lineDir.Z) > 1e-6)
            {
                double t1 = (expandedMin.Z - p1.Z) / lineDir.Z;
                double t2 = (expandedMax.Z - p1.Z) / lineDir.Z;
                
                if (IsPointOnLineSegmentInBox(p1, lineDir, t1, expandedMin, expandedMax) ||
                    IsPointOnLineSegmentInBox(p1, lineDir, t2, expandedMin, expandedMax))
                {
                    return true;
                }
            }
            
            // No intersection found
            return false;
        }
        
        /// <summary>
        /// Helper: Check if parametric point on line is within line segment bounds and inside box
        /// </summary>
        private static bool IsPointOnLineSegmentInBox(XYZ lineStart, XYZ lineDir, double t, XYZ boxMin, XYZ boxMax)
        {
            // Check if t is within line segment (0 to line length)
            if (t < -1e-6 || t > 1.0 + 1e-6) // Small tolerance
                return false;
            
            // Calculate point on line
            XYZ point = lineStart + lineDir * t;
            
            // Check if point is inside box
            return point.X >= boxMin.X && point.X <= boxMax.X &&
                   point.Y >= boxMin.Y && point.Y <= boxMax.Y &&
                   point.Z >= boxMin.Z && point.Z <= boxMax.Z;
        }

        // Extracts a solid from a geometry object
        private static Solid? GetSolidFromGeometry(GeometryElement geometry)
        {
            List<Solid> allSolids = new List<Solid>();
            
            foreach (GeometryObject geomObj in geometry)
            {
                if (geomObj is Solid s && s.Volume > 0)
                {
                    allSolids.Add(s);
                }
                else if (geomObj is GeometryInstance gi)
                {
                    foreach (GeometryObject instObj in gi.GetInstanceGeometry())
                    {
                        if (instObj is Solid s2 && s2.Volume > 0)
                        {
                            allSolids.Add(s2);
                        }
                    }
                }
            }
            
            // If no solids found, return null
            if (allSolids.Count == 0)
                return null;
            
            // For single solid, return it directly (original behavior)
            if (allSolids.Count == 1)
                return allSolids[0];
            
            // For multiple solids (compound walls), union them into one solid
            // This ensures we get one sleeve for the entire wall, not one per layer
            try
            {
                Solid resultSolid = allSolids[0];
                for (int i = 1; i < allSolids.Count; i++)
                {
                    resultSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        resultSolid, allSolids[i], BooleanOperationsType.Union);
                    
                    // Safety check: if union fails, fall back to first solid
                    if (resultSolid == null || resultSolid.Volume <= 0)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[MepIntersectionService] Wall union failed: Solid {i}/{allSolids.Count} union resulted in null or invalid volume. Falling back to first solid. This may result in individual sleeves per layer that will be clustered.");
                        resultSolid = allSolids[0];
                        break;
                    }
                }
                
                // Log successful union
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MepIntersectionService] Successfully united {allSolids.Count} wall layer solids into single solid");
                return resultSolid;
            }
            catch (Exception ex)
            {
                // If union fails for any reason, log and return the first solid as fallback
                // NOTE: Union failures are rare in Revit. If this occurs, only one sleeve will be placed
                // (first layer only). The clustering logic will handle multiple sleeves in nearby walls.
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[MepIntersectionService] Wall union failed with exception: {ex.Message}. Returning first solid only. Total layers detected: {allSolids.Count}.");
                return allSolids[0];
            }
        }

        // Intersects a solid with a line and returns the intersection points
        private static List<XYZ> GetIntersectionPoints(Solid solid, Line line, Action<string>? log = null)
        {
            var intersectionPoints = new List<XYZ>();
            try
            {
                int faceCount = solid.Faces.Size;
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"[Intersect] Solid face count = {faceCount}");
                foreach (Face face in solid.Faces)
                {
                    if (face == null) continue;
                    IntersectionResultArray? ira;
                    var res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        foreach (IntersectionResult ir in ira)
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                        if (intersectionPoints.Count > 0)
                        {
                            if (OptimizationFlags.UseDiagnosticMode)
                                log?.Invoke($"[Intersect] Found {intersectionPoints.Count} intersection point(s). First: {intersectionPoints[0]}");
                            // early exit optional? keep collecting for bbox
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"[Intersect] Exception while computing intersections: {ex.Message}");
            }
            return intersectionPoints;
        }

        // Creates a bounding box from a list of points
        private static BoundingBoxXYZ CreateBoundingBox(List<XYZ> points)
        {
            // ✅ CRITICAL FIX: Return null if points list is empty to prevent invalid bounding box
            if (points == null || points.Count == 0)
                return null;
            
            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
            foreach (var pt in points)
            {
                if (pt == null) continue;
                if (pt.X < minX) minX = pt.X;
                if (pt.Y < minY) minY = pt.Y;
                if (pt.Z < minZ) minZ = pt.Z;
                if (pt.X > maxX) maxX = pt.X;
                if (pt.Y > maxY) maxY = pt.Y;
                if (pt.Z > maxZ) maxZ = pt.Z;
            }
            
            // ✅ CRITICAL FIX: Validate that bounding box is valid (not all max/min values)
            if (minX == double.MaxValue || maxX == double.MinValue)
                return null;
            
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }

        // ✅ OOP REFACTORING: Removed duplicate GetBoundingBoxCenter - now uses BoundingBoxService.GetBoundingBoxCenter()

        // Collects structural elements within section box bounds only - MAJOR PERFORMANCE OPTIMIZATION
    public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc, Action<string> log, List<string>? selectedHostTypes = null)
        {
            var elements = new List<(Element, Transform?)>();
            log("Starting structural element collection.");

            // Get section box to drastically reduce search space
            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (doc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBox = view3D.GetSectionBox();
                    log($"Active view has a section box. Min: {sectionBox.Min}, Max: {sectionBox.Max}");
                }
                else
                {
                    log("No active section box found.");
                }
            }
            catch (Exception ex)
            {
                log($"Error getting section box: {ex.Message}");
            }

            // ✅ FIX: Filter structural categories based on UI host type selection
            var allCategories = new[] {
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Floors
            };

            var categories = allCategories;
            if (selectedHostTypes != null && selectedHostTypes.Count > 0)
            {
                var filteredCategories = new List<BuiltInCategory>();
                
                if (selectedHostTypes.Any(ht => ht.Equals("Walls", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_Walls);
                    log("UI Selection: Including Walls");
                }
                
                if (selectedHostTypes.Any(ht => ht.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_StructuralFraming);
                    log("UI Selection: Including Structural Framing");
                }
                
                if (selectedHostTypes.Any(ht => ht.Equals("Floors", StringComparison.OrdinalIgnoreCase)))
                {
                    filteredCategories.Add(BuiltInCategory.OST_Floors);
                    log("UI Selection: Including Floors");
                }
                
                if (filteredCategories.Count > 0)
                {
                    categories = filteredCategories.ToArray();
                    log($"✅ FILTERED: Only collecting {filteredCategories.Count} selected host types (was {allCategories.Length} total)");
                }
                else
                {
                    log("⚠️ WARNING: No valid host types selected, using all categories");
                }
            }
            else
            {
                log("No host type selection provided, collecting all structural categories");
            }

            try
            {
                // Use the existing CollectElements method from the command
                // This is already debugged and working
                var mepElements = new List<Element>();
                var wallElements = new List<Element>();
                
                // Get section box bounds
                if (sectionBox != null)
                {
                    Transform sectionTransform = sectionBox.Transform;
                    List<XYZ> corners = new List<XYZ>
                    {
                        sectionTransform.OfPoint(sectionBox.Min),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Max.Y, sectionBox.Min.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Max.Z)),
                        sectionTransform.OfPoint(new XYZ(sectionBox.Min.X, sectionBox.Max.Y, sectionBox.Max.Z)),
                        sectionBox.Max
                    };
                    
                    XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
                    XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));
                    
                    // Use the existing CollectElements method
                    CollectElements(doc, modelMin, modelMax, ref mepElements, ref wallElements);
                    
                    // Convert wall elements to the expected format
                    var links = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .ToList();
                    
                    foreach (var wall in wallElements)
                    {
                        // ✅ PHASE 2 OPTIMIZATION 3: Use cached transform
                        Transform? wallTransform = GetCachedTransform(wall.Document, links);
                        elements.Add((wall, wallTransform));
                    }
                }
                else
                {
                    // Fallback: collect all structural elements without section box filtering
                    var hostElements = new FilteredElementCollector(doc)
                        .WherePasses(new ElementMulticategoryFilter(categories))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    foreach (var e in hostElements) elements.Add((e, null));
                    
                    // Linked model elements
                    foreach (var link in new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>())
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) continue;

                        // ✅ PHASE 2 OPTIMIZATION 3: Use cached transform
                        var tr = GetCachedTransform(linkDoc, new List<RevitLinkInstance> { link });
                        var linked = new FilteredElementCollector(linkDoc)
                            .WherePasses(new ElementMulticategoryFilter(categories))
                            .WhereElementIsNotElementType()
                            .ToElements();
                        foreach (var e in linked) elements.Add((e, tr));
                    }
                }
                
                log($"Finished structural element collection. Total elements found: {elements.Count}");
                
                // ✅ FIX: Filter walls by minimum thickness if setting is enabled
                elements = FilterWallsByMinimumThicknessForTuples(elements).ToList();
                
                // ✅ FIX: Filter architectural floors if setting is enabled
                elements = FilterArchitecturalFloorsForTuples(elements).ToList();
                
                return elements;
            }
            catch (Exception ex)
            {
                log($"ERROR: Fallback structural collection failed: {ex.Message}");
                return elements;
            }
        }

        // Backwards-compatible overload: no-op logger
        public static List<(Element, Transform?)> CollectStructuralElementsForDirectIntersectionVisibleOnly(Document doc)
        {
            return CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, _ => { }, null);
        }
        
        // Copy of the working CollectElements method from the command
        private static void CollectElements(Document doc, XYZ modelMin, XYZ modelMax, 
                                         ref List<Element> mepElements, ref List<Element> wallElements)
        {
            Outline hostOutline = new Outline(modelMin, modelMax);

            // MEP categories
            BuiltInCategory[] mepCats = {
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
            };

            // Collect from host
            foreach (var cat in mepCats)
            {
                mepElements.AddRange(
                    new FilteredElementCollector(doc)
                        .OfCategory(cat)
                        .WhereElementIsNotElementType()
                        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                        .ToElements()
                );
            }

            // CRITICAL FIX: Also collect damper family instances that might not be properly categorized
            // as OST_DuctAccessory but are still dampers based on family name
            var damperFamilyInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                .Where(fi => {
                    var bbox = fi.get_BoundingBox(null);
                    if (bbox == null) return false;
                    return BoundingBoxService.BoundingBoxesIntersect(modelMin, modelMax, bbox.Min, bbox.Max);
                })
                .Cast<Element>()
                .ToList();
            
            mepElements.AddRange(damperFamilyInstances);

            // Collect walls and filter by minimum thickness if setting is enabled
            var collectedWalls = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType()
                .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
                .ToElements();
            
            // Filter by minimum wall thickness if setting is enabled
            collectedWalls = FilterWallsByMinimumThickness(collectedWalls).ToList();
            wallElements.AddRange(collectedWalls);

            // Collect from links
            var links = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var link in links)
            {
                Document linkDoc = link.GetLinkDocument();
                if (linkDoc == null) continue;

                // ✅ PHASE 2 OPTIMIZATION 3: Use cached transform
                Transform linkTransform = GetCachedTransform(linkDoc, links);
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

                Outline linkOutline = new Outline(actualMin, actualMax);

                foreach (var cat in mepCats)
                {
                    mepElements.AddRange(
                        new FilteredElementCollector(linkDoc)
                            .OfCategory(cat)
                            .WhereElementIsNotElementType()
                            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                            .ToElements()
                    );
                }

                // CRITICAL FIX: Also collect damper family instances from linked documents
                var linkedDamperFamilyInstances = new FilteredElementCollector(linkDoc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                    .Where(fi => {
                        var bbox = fi.get_BoundingBox(null);
                        if (bbox == null) return false;
                        // Transform the damper bbox to link coordinates for intersection test
                        var transformedMin = invTransform.OfPoint(bbox.Min);
                        var transformedMax = invTransform.OfPoint(bbox.Max);
                        var transformedBboxMin = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z));
                        var transformedBboxMax = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z));
                        return BoundingBoxService.BoundingBoxesIntersect(modelMin, modelMax, transformedBboxMin, transformedBboxMax);
                    })
                    .Cast<Element>()
                    .ToList();
                
                mepElements.AddRange(linkedDamperFamilyInstances);

                // Collect walls from link and filter by minimum thickness if setting is enabled
                var linkedWalls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
                    .ToElements();
                
                // Filter by minimum wall thickness if setting is enabled
                linkedWalls = FilterWallsByMinimumThickness(linkedWalls).ToList();
                wallElements.AddRange(linkedWalls);
            }
        }
        
        /// <summary>
        /// Collects damper family instances from host and linked documents for intersection detection
        /// </summary>
        public static List<(Element, Transform?)> CollectDamperElementsForIntersection(Document doc, Action<string> log)
        {
            var elements = new List<(Element, Transform?)>();
            log("Starting damper element collection for intersection detection.");

            // Get section box to filter by visible area
            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (doc.ActiveView is View3D view3D && view3D.IsSectionBoxActive)
                {
                    sectionBox = view3D.GetSectionBox();
                    log($"Active view has a section box for damper filtering. Min: {sectionBox.Min}, Max: {sectionBox.Max}");
                }
                else
                {
                    log("No active section box found for damper collection.");
                }
            }
            catch (Exception ex)
            {
                log($"Error getting section box for dampers: {ex.Message}");
            }

            // Collect dampers from host document
            try
            {
                var hostDampers = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                    .ToList();

                foreach (var damper in hostDampers)
                {
                    // Apply section box filtering if available
                    if (sectionBox != null)
                    {
                        var bbox = damper.get_BoundingBox(null);
                        if (bbox != null && !BoundingBoxService.BoundingBoxesIntersect(sectionBox.Min, sectionBox.Max, bbox.Min, bbox.Max))
                            continue;
                    }
                    elements.Add((damper, null));
                }

                log($"Found {hostDampers.Count} damper family instances in host document.");
            }
            catch (Exception ex)
            {
                log($"Error collecting host dampers: {ex.Message}");
            }

            // Collect dampers from linked documents
            try
            {
                var links = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var link in links)
                {
                    var linkDoc = link.GetLinkDocument();
                    if (linkDoc == null) continue;

                    var linkTransform = link.GetTotalTransform();
                    var linkedDampers = new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Damper") == true)
                        .ToList();

                    foreach (var damper in linkedDampers)
                    {
                        // Apply section box filtering if available (transform coordinates)
                        if (sectionBox != null)
                        {
                            var bbox = damper.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                var transformedMin = linkTransform.OfPoint(bbox.Min);
                                var transformedMax = linkTransform.OfPoint(bbox.Max);
                                var transformedBboxMin = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z));
                                var transformedBboxMax = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z));
                                
                                if (!BoundingBoxService.BoundingBoxesIntersect(sectionBox.Min, sectionBox.Max, transformedBboxMin, transformedBboxMax))
                                    continue;
                            }
                        }
                        elements.Add((damper, linkTransform));
                    }

                    log($"Found {linkedDampers.Count} damper family instances in linked document: {linkDoc.Title}");
                }
            }
            catch (Exception ex)
            {
                log($"Error collecting linked dampers: {ex.Message}");
            }

            log($"Finished damper element collection. Total dampers found: {elements.Count}");
            return elements;
        }

        // Helper method to calculate distance between MEP element and structural element
        private static double GetDistanceToMepElement(BoundingBoxXYZ mepBBox, BoundingBoxXYZ structBBox, Transform? linkTransform)
        {
            try
            {
                // Transform structural bbox if it's from a linked doc
                BoundingBoxXYZ transformedStructBBox = structBBox;
                if (linkTransform != null)
                {
                    var transformedMin = linkTransform.OfPoint(structBBox.Min);
                    var transformedMax = linkTransform.OfPoint(structBBox.Max);
                    transformedStructBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                        Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                    };
                }
                
                // Calculate center points
                var mepCenter = new XYZ(
                    (mepBBox.Min.X + mepBBox.Max.X) / 2,
                    (mepBBox.Min.Y + mepBBox.Max.Y) / 2,
                    (mepBBox.Min.Z + mepBBox.Max.Z) / 2
                );
                
                var structCenter = new XYZ(
                    (transformedStructBBox.Min.X + transformedStructBBox.Max.X) / 2,
                    (transformedStructBBox.Min.Y + transformedStructBBox.Max.Y) / 2,
                    (transformedStructBBox.Min.Z + transformedStructBBox.Max.Z) / 2
                );
                
                // Calculate distance in feet
                var distance = mepCenter.DistanceTo(structCenter);
                return UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Feet);
            }
            catch
            {
                return -1.0; // Return -1 if calculation fails
            }
        }

        public static List<(Element, BoundingBoxXYZ, XYZ)> FindDamperIntersections(
            Element damperElement,
            List<(Element, Transform?)> structuralElements,
            Transform? damperLinkTransform,
            Action<string> log)
        {
            var damperBBox = damperElement.get_BoundingBox(null);
            if (damperBBox == null)
            {
                log($"WARNING: Could not get bounding box for damper element {damperElement.Id}.");
                return new List<(Element, BoundingBoxXYZ, XYZ)>();
            }

            return FindDamperIntersectionsInternal(damperElement, damperBBox, structuralElements, damperLinkTransform, log);
        }

        private static List<(Element, BoundingBoxXYZ, XYZ)> FindDamperIntersectionsInternal(
            Element damperElement,
            BoundingBoxXYZ damperBBox,
            List<(Element, Transform?)> structuralElements,
            Transform? damperLinkTransform,
            Action<string> log)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();

            BoundingBoxXYZ hostDamperBBox = damperBBox;
            if (damperLinkTransform != null)
            {
                var transformed = BoundingBoxService.TransformBoundingBox(damperBBox, damperLinkTransform);
                if (transformed != null)
                {
                    hostDamperBBox = transformed;
                }
            }

            const double tolerance = 0.5; // 6 inches
            var expandedMin = new XYZ(
                hostDamperBBox.Min.X - tolerance,
                hostDamperBBox.Min.Y - tolerance,
                hostDamperBBox.Min.Z - tolerance);
            var expandedMax = new XYZ(
                hostDamperBBox.Max.X + tolerance,
                hostDamperBBox.Max.Y + tolerance,
                hostDamperBBox.Max.Z + tolerance);

            if (OptimizationFlags.UseDiagnosticMode)
                log($"[DamperIntersection] Processing element {damperElement.Id} with bbox Min=({hostDamperBBox.Min.X:F2}, {hostDamperBBox.Min.Y:F2}, {hostDamperBBox.Min.Z:F2}) Max=({hostDamperBBox.Max.X:F2}, {hostDamperBBox.Max.Y:F2}, {hostDamperBBox.Max.Z:F2})");

            foreach (var tuple in structuralElements)
            {
                Element structuralElement = tuple.Item1;
                Transform? linkTransform = tuple.Item2;

                try
                {
                    var structBBox = structuralElement.get_BoundingBox(null);
                    if (structBBox == null) continue;

                    if (linkTransform != null)
                    {
                        var transformedStruct = BoundingBoxService.TransformBoundingBox(structBBox, linkTransform);
                        if (transformedStruct != null)
                        {
                            structBBox = transformedStruct;
                        }
                    }

                    if (!BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max))
                        continue;

                    if (OptimizationFlags.UseDiagnosticMode)
                        log($"[DamperIntersection] Intersection candidate: damper {damperElement.Id} with structural {structuralElement.Id}");

                    var intersectionMin = new XYZ(
                        Math.Max(hostDamperBBox.Min.X, structBBox.Min.X),
                        Math.Max(hostDamperBBox.Min.Y, structBBox.Min.Y),
                        Math.Max(hostDamperBBox.Min.Z, structBBox.Min.Z));
                    var intersectionMax = new XYZ(
                        Math.Min(hostDamperBBox.Max.X, structBBox.Max.X),
                        Math.Min(hostDamperBBox.Max.Y, structBBox.Max.Y),
                        Math.Min(hostDamperBBox.Max.Z, structBBox.Max.Z));

                    if (intersectionMin.X > intersectionMax.X ||
                        intersectionMin.Y > intersectionMax.Y ||
                        intersectionMin.Z > intersectionMax.Z)
                    {
                        continue;
                    }

                    var intersectionBBox = new BoundingBoxXYZ
                    {
                        Min = intersectionMin,
                        Max = intersectionMax
                    };

                    var center = BoundingBoxService.GetBoundingBoxCenter(intersectionBBox);
                    results.Add((structuralElement, intersectionBBox, center));
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process damper intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }

            return results;
        }

        // ✅ OOP REFACTORING: Removed duplicate TransformBoundingBox - now uses BoundingBoxService.TransformBoundingBox()

        private static Line? GetElementLine(Element element, BoundingBoxXYZ mepBBox, Action<string> log)
        {
            if (element is FamilyInstance fi && fi.Symbol?.Family?.Name?.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                log($"[MepIntersectionService] Element {element.Id} identified as damper; using bounding-box intersection approach.");

                return null;
            }

            if (element.Location is LocationCurve locCurve && locCurve.Curve is Line curveLine)
            {
                return curveLine;
            }

            if (element is MEPCurve mepCurve)
            {
                try
                {
                    var connectors = mepCurve.ConnectorManager?.Connectors?.Cast<Connector>().Where(c => c != null).ToList();
                    if (connectors != null && connectors.Count >= 2)
                    {
                        var endpoints = connectors
                            .SelectMany((c, idx) => connectors
                                .Skip(idx + 1)
                                .Select(other => new { First = c, Second = other, Distance = c.Origin.DistanceTo(other.Origin) }))
                            .OrderByDescending(x => x.Distance)
                            .FirstOrDefault();

                        if (endpoints != null && endpoints.Distance > 0)
                        {
                            return Line.CreateBound(endpoints.First.Origin, endpoints.Second.Origin);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log($"[MepIntersectionService] Failed deriving line from MEPCurve connectors for element {element.Id}: {ex.Message}");
                }
            }

            try
            {
                var min = mepBBox.Min;
                var max = mepBBox.Max;
                var centerX = (min.X + max.X) * 0.5;
                var centerY = (min.Y + max.Y) * 0.5;
                var p1 = new XYZ(centerX, centerY, min.Z);
                var p2 = new XYZ(centerX, centerY, max.Z);

                if (p1.DistanceTo(p2) < 1e-6)
                {
                    p1 = new XYZ(min.X, centerY, (min.Z + max.Z) * 0.5);
                    p2 = new XYZ(max.X, centerY, (min.Z + max.Z) * 0.5);
                }

                if (p1.DistanceTo(p2) < 1e-6)
                {
                    p2 = new XYZ(p1.X + 1.0, p1.Y, p1.Z);
                }

                if (OptimizationFlags.UseDiagnosticMode)
                    log($"[MepIntersectionService] Fallback line derived from bounding box for element {element.Id}.");
                return Line.CreateBound(p1, p2);
            }
            catch (Exception ex)
            {
                log($"[MepIntersectionService] Failed to derive fallback line for element {element.Id}: {ex.Message}");
                return null;
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
                
                double minThicknessInternal = UnitUtils.ConvertToInternalUnits(minThicknessMm, UnitTypeId.Millimeters);
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
                            if (OptimizationFlags.UseDiagnosticMode)
                            {
                                double wallThicknessMm = UnitUtils.ConvertFromInternalUnits(wallThickness, UnitTypeId.Millimeters);
                                System.Diagnostics.Debug.WriteLine($"[MepIntersectionService] SKIP: Wall {wall.Id.IntegerValue} thickness {wallThicknessMm:F1}mm < {minThicknessMm:F1}mm minimum");
                            }
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MepIntersectionService] Filtered {skippedCount} walls below {minThicknessMm:F1}mm minimum thickness");
                }
                
                return filteredWalls;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MepIntersectionService] Error filtering walls by minimum thickness: {ex.Message}");
                return walls; // Return original list on error
            }
        }

        /// <summary>
        /// Filters walls by minimum thickness setting (for tuple format)
        /// Skips walls that are thinner than the user-specified minimum
        /// </summary>
        private static List<(Element, Transform?)> FilterWallsByMinimumThicknessForTuples(List<(Element, Transform?)> elements)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                double minThicknessMm = settings.MinWallThickness;
                
                // If setting is 0 or negative, don't filter (all walls allowed)
                if (minThicknessMm <= 0)
                    return elements;
                
                double minThicknessInternal = UnitUtils.ConvertToInternalUnits(minThicknessMm, UnitTypeId.Millimeters);
                var filteredElements = new List<(Element, Transform?)>();
                int skippedCount = 0;
                
                foreach (var (element, transform) in elements)
                {
                    if (element is Wall wallObj)
                    {
                        double wallThickness = wallObj.Width;
                        
                        if (wallThickness >= minThicknessInternal)
                        {
                            filteredElements.Add((element, transform));
                        }
                        else
                        {
                            skippedCount++;
                            if (OptimizationFlags.UseDiagnosticMode)
                            {
                                double wallThicknessMm = UnitUtils.ConvertFromInternalUnits(wallThickness, UnitTypeId.Millimeters);
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Log($"[MepIntersectionService] SKIP: Wall {element.Id.IntegerValue} thickness {wallThicknessMm:F1}mm < {minThicknessMm:F1}mm minimum");
                            }
                        }
                    }
                    else
                    {
                        // Not a wall, add it
                        filteredElements.Add((element, transform));
                    }
                }
                
                if (skippedCount > 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MepIntersectionService] Filtered {skippedCount} walls below {minThicknessMm:F1}mm minimum thickness");
                }
                
                return filteredElements;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MepIntersectionService] Error filtering walls by minimum thickness: {ex.Message}");
                return elements; // Return original list on error
            }
        }

        /// <summary>
        /// Filters architectural floors if the setting is enabled (for tuple format)
        /// Skips floors where Structural parameter is not checked
        /// </summary>
        private static List<(Element, Transform?)> FilterArchitecturalFloorsForTuples(List<(Element, Transform?)> elements)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                bool ignoreArchFloors = settings.IgnoreArchitecturalFloors;
                
                // If setting is disabled, don't filter (all floors allowed)
                if (!ignoreArchFloors)
                    return elements;
                
                var filteredElements = new List<(Element, Transform?)>();
                int skippedCount = 0;
                
                foreach (var (element, transform) in elements)
                {
                    if (element is Floor floor)
                    {
                        // Check Structural parameter
                        Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                        bool isStructural = structuralParam?.AsInteger() == 1;
                        
                        if (isStructural)
                        {
                            filteredElements.Add((element, transform));
                        }
                        else
                        {
                            skippedCount++;
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[MepIntersectionService] SKIP: Architectural floor {floor.Id.IntegerValue} (Structural parameter not checked)");
                        }
                    }
                    else
                    {
                        // Not a floor, add it (walls, framing, etc.)
                        filteredElements.Add((element, transform));
                    }
                }
                
                if (skippedCount > 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MepIntersectionService] Filtered {skippedCount} architectural floors");
                }
                
                return filteredElements;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MepIntersectionService] Error filtering architectural floors: {ex.Message}");
                return elements; // Return original list on error
            }
        }
    }
}