#nullable enable
#if !REVIT2024_OR_GREATER
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static partial class MepIntersectionService
    {
        // ✅ MEMORY OPTIMIZATION: LRU cache with max size to prevent unbounded growth
        // Large linked files can have thousands of structural elements - cache can grow to 100MB+ without limits
        private const int MAX_GEOMETRY_CACHE_SIZE = 5000; // Limit to 5000 entries (~10-20MB typical)
        private static readonly Lazy<Dictionary<string, Solid?>> _geometryCache = new Lazy<Dictionary<string, Solid?>>(() => new Dictionary<string, Solid?>());
        private static readonly Lazy<LinkedList<string>> _geometryCacheOrder = new Lazy<LinkedList<string>>(() => new LinkedList<string>()); // LRU tracking
        // Multi-solid cache (stores all layer solids) guarded by flag UseMultiSolidCache
        private static readonly Lazy<Dictionary<string, List<Solid>>> _geometryMultiSolidCache = new Lazy<Dictionary<string, List<Solid>>>(() => new Dictionary<string, List<Solid>>());

        // DIAGNOSTIC: Build stamp + static constructor instrumentation to trace type initialization issues in R2024.
        // If a TypeInitializationException persists, this logging will confirm whether our static constructor executes.
        private static readonly string _buildStamp = "MepIntersectionService BuildStamp 2025-11-20_17-10";
        // Intersection debug flag and log path (lightweight instrumentation for 2024 failure investigation)
        private static string IntersectionDebugFlagFile;
        private static string IntersectionDebugLogPath;
        private static bool IntersectionDebugEnabled;


        
        // LAZY INITIALIZATION - avoid static constructor that can throw
        private static IRevitUnitConversionService? _unitConverter;
        private static IRevitUnitConversionService UnitConverter
        {
            get
            {
                if (_unitConverter == null)
                {
#if REVIT2024_OR_GREATER
                    // In Revit 2024+, the UnitTypeId API exists but causes issues
                    // Use manual conversion until we can safely detect API availability at runtime
                    _unitConverter = new FallbackUnitConverter();
#else
                    // Revit 2023 and earlier - always use manual conversion
                    _unitConverter = new FallbackUnitConverter();
#endif
                }
                return _unitConverter;
            }
        }

        private sealed class FallbackUnitConverter : IRevitUnitConversionService
        {
            private const double MillimetersPerFoot = 304.8;

            public double ToInternalMillimeters(double value) => value / MillimetersPerFoot;
            public double FromInternalMillimeters(double value) => value * MillimetersPerFoot;
            public double ToInternalFeet(double value) => value;
            public double FromInternalFeet(double value) => value;
        }
        
        // ✅ TWO-TIER OPTIMIZATION: Spatial partitioning service for Tier 1 filtering
        // Only initialized when UseSpatialGrid flag is enabled
        private static SpatialPartitioningService? _spatialService = null;
        
        // PHASE 2 OPTIMIZATION 3: Transform cache (1.5x speedup)
        private static readonly Lazy<Dictionary<Document, Transform>> _transformCache = new Lazy<Dictionary<Document, Transform>>(() => new Dictionary<Document, Transform>());
        
        // PHASE 1 OPTIMIZATION 2: Category Whitelist (2x speedup)
        // ✅ CRITICAL FIX: Initialize in static constructor with validation to prevent TypeInitializationException in Revit 2024
        private static readonly BuiltInCategory[] MEP_CATEGORY_WHITELIST;
        private static readonly BuiltInCategory[] STRUCTURAL_CATEGORY_WHITELIST;
        
        static MepIntersectionService()
        {
            try
            {
                // MINIMAL static constructor - avoid ANY external dependencies
                IntersectionDebugFlagFile = "enable_intersection_debug.flag";
                IntersectionDebugLogPath = "intersection_debug_2024.log";
                IntersectionDebugEnabled = false; // Will be set lazily on first use

                // Initialize whitelists with safe fallback
                var mepList = new List<BuiltInCategory>
                {
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
                MEP_CATEGORY_WHITELIST = mepList.ToArray();

                var structList = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFoundation
                };
                STRUCTURAL_CATEGORY_WHITELIST = structList.ToArray();
            }
            catch (Exception ex)
            {
                // Last ditch logging for static init failure
                try 
                {
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logPath = System.IO.Path.Combine(appData, "JSE_MEP_Openings", "STATIC_INIT_FAIL.log");
                    var logDir = System.IO.Path.GetDirectoryName(logPath);
                    if (!System.IO.Directory.Exists(logDir)) System.IO.Directory.CreateDirectory(logDir);
                    
                    System.IO.File.WriteAllText(logPath, $"Static Init Failed: {ex}\nInner: {ex.InnerException}");
                }
                catch { }
                
                // Assign empty arrays to prevent null ref later, though the app is likely doomed
                MEP_CATEGORY_WHITELIST = new BuiltInCategory[0];
                STRUCTURAL_CATEGORY_WHITELIST = new BuiltInCategory[0];
                
                // Re-throw to ensure we don't fail silently
                throw; 
            }
        }
        
        // ✅ MEMORY OPTIMIZATION: LRU cache add with eviction
        private static void AddToGeometryCache(string key, Solid? solid)
        {
            // Remove if exists (move to end = most recently used)
            if (_geometryCache.Value.ContainsKey(key))
            {
                _geometryCacheOrder.Value.Remove(key);
                _geometryCache.Value.Remove(key);
            }
            
            // Add to cache
            _geometryCache.Value[key] = solid;
            _geometryCacheOrder.Value.AddLast(key);
            
            // ✅ MEMORY OPTIMIZATION: Evict oldest entries if over limit
            while (_geometryCache.Value.Count > MAX_GEOMETRY_CACHE_SIZE && _geometryCacheOrder.Value.First != null)
            {
                var oldestKey = _geometryCacheOrder.Value.First.Value;
                _geometryCache.Value.Remove(oldestKey);
                _geometryCacheOrder.Value.RemoveFirst();
            }
        }
        
        // ✅ MEMORY OPTIMIZATION: Get from cache and update LRU order
        private static bool TryGetFromGeometryCache(string key, out Solid? solid)
        {
            if (_geometryCache.Value.TryGetValue(key, out solid))
            {
                // Move to end (most recently used)
                _geometryCacheOrder.Value.Remove(key);
                _geometryCacheOrder.Value.AddLast(key);
                return true;
            }
            return false;
        }
        private static bool TryGetFromMultiSolidCache(string key, out List<Solid> solids)
        {
            solids = new List<Solid>();
            if (!OptimizationFlags.UseMultiSolidCache) return false;
            if (_geometryMultiSolidCache.Value.TryGetValue(key, out var stored) && stored != null && stored.Count > 0)
            {
                solids = stored;
                return true;
            }
            return false;
        }
        
        // ✅ STEP 3 OPTIMIZATION: Add to multi-solid cache (compound walls support)
        private static void AddToGeometryMultiSolidCache(string key, List<Solid> solids)
        {
            if (!OptimizationFlags.UseMultiSolidCache) return;
            if (solids == null || solids.Count == 0) return;
            
            // Store list of solids (handles compound walls with multiple layers)
            _geometryMultiSolidCache.Value[key] = new List<Solid>(solids); // Defensive copy
        }
        
        // Clear cache method for memory management
        public static void ClearGeometryCache()
        {
            if (_geometryCache.IsValueCreated) _geometryCache.Value.Clear();
            if (_geometryCacheOrder.IsValueCreated) _geometryCacheOrder.Value.Clear();
            if (_geometryMultiSolidCache.IsValueCreated) _geometryMultiSolidCache.Value.Clear();
        }
        
        // ✅ MEMORY OPTIMIZATION: Get cache statistics for monitoring
        public static (int count, int maxSize, double memoryEstimateMB) GetGeometryCacheStats()
        {
            // Rough estimate: 2-5 MB per 1000 entries (depends on solid complexity)
            double memoryEstimateMB = _geometryCache.Value.Count * 0.003; // 3KB per entry average
            return (_geometryCache.Value.Count, MAX_GEOMETRY_CACHE_SIZE, memoryEstimateMB);
        }
        
        /// <summary>
        /// PHASE 2 OPTIMIZATION 3: Get cached transform for a document
        /// </summary>
        public static Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null)
        {
            if (!_transformCache.Value.ContainsKey(doc))
            {
                var link = links.FirstOrDefault(l => l.GetLinkDocument()?.Title == doc.Title);
                if (link != null)
                {
                    var transform = link.GetTotalTransform();
                    _transformCache.Value[doc] = transform;
                    log?.Invoke($"[TransformCache] Cached transform for document: {doc.Title}");
                }
                else
                {
                    _transformCache.Value[doc] = Transform.Identity;
                    log?.Invoke($"[TransformCache] No link found for document: {doc.Title}, using Identity transform");
                }
            }
            return _transformCache.Value[doc];
        }
        
        /// <summary>
        /// PHASE 2 OPTIMIZATION 3: Clear transform cache
        /// </summary>
        public static void ClearTransformCache()
        {
            if (_transformCache.IsValueCreated) _transformCache.Value.Clear();
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
            // ✅ CRITICAL FIX: Check category FIRST - Duct Accessories are dampers
            // This ensures all Duct Accessories are detected, even if family name doesn't contain "Damper"
            if (element?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
            {
                // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                if (element is FamilyInstance fi && fi.Symbol != null)
                {
                    var familyName = fi.Symbol.Family?.Name ?? "";
                    var typeName = fi.Symbol.Name ?? "";
                    var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                    
                    if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                        return false;
                }
                
                // ✅ All other Duct Accessories are treated as dampers
                return true;
            }
            
            // ✅ FALLBACK: Check family name for non-DuctAccessory elements (shouldn't happen, but safe)
            if (element is FamilyInstance fi2)
            {
                var familyName = fi2.Symbol?.Family?.Name ?? "";
                var typeName = fi2.Symbol?.Name ?? "";
                var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                
                // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                    return false;
                
                return familyName.Contains("Damper", StringComparison.OrdinalIgnoreCase);
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
                        for (int i = 0; i < entriesToRemove && _geometryCacheOrder.Value.First != null; i++)
                        {
                            var oldestKey = _geometryCacheOrder.Value.First.Value;
                            _geometryCache.Value.Remove(oldestKey);
                            _geometryCacheOrder.Value.RemoveFirst();
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

            // ✅ DIAGNOSTIC LOG: Use SafeFileLogger (should log to safefilelogger_diagnostic.log)
            try
            {
                var buildTime = System.IO.File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location);
                SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"🔨 DLL BUILD TIME: {buildTime:yyyy-MM-dd HH:mm:ss}");
                SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[{DateTime.Now:HH:mm:ss.fff}] FindIntersectionsBatchInternal CALLED - MEP={mepElements.Count}, Structural={structuralElements.Count}");
            }
            catch (Exception diagEx)
            {
                // If SafeFileLogger fails, try direct write as fallback
                try
                {
                    var versionTag = Helpers.VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = System.IO.Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!System.IO.Directory.Exists(logDir)) System.IO.Directory.CreateDirectory(logDir);
                    var logPath = System.IO.Path.Combine(logDir, "DIAGNOSTIC_TEST.log");
                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] FindIntersectionsBatchInternal CALLED (fallback) - MEP={mepElements.Count}, Structural={structuralElements.Count}\n");
                }
                catch { }
            }

            // ✅ PERFORMANCE DIAGNOSTICS: Track timing and statistics
            var overallStopwatch = System.Diagnostics.Stopwatch.StartNew();
            var preprocessStopwatch = System.Diagnostics.Stopwatch.StartNew();
            int totalCacheHits = 0;
            int totalCacheMisses = 0;
            int totalSpatiallyFiltered = 0;
            int totalGeometryComputations = 0;
            int totalKnownPairsSkipped = 0;
            int totalIntersectionTests = 0;
            // Curve pre-filter instrumentation (cheap line vs bbox rejection stats)
            int curvePreFilterTested = 0;
            int curvePreFilterRejected = 0;
            // Spatial tier metrics aggregation
            int totalTier1NearbyElements = 0;
            int totalTier2PreciseCandidates = 0;
            int totalTier2Rejected = 0;
            // Geometry timing accumulators (ms)
            long totalGeometryExtractionMs = 0;
            long totalSolidEnumerationMs = 0;
            long totalSolidTransformMs = 0;
            long totalIntersectionCalcMs = 0;
            long totalPrecomputeMs = 0;
            int precomputeElements = 0;

            // Lazy debug flag check on first use
            if (!IntersectionDebugEnabled)
            {
                try { IntersectionDebugEnabled = System.IO.File.Exists(IntersectionDebugFlagFile); } catch { }
            }
            if (IntersectionDebugEnabled)
            {
                SafeFileLogger.SafeAppendText(IntersectionDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ENTER FindIntersectionsBatchInternal MEP={mepElements.Count} STRUCT={structuralElements.Count} KnownPairs={knownValidPairs?.Count ?? 0} SkipKnownGeom={skipKnownPairsGeometryCheck}\n");
            }

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
                string cacheKey = BuildStableCacheKey(structElement, structTransform);
                if (IntersectionDebugEnabled && structuralData.Count < 50)
                {
                    double w = structBBox.Max.X - structBBox.Min.X;
                    double h = structBBox.Max.Y - structBBox.Min.Y;
                    double d = structBBox.Max.Z - structBBox.Min.Z;
                    SafeFileLogger.SafeAppendText(IntersectionDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] STRUCT[{structElement.Id.IntegerValue}] BBOX_ft W={w:F3} H={h:F3} D={d:F3} Cat={(BuiltInCategory)structElement.Category.Id.IntegerValue}\n");
                }
                structuralData.Add((structElement, structTransform, structBBox, cacheKey));
            }
            if (IntersectionDebugEnabled)
            {
                SafeFileLogger.SafeAppendText(IntersectionDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Structural preprocessed count={structuralData.Count}\n");
            }
            preprocessStopwatch.Stop();

            // ✅ DIAGNOSTIC LOG: Preprocessing complete
            try
            {
                SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[{DateTime.Now:HH:mm:ss.fff}] PREPROCESSING COMPLETE - Structural elements processed: {structuralData.Count}, Time: {preprocessStopwatch.ElapsedMilliseconds}ms");
            }
            catch { }

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

            // ✅ TWO-TIER OPTIMIZATION: Initialize spatial service if enabled
            var spatialBuildStopwatch = System.Diagnostics.Stopwatch.StartNew();

            // 🔍 DIAGNOSTIC: Show flag status at runtime
            try { SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[SPATIAL_GRID_INIT] OptimizationFlags.UseSpatialGrid = {OptimizationFlags.UseSpatialGrid}"); } catch { }

            if (OptimizationFlags.UseSpatialGrid)
            {
                // ✅ ADAPTIVE GRID SIZING: Calculate optimal grid size based on structural element distribution
                // Target: 3-5 elements per cell on average for best performance
                double gridSize = 1.0; // Default fallback

                if (structuralData.Count > 0)
                {
                    // Calculate bounding box of all structural elements (work area extent)
                    double minX = structuralData.Min(sd => sd.bbox.Min.X);
                    double maxX = structuralData.Max(sd => sd.bbox.Max.X);
                    double minY = structuralData.Min(sd => sd.bbox.Min.Y);
                    double maxY = structuralData.Max(sd => sd.bbox.Max.Y);
                    double minZ = structuralData.Min(sd => sd.bbox.Min.Z);
                    double maxZ = structuralData.Max(sd => sd.bbox.Max.Z);

                    double width = maxX - minX;
                    double height = maxY - minY;
                    double depth = maxZ - minZ;

                    // Calculate volume and element density
                    double volume = width * height * depth;
                    double elementsPerCubicFoot = structuralData.Count / Math.Max(volume, 1.0);

                    // Target: 4 elements per cell (sweet spot for spatial filtering)
                    const double TARGET_ELEMENTS_PER_CELL = 4.0;
                    double idealGridSize = Math.Pow(TARGET_ELEMENTS_PER_CELL / Math.Max(elementsPerCubicFoot, 0.001), 1.0 / 3.0);

                    // Clamp to reasonable range: 0.5 ft (tight spaces) to 5.0 ft (large spaces)
                    gridSize = Math.Max(0.5, Math.Min(idealGridSize, 5.0));

                    if (OptimizationFlags.UseDiagnosticMode)
                    {
                        log($"[SpatialGrid] Work area: {width:F1}×{height:F1}×{depth:F1} ft, {structuralData.Count} elements, density: {elementsPerCubicFoot:F3} elem/ft³");
                        log($"[SpatialGrid] Adaptive grid size: {gridSize:F2} ft (target {TARGET_ELEMENTS_PER_CELL} elements/cell)");
                    }
                }

                _spatialService = new SpatialPartitioningService(gridSize);
                // Build grid with structural data (need to convert to format expected by BuildGrid)
                var structuralDataForGrid = structuralData.Select(sd => (sd.element, sd.transform, sd.bbox, (Solid?)null)).ToList();
                _spatialService.BuildGrid(structuralDataForGrid);
                spatialBuildStopwatch.Stop();

                try { SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[SPATIAL_GRID_INIT] _spatialService created successfully"); } catch { }

                if (OptimizationFlags.UseDiagnosticMode)
                {
                    var (totalCells, usedCells, avgElements) = _spatialService.GetStatistics();
                    log($"[TwoTier] Spatial grid initialized in {spatialBuildStopwatch.ElapsedMilliseconds}ms: {usedCells} cells used, avg {avgElements:F1} elements per cell");
                }
            }
            else
            {
                spatialBuildStopwatch.Stop();
                try { SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[SPATIAL_GRID_INIT] Spatial grid DISABLED (flag is false)"); } catch { }
            }

            // OPTIONAL PRECOMPUTE: Pre-extract host solids (can improve later lookups when many MEP elements share same hosts)
            if (OptimizationFlags.PrecomputeHostSolids)
            {
                var precomputeSw = System.Diagnostics.Stopwatch.StartNew();
                foreach (var (element, transform, bbox, cacheKey) in structuralData)
                {
                    // Skip if already in multi-solid cache
                    if (OptimizationFlags.UseMultiSolidCache && TryGetFromMultiSolidCache(cacheKey, out _))
                    {
                        continue;
                    }
                    if (TryGetFromGeometryCache(cacheKey, out _))
                    {
                        continue; // single solid cached already
                    }
                    var geomSw = System.Diagnostics.Stopwatch.StartNew();
                    try
                    {
                        var options = Helpers.GeometryOptionsFactory.CreateIntersectionOptions();
                        var geometry = element.get_Geometry(options);
                        geomSw.Stop();
                        totalGeometryExtractionMs += geomSw.ElapsedMilliseconds;
                        if (geometry == null) continue;
                        var enumSw = System.Diagnostics.Stopwatch.StartNew();
                        var solids = GetSolidsFromGeometry(geometry);
                        enumSw.Stop();
                        totalSolidEnumerationMs += enumSw.ElapsedMilliseconds;
                        if (solids != null && solids.Count > 0 && transform != null)
                        {
                            var txSw = System.Diagnostics.Stopwatch.StartNew();
                            var transformed = new List<Solid>();
                            foreach (var s in solids)
                            {
                                if (s != null) transformed.Add(SolidUtils.CreateTransformed(s, transform));
                            }
                            solids = transformed;
                            txSw.Stop();
                            totalSolidTransformMs += txSw.ElapsedMilliseconds;
                        }
                        // Cache first solid for legacy callers + full list for multi-solid use
                        AddToGeometryCache(cacheKey, solids != null && solids.Count > 0 ? solids[0] : null);
                        AddToGeometryMultiSolidCache(cacheKey, solids);
                        precomputeElements++;
                    }
                    catch { }
                    // Memory guard: abort precompute if cache grows too large
                    var (_, _, cacheMBCurrent) = GetGeometryCacheStats();
                    if (cacheMBCurrent > 20.0) // threshold MB
                    {
                        log?.Invoke($"[PrecomputeHostSolids] Aborting precompute early – cache reached {cacheMBCurrent:F1} MB");
                        break;
                    }
                }
                precomputeSw.Stop();
                totalPrecomputeMs = precomputeSw.ElapsedMilliseconds;
                log?.Invoke($"[PrecomputeHostSolids] Precomputed {precomputeElements} structural geometries in {totalPrecomputeMs}ms");
            }

            if (OptimizationFlags.UseParallelClashSearch)
            {
                // ==========================================================================================
                // PARALLEL BROAD PHASE STRATEGY (Multithreading Safe)
                // ==========================================================================================
                // Phase 1: Extraction (Main Thread)
                // We already extracted 'structuralData'. Now extract MEP data.
                // (Done below in the loop, but we need to pull it out to parallelize)

                var mepDataList = new List<(Element mepElement, Transform? mepTransform, BoundingBoxXYZ mepBBox, Line? line, bool isDamper)>();
                // Also keep track of indices for results

                foreach (var (mepElement, mepTransform) in mepElements)
                {
                    var mepBBox = mepElement.get_BoundingBox(null);
                    if (mepBBox == null) continue;

                    // Transform logic (copied from downstream)
                    if (mepTransform != null)
                    {
                        var tMin = mepTransform.OfPoint(mepBBox.Min);
                        var tMax = mepTransform.OfPoint(mepBBox.Max);
                        mepBBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(Math.Min(tMin.X, tMax.X), Math.Min(tMin.Y, tMax.Y), Math.Min(tMin.Z, tMax.Z)),
                            Max = new XYZ(Math.Max(tMin.X, tMax.X), Math.Max(tMin.Y, tMax.Y), Math.Max(tMin.Z, tMax.Z))
                        };
                    }

                    bool isDamper = IsDamperElement(mepElement) || mepElement.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory;
                    Line? line = null;
                    if (!isDamper)
                    {
                        var lineResult = GetElementLineWithSource(mepElement, mepBBox, null); // Log null for speed
                        line = lineResult.line;
                        // Transform line if needed
                        if (mepTransform != null && !lineResult.isFallbackLine && line != null)
                        {
                            line = Line.CreateBound(mepTransform.OfPoint(line.GetEndPoint(0)), mepTransform.OfPoint(line.GetEndPoint(1)));
                        }
                    }

                    mepDataList.Add((mepElement, mepTransform, mepBBox, line, isDamper));
                }

                // Phase 2: Parallel Search (Multi-Thread)
                // Find Candidates: (MEP_Index, Struct_Index)
                var candidatesToCheck = new System.Collections.Concurrent.ConcurrentBag<(int mepIdx, int structIdx)>();

                System.Threading.Tasks.Parallel.ForEach(mepDataList, (mepData, state, index) =>
                {
                    // capture index properly
                    int mIdx = (int)index;
                    var mepBBox = mepData.mepBBox;
                    double mepCenterZ = (mepBBox.Min.Z + mepBBox.Max.Z) * 0.5;
                    double MAX_V_SEP = 20.0;
                    const double tolerance = 0.2;

                    var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
                    var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);
                    var expandedBBox = new BoundingBoxXYZ { Min = expandedMin, Max = expandedMax };

                    // Candidates source
                    IList<(Element element, Transform? transform, BoundingBoxXYZ bbox, string cacheKey)> structsToScan = structuralData;

                    // Use Spatial Grid if available (Thread safe read?) - Yes usually, as strictly read-only after build.
                    if (_spatialService != null && OptimizationFlags.UseSpatialGrid)
                    {
                        var nearby = _spatialService.GetNearbyElements(expandedBBox); // Assumes thread-safety
                                                                                      // Map back to structuralData indices or objects?
                                                                                      // This implementation of SpatialService returns structs containing Element.
                                                                                      // We need to match them back to 'structuralData' or just use them.
                                                                                      // SpatialService returns (Element, Transform, BBox, Solid). 
                                                                                      // But we need the index in 'structuralData' OR just the object.
                                                                                      // Let's iterate structuralData for safety unless we map indices.
                                                                                      // For simplicity in this broad phase refactor: fallback to full scan OR unsafe spatial.
                                                                                      // *Safe Plan*: Use Spatial Service's BBox check logic locally if possible.
                                                                                      // Actually, let's stick to simple Z-Filter + Iterate All for Parallel. 
                                                                                      // Why? Because iterating 5000 structs in 12 threads is faster than overhead of spatial mapping sometimes.
                                                                                      // But Spatial is better. 
                                                                                      // Let's use the 'nearby' from spatial service if possible. 
                                                                                      // The SpatialService probably returns a List of *internal* structs. 
                                                                                      // Let's assume _spatialService.GetNearbyElements is thread-safe (reads generic list/dict). 
                                                                                      // Note: We need 'structuralData' items. 

                        // Re-implementation of simple spatial filter for parallel loop to be safe:
                        // Proceed with iterating 'structsToScan' = structuralData (Brute force parallel is fast enough usually)
                    }

                    // Brute force check in parallel (N*M / Cores)
                    for (int sIdx = 0; sIdx < structuralData.Count; sIdx++)
                    {
                        var sData = structuralData[sIdx];

                        // Z-Filter
                        double sCenterZ = (sData.bbox.Min.Z + sData.bbox.Max.Z) * 0.5;
                        if (Math.Abs(mepCenterZ - sCenterZ) > MAX_V_SEP) continue;

                        // BBox Overlap
                        if (mepBBox.Max.X < sData.bbox.Min.X || mepBBox.Min.X > sData.bbox.Max.X) continue;
                        if (mepBBox.Max.Y < sData.bbox.Min.Y || mepBBox.Min.Y > sData.bbox.Max.Y) continue;
                        if (mepBBox.Max.Z < sData.bbox.Min.Z || mepBBox.Min.Z > sData.bbox.Max.Z) continue;

                        // It's a candidate!
                        candidatesToCheck.Add((mIdx, sIdx));
                    }
                });

                // Phase 3: Validation (Main Thread)
                // Process confirmed candidates
                foreach (var pair in candidatesToCheck)
                {
                    var mepEntry = mepDataList[pair.mepIdx];
                    var structEntry = structuralData[pair.structIdx];

                    // Handle Dampers
                    if (mepEntry.isDamper)
                    {
                        // Call damper logic (needs re-implementation or calling existing method)
                        // For now, fall back to existing method logic inside the loop?
                        // No, we already have the pair.
                        // We can just call "FindIntersection" logic.
                        // Actually, Dampers return multiple "blades".
                        // We might need to just run standard damper logic for Dampers.
                        // Let's call FindDamperIntersectionsInternal for this pair? 
                        // Easier: If it's a damper, skip this candidate logic and let it run FULL loop?
                        // Optimization: Only run damper logic for overlapping bounding boxes.
                        var specificStructList = new List<(Element, Transform?)> { (structEntry.element, structEntry.transform) };
                        var damperRes = FindDamperIntersectionsInternal(mepEntry.mepElement, mepEntry.mepBBox, specificStructList, mepEntry.mepTransform, log);
                        results.AddRange(damperRes.Select(i => (mepEntry.mepElement, i.Item1, i.Item2, i.Item3)));
                        continue;
                    }

                    // Handle Normal Elements (Line vs Solid)
                    // Verify Line vs BBox first (fast)
                    if (mepEntry.line != null)
                    {
                        // transform struct BBox to check line? 
                        // No, structural BBox is in Host Coords. Line is in Host Coords.
                        // Check curve?
                        // if (OptimizationFlags.UseCurveInBoundingBoxFilter) ...
                        // We can use the cached line.
                    }

                    // Final Solid Check
                    // Fetch Solid (This triggers the lazy load/transform on Main Thread)
                    Solid? structSolid = null;

                    // Use Cache Logic
                    if (TryGetFromGeometryCache(structEntry.cacheKey, out var cachedSolid)) structSolid = cachedSolid;
                    else
                    {
                        // Compute and Cache (extracted from original loop)
                        var options = Helpers.GeometryOptionsFactory.CreateIntersectionOptions();
                        var g = structEntry.element.get_Geometry(options);
                        var solids = GetSolidsFromGeometry(g);
                        if (solids != null && solids.Count > 0)
                        {
                            if (structEntry.transform != null)
                                structSolid = SolidUtils.CreateTransformed(solids[0], structEntry.transform);
                            else structSolid = solids[0];
                            AddToGeometryCache(structEntry.cacheKey, structSolid);
                        }
                    }

                    if (structSolid == null) continue;

                    // INTERSECT
                    // (Requires 'line' for MEP or Solid for MEP)
                    // If we have a line:
                    var filter = new ElementIntersectsSolidFilter(structSolid); // Wait, this filter is for Collector.
                                                                                // We need manual intersection:
                                                                                // BooleanOperations? Or SolidCurveIntersection?
                    if (mepEntry.line != null)
                    {
                        using (var sci = structSolid.IntersectWithCurve(mepEntry.line, new SolidCurveIntersectionOptions()))
                        {
                            if (sci.SegmentCount > 0)
                            {
                                // ✅ PARALLEL FIX: Extract actual intersection points from SCI
                                // Previous code incorrectly returned XYZ.Zero using a 'Point calc needed?' placeholder
                                var sciPoints = new List<XYZ>();
                                for (int i = 0; i < sci.SegmentCount; i++)
                                {
                                    var curve = sci.GetCurveSegment(i);
                                    sciPoints.Add(curve.GetEndPoint(0));
                                    sciPoints.Add(curve.GetEndPoint(1));
                                }
                                
                                // Calculate proper intersection bbox and center (Using local static helper)
                                var intsBBox = CreateBoundingBox(sciPoints);
                                if (intsBBox != null)
                                {
                                    var intsCenter = BoundingBoxService.GetBoundingBoxCenter(intsBBox);
                                    
                                    // Validate center is not Zero and add to results
                                    if (intsCenter != null && (Math.Abs(intsCenter.X) > 1e-9 || Math.Abs(intsCenter.Y) > 1e-9 || Math.Abs(intsCenter.Z) > 1e-9))
                                    {
                                        results.Add((mepEntry.mepElement, structEntry.element, intsBBox, intsCenter));
                                        log?.Invoke($"[PARALLEL-FIX] Fixed Zero-Point at Center={intsCenter}");
                                    }
                                }
                            }
                        }
                    }
                }

                return results;
            }

            // ==========================================================================================
            // END PARALLEL STRATEGY - FALLBACK TO LEGACY (Sequential)
            // ==========================================================================================

            var mepProcessingStopwatch = System.Diagnostics.Stopwatch.StartNew();
            int mepIndex = 0;
            foreach (var (mepElement, mepTransform) in mepElements)
            {
                // Pre-calculation for MEP element
                var perMepStopwatch = System.Diagnostics.Stopwatch.StartNew();
                int spatiallyFiltered = 0;
                int geometrySkippedForKnownPairs = 0;
                int mepCacheHits = 0;
                int mepCacheMisses = 0;
                int nearbyElementsCount = 0;
                int preciseCandidatesCount = 0;
                int rtreeFiltered = 0;
                curvePreFilterTested = 0;
                curvePreFilterRejected = 0;
                const double tolerance = 1.0;
                
                var mepBBox = mepElement.get_BoundingBox(null);
                if (mepBBox == null) continue;

                // Apply transform if needed
                if (mepTransform != null)
                {
                    var tMin = mepTransform.OfPoint(mepBBox.Min);
                    var tMax = mepTransform.OfPoint(mepBBox.Max);
                    mepBBox = new BoundingBoxXYZ
                    {
                        Min = new XYZ(Math.Min(tMin.X, tMax.X), Math.Min(tMin.Y, tMax.Y), Math.Min(tMin.Z, tMax.Z)),
                        Max = new XYZ(Math.Max(tMin.X, tMax.X), Math.Max(tMin.Y, tMax.Y), Math.Max(tMin.Z, tMax.Z))
                    };
                }

                var expandedMin = new XYZ(mepBBox.Min.X - 1.0, mepBBox.Min.Y - 1.0, mepBBox.Min.Z - 1.0);
                var expandedMax = new XYZ(mepBBox.Max.X + 1.0, mepBBox.Max.Y + 1.0, mepBBox.Max.Z + 1.0);

                var lineResult = GetElementLineWithSource(mepElement, mepBBox, log);
                var line = lineResult.line;
                if (mepTransform != null && line != null && !lineResult.isFallbackLine)
                {
                    line = Line.CreateBound(mepTransform.OfPoint(line.GetEndPoint(0)), mepTransform.OfPoint(line.GetEndPoint(1)));
                }

                // ✅ CRITICAL FIX: Skip MEP element if line is null (e.g., dampers or failed line extraction)
                // Without this check, GetIntersectionPoints gets called with null line, returns empty points,
                // creates null bbox, gets (0,0,0) center, triggers FIXUP in ClashZoneService
                if (line == null)
                {
                    if (OptimizationFlags.UseDiagnosticMode)
                        log?.Invoke($"[SKIP-NULL-LINE] Skipping MEP element {mepElement.Id} - line is null (damper or failed line extraction)");
                    continue; // Skip to next MEP element
                }

                // Inner Loop: Structural Elements
                foreach (var structEntry in structuralData)
                {
                    var structElement = structEntry.element;
                    var structTransform = structEntry.transform;
                    var structBBox = structEntry.bbox;
                    var cacheKey = structEntry.cacheKey;



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

                    // ✅ PRIORITY 1 OPTIMIZATION: Re-enable TestCurveInBoundingBox filter for cheap rejection
                    // Fast curve-in-bbox test before expensive solid intersection (10-15% faster)
                    // Controlled by optimization flag for safe rollout
                    if (OptimizationFlags.UseCurveInBoundingBoxFilter)
                    {
                        curvePreFilterTested++;
                        if (!TestCurveInBoundingBox(line, structBBox, tolerance))
                        {
                            curvePreFilterRejected++;
                            spatiallyFiltered++;
                            continue; // Skip expensive solid intersection
                        }
                    }

                    // ✅ MEMORY OPTIMIZATION: Lazy solid loading - only compute after all cheap checks pass
                    // cacheKey is already available from foreach loop deconstruction

                    // ✅ PERFORMANCE PROFILING: Track solid extraction time (major bottleneck)
                    var solidExtractionStopwatch = System.Diagnostics.Stopwatch.StartNew(); // includes cache retrieval

                    // ✅ R2024 FIX: Get ALL solids instead of trying to union them (BooleanOperations fails in R2024)
                    // For compound walls, this returns multiple solids (one per layer)
                    // We check intersection against ALL layers to avoid missing intersections
                    List<Solid> solids = null;
                    totalIntersectionTests++;

                    // MULTI-SOLID CACHE: Attempt list retrieval first if enabled
                    if (OptimizationFlags.UseMultiSolidCache && TryGetFromMultiSolidCache(cacheKey, out var cachedList))
                    {
                        solids = cachedList;
                        mepCacheHits++;
                        totalCacheHits++;
                        solidExtractionStopwatch.Stop();
                    }
                    else if (!TryGetFromGeometryCache(cacheKey, out var cachedSolid))
                    {
                        // Cache miss - compute solids now (only for elements that passed all filters)
                        mepCacheMisses++;
                        totalCacheMisses++;
                        totalGeometryComputations++;

                        // ✅ PERFORMANCE PROFILING: Track geometry extraction (expensive operation)
                        var geometryExtractionStopwatch = System.Diagnostics.Stopwatch.StartNew();
                        var options = Helpers.GeometryOptionsFactory.CreateIntersectionOptions();
                        var geometry = structElement.get_Geometry(options);
                        geometryExtractionStopwatch.Stop();
                        totalGeometryExtractionMs += geometryExtractionStopwatch.ElapsedMilliseconds;

                        if (geometry != null)
                        {
                            // ✅ PERFORMANCE PROFILING: Track solid extraction from geometry
                            var solidExtractionFromGeometryStopwatch = System.Diagnostics.Stopwatch.StartNew();
                            solids = GetSolidsFromGeometry(geometry);
                            solidExtractionFromGeometryStopwatch.Stop();
                            totalSolidEnumerationMs += solidExtractionFromGeometryStopwatch.ElapsedMilliseconds;

                            // ✅ PERFORMANCE PROFILING: Log slow geometry operations
                            if (OptimizationFlags.UseDiagnosticMode && (geometryExtractionStopwatch.ElapsedMilliseconds > 50 || solidExtractionFromGeometryStopwatch.ElapsedMilliseconds > 50))
                            {
                                log($"[PROFILING] Slow geometry extraction: Element={structElement.Id}, Geometry={geometryExtractionStopwatch.ElapsedMilliseconds}ms, Solids={solidExtractionFromGeometryStopwatch.ElapsedMilliseconds}ms, Count={solids?.Count ?? 0}");
                            }

                            // Transform all solids if needed
                            if (solids != null && solids.Count > 0 && structTransform != null)
                            {
                                var transformStopwatch = System.Diagnostics.Stopwatch.StartNew();
                                var transformedSolids = new List<Solid>();
                                foreach (var s in solids)
                                {
                                    if (s != null)
                                    {
                                        transformedSolids.Add(SolidUtils.CreateTransformed(s, structTransform));
                                    }
                                }
                                solids = transformedSolids;
                                transformStopwatch.Stop();
                                totalSolidTransformMs += transformStopwatch.ElapsedMilliseconds;

                                if (OptimizationFlags.UseDiagnosticMode && transformStopwatch.ElapsedMilliseconds > 20)
                                {
                                    log($"[PROFILING] Slow solid transform: Element={structElement.Id}, Time={transformStopwatch.ElapsedMilliseconds}ms, Solids={solids.Count}");
                                }
                            }

                            // Cache the first solid for backwards compatibility with existing cache structure
                            AddToGeometryCache(cacheKey, solids != null && solids.Count > 0 ? solids[0] : null);
                            AddToGeometryMultiSolidCache(cacheKey, solids);
                            solidExtractionStopwatch.Stop();
                        }
                    }
                    else
                    {
                        mepCacheHits++;
                        totalCacheHits++;
                        // For now, wrap cached single solid in a list
                        // TODO: Update cache structure to store List<Solid>
                        solids = cachedSolid != null ? new List<Solid> { cachedSolid } : new List<Solid>();
                        solidExtractionStopwatch.Stop();
                    }

                    if (solids == null || solids.Count == 0) continue;

                    // ✅ R2024 FIX: Check intersection against ALL solids (for compound walls with multiple layers)
                    // ✅ PERFORMANCE PROFILING: Track intersection point calculation (major bottleneck)
                    var intersectionCalculationStopwatch = System.Diagnostics.Stopwatch.StartNew();
                    var allIntersectionPoints = new List<XYZ>();
                    int solidIntersectionCount = 0;
                    foreach (var solid in solids)
                    {
                        if (solid == null || solid.Volume <= 0) continue;

                        var layerIntersectionStopwatch = System.Diagnostics.Stopwatch.StartNew();
                        var layerIntersectionPoints = GetIntersectionPoints(solid, line, log);
                        layerIntersectionStopwatch.Stop();
                        solidIntersectionCount++;

                        // ✅ PERFORMANCE PROFILING: Log slow intersection calculations
                        if (OptimizationFlags.UseDiagnosticMode && layerIntersectionStopwatch.ElapsedMilliseconds > 50)
                        {
                            log($"[PROFILING] Slow intersection: Element={structElement.Id}, Solid={solidIntersectionCount}, Time={layerIntersectionStopwatch.ElapsedMilliseconds}ms, Points={layerIntersectionPoints?.Count ?? 0}");
                        }

                        if (layerIntersectionPoints != null && layerIntersectionPoints.Count > 0)
                        {
                            allIntersectionPoints.AddRange(layerIntersectionPoints);
                        }
                    }
                    intersectionCalculationStopwatch.Stop();
                    totalIntersectionCalcMs += intersectionCalculationStopwatch.ElapsedMilliseconds;

                    // ✅ PERFORMANCE PROFILING: Log total intersection calculation time
                    if (OptimizationFlags.UseDiagnosticMode && intersectionCalculationStopwatch.ElapsedMilliseconds > 100)
                    {
                        log($"[PROFILING] Slow total intersection: Element={structElement.Id}, Time={intersectionCalculationStopwatch.ElapsedMilliseconds}ms, Solids={solids.Count}, Points={allIntersectionPoints.Count}");
                    }

                    if (allIntersectionPoints.Count > 0)
                    {
                        var bbox = CreateBoundingBox(allIntersectionPoints);

                        // ✅ CRITICAL FIX: Validate bounding box is not null before using it
                        if (bbox == null)
                        {
                            log?.Invoke($"[⚠️ SKIP-NULL-BBOX] Skipping intersection with null bounding box: MEP={mepElement.Id}, Structural={structElement.Id}, IntersectionPoints={allIntersectionPoints.Count}");
                            continue;
                        }

                        // Use bounding box center (average of entry/exit points) to get mid-depth of host
                        var center = BoundingBoxService.GetBoundingBoxCenter(bbox);

                        // ✅ UNCONDITIONAL DEBUG: Always log intersection point calculation
                        log?.Invoke($"[INTERSECTION-CALC] MEP={mepElement.Id}, Structural={structElement.Id}, Points={allIntersectionPoints.Count}, BBox={(bbox != null ? $"Min={bbox.Min}, Max={bbox.Max}" : "NULL")}, Center={center}");

                        // ✅ CRITICAL FIX: Validate intersection point is NOT zero before adding
                        // This prevents creating clash zones with (0,0,0) intersection points
                        if (center != null && Math.Abs(center.X) > 1e-9 && Math.Abs(center.Y) > 1e-9 && Math.Abs(center.Z) > 1e-9)
                        {
                            results.Add((mepElement, structElement, bbox, center));

                            if (OptimizationFlags.UseDiagnosticMode && solids.Count > 1)
                            {
                                log?.Invoke($"[R2024-MULTILAYER] Found intersection in compound wall with {solids.Count} layers: MEP={mepElement.Id}, Structural={structElement.Id}, TotalPoints={allIntersectionPoints.Count}");
                            }
                        }
                        else
                        {
                            // Log warning but don't add invalid intersection
                            log?.Invoke($"[⚠️ SKIP-ZERO] Skipping intersection with zero center point: MEP={mepElement.Id}, Structural={structElement.Id}, Center={center}");
                        }
                    }

                    // ✅ PERFORMANCE DIAGNOSTICS: Track aggregates
                    totalSpatiallyFiltered += spatiallyFiltered;
                    totalKnownPairsSkipped += geometrySkippedForKnownPairs;
                    // Aggregate spatial tier metrics
                    totalTier1NearbyElements += nearbyElementsCount;
                    totalTier2PreciseCandidates += preciseCandidatesCount;
                    totalTier2Rejected += rtreeFiltered;

                    if (geometrySkippedForKnownPairs > 0 && OptimizationFlags.UseDiagnosticMode)
                    {
                        log($"[OPTIMIZATION] MEP {mepElement.Id}: Skipped geometry checks for {geometrySkippedForKnownPairs} known valid pairs");
                    }

                    perMepStopwatch.Stop();
                    if (OptimizationFlags.UseDiagnosticMode)
                    {
                        var cacheHitRate = mepCacheHits + mepCacheMisses > 0 ? 100.0 * mepCacheHits / (mepCacheHits + mepCacheMisses) : 0;
                        log($"[PERF] MEP {mepIndex}/{mepElements.Count} ({mepElement.Id}): {perMepStopwatch.ElapsedMilliseconds}ms, CacheHits={mepCacheHits}, CacheMisses={mepCacheMisses}, HitRate={cacheHitRate:F1}%");

                        if (OptimizationFlags.UseSpatialGrid && _spatialService != null)
                        {
                            log($"[BatchIntersection] MEP {mepElement.Id}: spatially filtered {spatiallyFiltered}/{structuralData.Count} structural elements after two-tier filtering");

                            // ✅ TWO-TIER SUMMARY: Log filtering effectiveness
                            if (OptimizationFlags.UseRTreeFilter && rtreeFiltered > 0)
                            {
                                var totalFiltered = structuralData.Count - structuralData.Count;
                                var tier1Rejected = structuralData.Count - nearbyElementsCount;
                                var tier2Rejected = rtreeFiltered;
                                log($"[TwoTier] Summary MEP {mepElement.Id}: Tier1 rejected {tier1Rejected}, Tier2 rejected {tier2Rejected}, Total rejected {totalFiltered}/{structuralData.Count} ({100.0 * totalFiltered / structuralData.Count:F1}%)");
                            }
                        }
                        else
                        {
                            log($"[BatchIntersection] MEP {mepElement.Id}: spatially filtered {spatiallyFiltered}/{structuralData.Count} structural elements");
                        }
                    }
                }
            }

                mepProcessingStopwatch.Stop();
                overallStopwatch.Stop();

                // ✅ DIAGNOSTIC LOG: Processing complete
                try
                {
                    SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[{DateTime.Now:HH:mm:ss.fff}] PROCESSING COMPLETE - Found {results.Count} intersections, Total time: {overallStopwatch.ElapsedMilliseconds}ms, MEP processing: {mepProcessingStopwatch.ElapsedMilliseconds}ms");
                    SafeFileLogger.SafeAppendText("DIAGNOSTIC_TEST.log", $"[{DateTime.Now:HH:mm:ss.fff}] PERFORMANCE STATS - Cache hits: {totalCacheHits}, Cache misses: {totalCacheMisses}, Spatially filtered: {totalSpatiallyFiltered}, Geometry computations: {totalGeometryComputations}");
                }
                catch { }

                if (OptimizationFlags.UseDiagnosticMode)
                {
                    log($"[BatchIntersection] Found {results.Count} total intersections");
                    log($"\n========== PERFORMANCE SUMMARY ==========");
                    log($"Overall Time: {overallStopwatch.ElapsedMilliseconds}ms ({overallStopwatch.Elapsed.TotalSeconds:F2}s)");
                    log($"  - Preprocessing: {preprocessStopwatch.ElapsedMilliseconds}ms ({100.0 * preprocessStopwatch.ElapsedMilliseconds / Math.Max(overallStopwatch.ElapsedMilliseconds, 1):F1}%)");
                    log($"  - Spatial Build: {spatialBuildStopwatch.ElapsedMilliseconds}ms ({100.0 * spatialBuildStopwatch.ElapsedMilliseconds / Math.Max(overallStopwatch.ElapsedMilliseconds, 1):F1}%)");
                    log($"  - MEP Processing: {mepProcessingStopwatch.ElapsedMilliseconds}ms ({100.0 * mepProcessingStopwatch.ElapsedMilliseconds / Math.Max(overallStopwatch.ElapsedMilliseconds, 1):F1}%)");
                    log($"\nCache Statistics:");
                    log($"  - Cache Hits: {totalCacheHits}");
                    log($"  - Cache Misses: {totalCacheMisses}");
                    var overallHitRate = totalCacheHits + totalCacheMisses > 0 ? 100.0 * totalCacheHits / (totalCacheHits + totalCacheMisses) : 0;
                    log($"  - Hit Rate: {overallHitRate:F1}%");
                    log($"  - Geometry Computations: {totalGeometryComputations}");
                    log($"\nFiltering Effectiveness:");
                    log($"  - Total Intersection Tests: {totalIntersectionTests}");
                    log($"  - Spatially Filtered: {totalSpatiallyFiltered}");
                    log($"  - Known Pairs Skipped: {totalKnownPairsSkipped}");
                    if (OptimizationFlags.UseSpatialGrid)
                    {
                        log($"  - Tier1 Nearby Elements (sum): {totalTier1NearbyElements}");
                        if (OptimizationFlags.UseRTreeFilter)
                        {
                            log($"  - Tier2 Precise Candidates (sum): {totalTier2PreciseCandidates}");
                            log($"  - Tier2 Rejected (sum): {totalTier2Rejected}");
                        }
                    }
                    if (OptimizationFlags.UseCurveInBoundingBoxFilter)
                    {
                        var curvePreFilterRate = curvePreFilterTested > 0 ? 100.0 * curvePreFilterRejected / curvePreFilterTested : 0.0;
                        log($"  - Curve Pre-Filter Tested: {curvePreFilterTested}");
                        log($"  - Curve Pre-Filter Rejected: {curvePreFilterRejected} ({curvePreFilterRate:F1}% rejection)");
                    }
                    var avgTimePerMep = mepElements.Count > 0 ? mepProcessingStopwatch.ElapsedMilliseconds / (double)mepElements.Count : 0;
                    log($"  - Avg Time/MEP: {avgTimePerMep:F1}ms");
                    var zonesPerSecond = overallStopwatch.Elapsed.TotalSeconds > 0 ? results.Count / overallStopwatch.Elapsed.TotalSeconds : 0;
                    log($"  - Throughput: {zonesPerSecond:F1} zones/second");
                    log($"========================================\n");
                }

                // Geometry extraction metrics summary (flag controlled)
                if (OptimizationFlags.LogGeometryExtractionMetrics)
                {
                    log($"[GeometryMetrics] Extraction={totalGeometryExtractionMs}ms, SolidEnum={totalSolidEnumerationMs}ms, SolidTransform={totalSolidTransformMs}ms, IntersectionCalc={totalIntersectionCalcMs}ms, Precompute={totalPrecomputeMs}ms");
                    var (cacheCount2, cacheMax2, cacheMb2) = GetGeometryCacheStats();
                    log($"[GeometryMetrics] CacheEntries={cacheCount2}/{cacheMax2} (~{cacheMb2:F1}MB) HitRate={(totalCacheHits + totalCacheMisses > 0 ? 100.0 * totalCacheHits / (totalCacheHits + totalCacheMisses) : 0):F1}%");
                    if (OptimizationFlags.PrecomputeHostSolids)
                    {
                        log($"[GeometryMetrics] PrecomputedElements={precomputeElements}");
                    }
                }
                return results;

        }

        // Stable cache key builder (deterministic) guarded by flag
        private static string BuildStableCacheKey(Element e, Transform? t)
        {
            if (!OptimizationFlags.UseStableGeometryCacheKeys)
            {
                return $"{e.Id.IntegerValue}_{t?.GetHashCode() ?? 0}"; // legacy unstable
            }
            // Use UniqueId + rounded origin & basis vectors for transform (if any)
            if (t == null)
            {
                return $"{e.UniqueId}_A"; // Active doc
            }
            XYZ o = t.Origin;
            XYZ bx = t.BasisX; XYZ by = t.BasisY; XYZ bz = t.BasisZ;
            string fmt(XYZ v, int dpPos, int dpDir) => $"{Math.Round(v.X, dpPos)},{Math.Round(v.Y, dpPos)},{Math.Round(v.Z, dpPos)}";
            // Origin 3dp, basis 2dp
            return $"{e.UniqueId}_{fmt(o,3,3)}_{Math.Round(bx.X,2)},{Math.Round(bx.Y,2)},{Math.Round(bx.Z,2)}_{Math.Round(by.X,2)},{Math.Round(by.Y,2)},{Math.Round(by.Z,2)}_{Math.Round(bz.X,2)},{Math.Round(bz.Y,2)},{Math.Round(bz.Z,2)}";
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
                    // ✅ R2024 FIX: Get ALL solids for compound walls
                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    List<Solid> solids;
                    
                    if (!TryGetFromGeometryCache(cacheKey, out var cachedSolid))
                    {
                        var options = Helpers.GeometryOptionsFactory.CreateIntersectionOptions();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;

                        solids = GetSolidsFromGeometry(geometry);
                        if (solids != null && solids.Count > 0 && linkTransform != null)
                        {
                            var transformedSolids = new List<Solid>();
                            foreach (var s in solids)
                            {
                                if (s != null) transformedSolids.Add(SolidUtils.CreateTransformed(s, linkTransform));
                            }
                            solids = transformedSolids;
                        }
                        
                        // ✅ MEMORY OPTIMIZATION: Add to LRU cache (will evict oldest if over limit)
                        AddToGeometryCache(cacheKey, solids != null && solids.Count > 0 ? solids[0] : null);
                    }
                    else
                    {
                        solids = cachedSolid != null ? new List<Solid> { cachedSolid } : new List<Solid>();
                    }
                    
                    if (solids == null || solids.Count == 0) continue;

                    // ✅ R2024 FIX: Check intersection against ALL solids (for compound walls)
                    var allIntersectionPoints = new List<XYZ>();
                    foreach (var solid in solids)
                    {
                        if (solid == null || solid.Volume <= 0) continue;
                        var layerPoints = GetIntersectionPoints(solid, line, log);
                        if (layerPoints != null && layerPoints.Count > 0) allIntersectionPoints.AddRange(layerPoints);
                    }
                    var intersectionPoints = allIntersectionPoints;
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
                    // ✅ R2024 FIX: Get ALL solids for compound walls
                    string cacheKey = $"{structuralElement.Id.IntegerValue}_{linkTransform?.GetHashCode() ?? 0}";
                    List<Solid> solids;
                    if (!TryGetFromGeometryCache(cacheKey, out var cachedSolid))
                    {
                        var options = Helpers.GeometryOptionsFactory.CreateIntersectionOptions();
                        var geometry = structuralElement.get_Geometry(options);
                        if (geometry == null) continue;
                        solids = GetSolidsFromGeometry(geometry);
                        if (solids != null && solids.Count > 0 && linkTransform != null)
                        {
                            var transformedSolids = new List<Solid>();
                            foreach (var s in solids)
                            {
                                if (s != null) transformedSolids.Add(SolidUtils.CreateTransformed(s, linkTransform));
                            }
                            solids = transformedSolids;
                        }
                        // ✅ MEMORY OPTIMIZATION: Add to LRU cache (will evict oldest if over limit)
                        AddToGeometryCache(cacheKey, solids != null && solids.Count > 0 ? solids[0] : null);
                    }
                    else
                    {
                        solids = cachedSolid != null ? new List<Solid> { cachedSolid } : new List<Solid>();
                    }
                    if (solids == null || solids.Count == 0) continue;

                    // ✅ R2024 FIX: Check intersection against ALL solids (for compound walls)
                    var allIntersectionPoints = new List<XYZ>();
                    foreach (var solid in solids)
                    {
                        if (solid == null || solid.Volume <= 0) continue;
                        var layerPoints = GetIntersectionPoints(solid, hostLine, log);
                        if (layerPoints != null && layerPoints.Count > 0) allIntersectionPoints.AddRange(layerPoints);
                    }
                    var intersectionPoints = allIntersectionPoints;
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
            
            // ✅ FIX: Test both endpoints AND curve midpoint
            // A curve can pass through a box even if both endpoints are outside
            var p1 = curve.GetEndPoint(0);
            var p2 = curve.GetEndPoint(1);
            var midPoint = curve.Evaluate(0.5, true);
            
            // ✅ FIX: Check if any point is inside the expanded box using manual check
            // Outline.ContainsPoint doesn't exist in Revit API, so we check manually
            if (IsPointInsideBoundingBox(p1, expandedMin, expandedMax) ||
                IsPointInsideBoundingBox(p2, expandedMin, expandedMax) ||
                IsPointInsideBoundingBox(midPoint, expandedMin, expandedMax))
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
        /// Helper: Check if a point is inside a bounding box
        /// </summary>
        private static bool IsPointInsideBoundingBox(XYZ point, XYZ boxMin, XYZ boxMax)
        {
            return point.X >= boxMin.X && point.X <= boxMax.X &&
                   point.Y >= boxMin.Y && point.Y <= boxMax.Y &&
                   point.Z >= boxMin.Z && point.Z <= boxMax.Z;
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

        // Extracts all solids from a geometry object (returns list for compound walls with multiple layers)
        // ✅ R2024 FIX: Return ALL solids instead of trying to union them (BooleanOperations fails in R2024)
        private static List<Solid> GetSolidsFromGeometry(GeometryElement geometry)
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
            
            return allSolids;
        }
        
        // Legacy method for backwards compatibility - returns first solid only
        // NOTE: This may miss intersections in compound walls if MEP passes through a different layer
        // Consider using GetSolidsFromGeometry instead for R2024 compatibility
        private static Solid? GetSolidFromGeometry(GeometryElement geometry)
        {
            var solids = GetSolidsFromGeometry(geometry);
            return solids.Count > 0 ? solids[0] : null;
        }

        // Intersects a solid with a line and returns the intersection points
        private static List<XYZ> GetIntersectionPoints(Solid solid, Line line, Action<string>? log = null)
        {
            var intersectionPoints = new List<XYZ>();
            
            // ✅ CRITICAL DEBUG: Log input parameters
            if (OptimizationFlags.UseDiagnosticMode)
            {
                log?.Invoke($"[GetIntersectionPoints] ENTER: Solid={(solid != null ? $"Valid(Volume={solid.Volume})" : "NULL")}, Line={(line != null ? $"Valid({line.GetEndPoint(0)} to {line.GetEndPoint(1)})" : "NULL")}");
            }
            
            try
            {
                // ✅ CRITICAL FIX: Check if line is null before using it
                if (line == null)
                {
                    log?.Invoke($"[GetIntersectionPoints] ❌ Line is NULL - returning empty intersection points");
                    return intersectionPoints;
                }
                
                int faceCount = solid.Faces.Size;
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"[Intersect] Solid face count = {faceCount}");
                    
                int facesChecked = 0;
                int facesWithIntersections = 0;
                
                foreach (Face face in solid.Faces)
                {
                    if (face == null) continue;
                    facesChecked++;
                    
                    IntersectionResultArray? ira;
                    var res = face.Intersect(line, out ira);
                    
                    if (OptimizationFlags.UseDiagnosticMode && res != SetComparisonResult.Disjoint)
                    {
                        log?.Invoke($"[Intersect] Face {facesChecked}/{faceCount}: Result={res}, IRA={(ira != null ? $"Size={ira.Size}" : "NULL")}");
                    }
                    
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        facesWithIntersections++;
                        foreach (Autodesk.Revit.DB.IntersectionResult ir in ira)
                        {
                            var pt = GetIntersectionPointFromRevitResult(ir);
                            intersectionPoints.Add(pt);
                            
                            if (OptimizationFlags.UseDiagnosticMode)
                                log?.Invoke($"[Intersect] Found intersection point: {pt}");
                        }
                    }
                }
                
                if (OptimizationFlags.UseDiagnosticMode)
                {
                    log?.Invoke($"[GetIntersectionPoints] EXIT: FacesChecked={facesChecked}, FacesWithIntersections={facesWithIntersections}, TotalPoints={intersectionPoints.Count}");
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[Intersect] Exception while computing intersections: {ex.Message}");
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"[Intersect] Stack trace: {ex.StackTrace}");
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
                    .Where(fi => 
                    {
                        var familyName = fi.Symbol?.Family?.Name ?? "";
                        var typeName = fi.Symbol?.Name ?? "";
                        var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                        
                        // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                        if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                            return false;
                        
                        return familyName.Contains("Damper");
                    })
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
                        .Where(fi => 
                        {
                            var familyName = fi.Symbol?.Family?.Name ?? "";
                            var typeName = fi.Symbol?.Name ?? "";
                            var combinedName = $"{familyName} {typeName}".ToUpperInvariant();
                            
                            // ✅ EXCLUDE: Skip VCD and VOLUME dampers (not in walls)
                            if (combinedName.Contains("VCD") || combinedName.Contains("VOLUME"))
                                return false;
                            
                            return familyName.Contains("Damper");
                        })
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
                return UnitConverter.FromInternalFeet(distance);
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

            // ✅ ALWAYS LOG: Always log damper intersection processing (not just in diagnostic mode)
            log($"[DamperIntersection] Processing damper {damperElement.Id} with bbox Min=({hostDamperBBox.Min.X:F2}, {hostDamperBBox.Min.Y:F2}, {hostDamperBBox.Min.Z:F2}) Max=({hostDamperBBox.Max.X:F2}, {hostDamperBBox.Max.Y:F2}, {hostDamperBBox.Max.Z:F2})");
            log($"[DamperIntersection] Testing against {structuralElements.Count} structural elements");

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

                    bool bboxesIntersect = BoundingBoxService.BoundingBoxesIntersect(expandedMin, expandedMax, structBBox.Min, structBBox.Max);
                    if (!bboxesIntersect)
                    {
                        // ✅ LOG FIRST FEW: Log first few non-intersections to debug (limit to avoid spam)
                        if (results.Count == 0 && structuralElements.IndexOf(tuple) < 3)
                        {
                            log($"[DamperIntersection] ⚠️ No bbox intersection: damper {damperElement.Id} (expanded: Min=({expandedMin.X:F2},{expandedMin.Y:F2},{expandedMin.Z:F2}) Max=({expandedMax.X:F2},{expandedMax.Y:F2},{expandedMax.Z:F2})) vs structural {structuralElement.Id} (Min=({structBBox.Min.X:F2},{structBBox.Min.Y:F2},{structBBox.Min.Z:F2}) Max=({structBBox.Max.X:F2},{structBBox.Max.Y:F2},{structBBox.Max.Z:F2}))");
                        }
                        continue;
                    }

                    log($"[DamperIntersection] ✅ Intersection candidate: damper {damperElement.Id} with structural {structuralElement.Id}");

                    // ✅ FIX: Check if damper is completely contained within wall (buried inside)
                    // This handles the case where damper is fully inside the wall geometry
                    bool isDamperContainedInWall = 
                        hostDamperBBox.Min.X >= structBBox.Min.X &&
                        hostDamperBBox.Min.Y >= structBBox.Min.Y &&
                        hostDamperBBox.Min.Z >= structBBox.Min.Z &&
                        hostDamperBBox.Max.X <= structBBox.Max.X &&
                        hostDamperBBox.Max.Y <= structBBox.Max.Y &&
                        hostDamperBBox.Max.Z <= structBBox.Max.Z;

                    // ✅ ALWAYS LOG: Log containment check (not just in diagnostic mode)
                    if (isDamperContainedInWall)
                    {
                        log($"[DamperIntersection] ✅ Damper {damperElement.Id} is COMPLETELY CONTAINED within structural {structuralElement.Id} (buried inside wall)");
                    }

                    var intersectionMin = new XYZ(
                        Math.Max(hostDamperBBox.Min.X, structBBox.Min.X),
                        Math.Max(hostDamperBBox.Min.Y, structBBox.Min.Y),
                        Math.Max(hostDamperBBox.Min.Z, structBBox.Min.Z));
                    var intersectionMax = new XYZ(
                        Math.Min(hostDamperBBox.Max.X, structBBox.Max.X),
                        Math.Min(hostDamperBBox.Max.Y, structBBox.Max.Y),
                        Math.Min(hostDamperBBox.Max.Z, structBBox.Max.Z));

                    // ✅ FIX: If damper is contained, use damper's bounding box as intersection
                    // Otherwise, check if intersection bounding box is valid
                    if (isDamperContainedInWall)
                    {
                        // Damper is completely inside wall - use damper's bbox as intersection
                        intersectionMin = hostDamperBBox.Min;
                        intersectionMax = hostDamperBBox.Max;
                    }
                    else if (intersectionMin.X > intersectionMax.X ||
                             intersectionMin.Y > intersectionMax.Y ||
                             intersectionMin.Z > intersectionMax.Z)
                    {
                        // No valid intersection bounding box
                        if (OptimizationFlags.UseDiagnosticMode)
                        {
                            log($"[DamperIntersection] ⚠️ No valid intersection bbox: damper {damperElement.Id} vs structural {structuralElement.Id} (Min=({intersectionMin.X:F3},{intersectionMin.Y:F3},{intersectionMin.Z:F3}) Max=({intersectionMax.X:F3},{intersectionMax.Y:F3},{intersectionMax.Z:F3}))");
                        }
                        continue;
                    }

                    var intersectionBBox = new BoundingBoxXYZ
                    {
                        Min = intersectionMin,
                        Max = intersectionMax
                    };

                    // ✅ FIX: For dampers, use damper's insertion point (LocationPoint) as center - this is the geometric center of damper body
                    // This excludes connectors and ensures placement point is at center of damper's width and height
                    // Fallback to bounding box center if LocationPoint is not available
                    XYZ damperCenter;
                    if (damperElement is FamilyInstance familyInstance)
                    {
                        // Use insertion point (LocationPoint) - this is the geometric center of the damper body, excluding connectors
                        var locationPoint = familyInstance.Location as LocationPoint;
                        if (locationPoint != null)
                        {
                            damperCenter = locationPoint.Point;
                            // Transform if damper is in a linked document
                            if (damperLinkTransform != null)
                            {
                                damperCenter = damperLinkTransform.OfPoint(damperCenter);
                            }
                        }
                        else
                        {
                            // Fallback to transform origin (geometric center)
                            var transform = familyInstance.GetTransform();
                            damperCenter = transform.Origin;
                            if (damperLinkTransform != null)
                            {
                                damperCenter = damperLinkTransform.OfPoint(damperCenter);
                            }
                        }
                    }
                    else
                    {
                        // Fallback to bounding box center if not a FamilyInstance
                        damperCenter = BoundingBoxService.GetBoundingBoxCenter(hostDamperBBox);
                    }
                    
                    // Project damper center onto wall plane (structural element face)
                    // For walls, project along wall normal; for floors/framing, use intersection bbox center as fallback
                    XYZ intersectionPoint;
                    if (structuralElement is Wall wall)
                    {
                        // Get wall normal and project damper center onto wall face
                        var wallNormal = wall.Orientation;
                        if (linkTransform != null)
                        {
                            wallNormal = linkTransform.OfVector(wallNormal);
                        }
                        
                        // Get wall face origin (use wall location curve start point)
                        var wallLocation = wall.Location as LocationCurve;
                        XYZ wallFaceOrigin = wallLocation?.Curve?.GetEndPoint(0) ?? damperCenter;
                        if (linkTransform != null && wallFaceOrigin != null)
                        {
                            wallFaceOrigin = linkTransform.OfPoint(wallFaceOrigin);
                        }
                        
                        // Project damper center onto wall plane
                        double distance = (damperCenter - wallFaceOrigin).DotProduct(wallNormal);
                        intersectionPoint = damperCenter - wallNormal.Multiply(distance);
                        
                        if (OptimizationFlags.UseDiagnosticMode)
                            log($"[DamperIntersection] Using damper insertion point ({damperCenter.X:F3}, {damperCenter.Y:F3}, {damperCenter.Z:F3}) projected onto wall = ({intersectionPoint.X:F3}, {intersectionPoint.Y:F3}, {intersectionPoint.Z:F3})");
                    }
                    else
                    {
                        // For floors/framing, use intersection bbox center (fallback)
                        intersectionPoint = BoundingBoxService.GetBoundingBoxCenter(intersectionBBox);
                        
                        if (OptimizationFlags.UseDiagnosticMode)
                            log($"[DamperIntersection] Using intersection bbox center for non-wall element: ({intersectionPoint.X:F3}, {intersectionPoint.Y:F3}, {intersectionPoint.Z:F3})");
                    }
                    
                    results.Add((structuralElement, intersectionBBox, intersectionPoint));
                }
                catch (Exception ex)
                {
                    log($"ERROR: Failed to process damper intersection for element {structuralElement.Id}: {ex.Message}");
                }
            }

            return results;
        }

        // ✅ OOP REFACTORING: Removed duplicate TransformBoundingBox - now uses BoundingBoxService.TransformBoundingBox()

        /// <summary>
        /// Gets the centerline of an MEP element and indicates whether it's from the fallback path.
        /// Returns (line, isFallbackLine) where isFallbackLine=true means line is already in host coordinates.
        /// </summary>
        private static (Line? line, bool isFallbackLine) GetElementLineWithSource(Element element, BoundingBoxXYZ mepBBox, Action<string> log)
        {
            if (element is FamilyInstance fi && fi.Symbol?.Family?.Name?.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                log?.Invoke($"[MepIntersectionService] Element {element.Id} identified as damper; using bounding-box intersection approach.");
                return (null, false);
            }

            if (element.Location is LocationCurve locCurve && locCurve.Curve is Line curveLine)
            {
                if (OptimizationFlags.UseDiagnosticMode)
                    log?.Invoke($"[GetElementLine] Using LocationCurve path for element {element.Id} - line is in '{element.Document.Title}' coordinates (needs transform to host)");
                return (curveLine, false); // LocationCurve is in linked doc coordinates, needs transform
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
                            if (OptimizationFlags.UseDiagnosticMode)
                                log?.Invoke($"[GetElementLine] Using Connector path for element {element.Id} - line is in '{element.Document.Title}' coordinates (needs transform to host)");
                            return (Line.CreateBound(endpoints.First.Origin, endpoints.Second.Origin), false); // Connector is in linked doc coordinates, needs transform
                        }
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[MepIntersectionService] Failed deriving line from MEPCurve connectors for element {element.Id}: {ex.Message}");
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
                    log?.Invoke($"[GetElementLine] Using Fallback path for element {element.Id} - line created from transformed bbox (already in host coordinates, no transform needed)");
                return (Line.CreateBound(p1, p2), true); // Fallback line is created from transformed bbox, already in host coordinates
            }
            catch (Exception ex)
            {
                log?.Invoke($"[MepIntersectionService] Failed to derive fallback line for element {element.Id}: {ex.Message}");
                return (null, false);
            }
        }
        
        // Legacy method for compatibility with other code paths
        private static Line? GetElementLine(Element element, BoundingBoxXYZ mepBBox, Action<string> log)
        {
            return GetElementLineWithSource(element, mepBBox, log).line;
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
                
                double minThicknessInternal = UnitConverter.ToInternalMillimeters(minThicknessMm);
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
                                double wallThicknessMm = UnitConverter.FromInternalMillimeters(wallThickness);
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
                
                double minThicknessInternal = UnitConverter.ToInternalMillimeters(minThicknessMm);
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
                                double wallThicknessMm = UnitConverter.FromInternalMillimeters(wallThickness);
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

        /// <summary>
        /// Helper method to get intersection point from Revit's IntersectionResult (handles both Point and XYZPoint properties)
        /// </summary>
        private static XYZ GetIntersectionPointFromRevitResult(Autodesk.Revit.DB.IntersectionResult ir)
        {
            try
            {
                // Try Point property first (Revit 2024+)
                var pointProperty = ir.GetType().GetProperty("Point");
                if (pointProperty != null)
                    return (XYZ)pointProperty.GetValue(ir);
            }
            catch { }
            
            try
            {
                // Try XYZPoint property (Revit 2020-2023)
                var xyzPointProperty = ir.GetType().GetProperty("XYZPoint");
                if (xyzPointProperty != null)
                    return (XYZ)xyzPointProperty.GetValue(ir);
            }
            catch { }
            
            throw new InvalidOperationException("Unable to get intersection point from IntersectionResult");
        }
    }
}
#endif // !REVIT2024_OR_GREATER

#if REVIT2024_OR_GREATER
// ========================================================================================================
// REVIT 2024+ MINIMAL IMPLEMENTATION
// This version eliminates ALL static caches to avoid TypeInitializationException in R2024 environment
// ========================================================================================================
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static partial class MepIntersectionService
    {
        // NO STATIC CACHES - avoid TypeInitializationException
        private const double MillimetersPerFoot = 304.8;
        
        // Simple instance-based conversion (no static initialization)
        private static double ToInternalMillimeters(double value) => value / MillimetersPerFoot;
        private static double FromInternalMillimeters(double value) => value * MillimetersPerFoot;
        
        // Category whitelists (primitives only - safe in static context)
        private static readonly BuiltInCategory[] MEP_CATEGORY_WHITELIST = {
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
        
        private static readonly BuiltInCategory[] STRUCTURAL_CATEGORY_WHITELIST = {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFoundation
        };
        
        // ✅ R24: FindIntersectionsBatch implementation (Fixed with Face.Intersect)
        // ✅ CRITICAL: NO STATIC CACHES - removed to avoid TypeInitializationException
        // Geometry is computed on-demand without caching

        public static List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsBatch(
            List<(Element, Transform?)> mepElements,
            List<(Element, Transform?)> structuralElements,
            Action<string> log,
            HashSet<(int mepId, int structuralId)>? knownValidPairs = null,
            bool skipKnownPairsGeometryCheck = false)
        {
            var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
            
            try
            {
                log?.Invoke($"[R24-FIX] FindIntersectionsBatch called: {mepElements.Count} MEP + {structuralElements.Count} structural");
                
                foreach (var (mepElem, mepTransform) in mepElements)
                {
                    try
                    {
                        // Get MEP Line (Centerline)
                        var mepLine = GetElementLine(mepElem);
                        if (mepLine == null) continue;
                        
                        // Transform MEP line if needed
                        if (mepTransform != null && !mepTransform.IsIdentity)
                        {
                            mepLine = mepLine.CreateTransformed(mepTransform) as Line;
                        }
                        
                        if (mepLine == null) continue;

                        var mepBBox = mepElem.get_BoundingBox(null);
                        
                        foreach (var (structElem, structTransform) in structuralElements)
                        {
                            try
                            {
                                // Spatial Filter (Bounding Box)
                                var structBBox = structElem.get_BoundingBox(null);
                                if (structBBox != null && mepBBox != null)
                                {
                                    // Transform structural bbox if needed
                                    if (structTransform != null && !structTransform.IsIdentity)
                                    {
                                        var min = structTransform.OfPoint(structBBox.Min);
                                        var max = structTransform.OfPoint(structBBox.Max);
                                        structBBox = new BoundingBoxXYZ
                                        {
                                            Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                                            Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
                                        };
                                    }
                                    
                                    // Simple overlap check
                                    if (!BoundingBoxesOverlap(mepBBox, structBBox)) continue;
                                }

                                // ✅ Use EfficientIntersectionService for solid intersection
                                // This reuses the robust logic we already have
                                var points = EfficientIntersectionService.PerformSolidIntersection(mepLine, structElem, structTransform);
                                
                                // Create Result
                                if (points.Count > 0)
                                {
                                    var bbox = CreateBoundingBox(points);
                                    if (bbox != null)
                                    {
                                        var center = (bbox.Min + bbox.Max) * 0.5;
                                        results.Add((mepElem, structElem, bbox, center));
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // log?.Invoke($"[R24-FIX] Struct element {structElem.Id} failed: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log?.Invoke($"[R24-FIX] MEP element {mepElem.Id} failed: {ex.Message}");
                    }
                }
                
                log?.Invoke($"[R24-FIX] Found {results.Count} intersections");
                return results;
            }
            catch (Exception ex)
            {
                log?.Invoke($"[R24-FIX] FATAL ERROR: {ex.Message}");
                return results;
            }
        }

        // --- Helper Methods ---

        private static Line GetElementLine(Element element)
        {
            if (element.Location is LocationCurve lc && lc.Curve is Line line)
            {
                return line;
            }
            return null;
        }

        private static bool BoundingBoxesOverlap(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
        {
            return bb1.Max.X >= bb2.Min.X && bb1.Min.X <= bb2.Max.X &&
                   bb1.Max.Y >= bb2.Min.Y && bb1.Min.Y <= bb2.Max.Y &&
                   bb1.Max.Z >= bb2.Min.Z && bb1.Min.Z <= bb2.Max.Z;
        }

        private static List<Solid> GetSolidsFromGeometry(Element element)
        {
            var solids = new List<Solid>();
            var options = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true, IncludeNonVisibleObjects = true };
            var geomElem = element.get_Geometry(options);
            if (geomElem == null) return solids;

            foreach (var obj in geomElem)
            {
                if (obj is Solid s && s.Volume > 0) solids.Add(s);
                else if (obj is GeometryInstance gi)
                {
                    foreach (var instObj in gi.GetInstanceGeometry())
                    {
                        if (instObj is Solid instS && instS.Volume > 0) solids.Add(instS);
                    }
                }
            }
            return solids;
        }

        private static List<XYZ> GetIntersectionPoints(Solid solid, Line line)
        {
            var points = new List<XYZ>();
            foreach (Face face in solid.Faces)
            {
                var res = face.Intersect(line, out IntersectionResultArray ira);
                if (res == SetComparisonResult.Overlap && ira != null)
                {
                    foreach (IntersectionResult ir in ira)
                    {
                        points.Add(ir.XYZPoint);
                    }
                }
            }
            return points;
        }

        private static BoundingBoxXYZ CreateBoundingBox(List<XYZ> points)
        {
            if (points == null || points.Count == 0) return null;
            double minX = points.Min(p => p.X);
            double minY = points.Min(p => p.Y);
            double minZ = points.Min(p => p.Z);
            double maxX = points.Max(p => p.X);
            double maxY = points.Max(p => p.Y);
            double maxZ = points.Max(p => p.Z);
            return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
        }

        // Stub methods to maintain API compatibility
        public static bool IsMepCategoryWhitelisted(Element element) => 
            MEP_CATEGORY_WHITELIST.Contains((BuiltInCategory)element.Category.Id.IntegerValue);
        
        public static bool IsStructuralCategoryWhitelisted(Element element) => 
            STRUCTURAL_CATEGORY_WHITELIST.Contains((BuiltInCategory)element.Category.Id.IntegerValue);
        
        // ✅ R2024 FIX: No static caches - these methods are no-ops
        public static void ClearGeometryCache() 
        {
            // No-op: No static cache in R2024 to clear
        }
        
        public static void ClearTransformCache() 
        { 
            // No-op: No static cache in R24
        }
        
        public static (int count, int maxSize, double memoryEstimateMB) GetGeometryCacheStats() 
        {
            // No static cache in R2024 - return zeros
            return (0, 0, 0.0);
        }
        
        public static Transform GetCachedTransform(Document doc, List<RevitLinkInstance> links, Action<string>? log = null)
        {
            try
            {
                var link = links?.FirstOrDefault(l => l.GetLinkDocument()?.Title == doc.Title);
                return link?.GetTotalTransform() ?? Transform.Identity;
            }
            catch (TypeInitializationException tiex)
            {
                // ✅ R2024 FIX: Handle TypeInitializationException from static initialization
                log?.Invoke($"[R24-FIX] TypeInitializationException in GetCachedTransform: {tiex.Message}, Inner: {tiex.InnerException?.Message ?? "None"}");
                // Fallback: Direct transform without cache
                var link = links?.FirstOrDefault(l => l.GetLinkDocument()?.Title == doc.Title);
                return link?.GetTotalTransform() ?? Transform.Identity;
            }
        }
        
        // Legacy single-element methods
        public static List<(Element structuralElement, BoundingBoxXYZ bbox, XYZ center)> FindIntersections(
            Element mepElement,
            List<(Element, Transform?)> structuralElements,
            Transform? mepTransform,
            Action<string>? log = null)
        {
            var results = new List<(Element, BoundingBoxXYZ, XYZ)>();
            var mepList = new List<(Element, Transform?)> { (mepElement, mepTransform) };
            var intersections = FindIntersectionsBatch(mepList, structuralElements, log ?? (_ => { }));
            foreach (var intersection in intersections)
            {
                results.Add((intersection.Item2, intersection.Item3, intersection.Item4));
            }
            return results;
        }
        
        public static List<(Element structuralElement, BoundingBoxXYZ bbox, XYZ center)> FindIntersections(
            Line mepLine,
            BoundingBoxXYZ mepBBox,
            List<(Element, Transform?)> structuralElements,
            Action<string>? log = null)
        {
            // R24 minimal: Line-based intersection not fully supported in this minimal version
            return new List<(Element, BoundingBoxXYZ, XYZ)>();
        }
        
        public static List<(Element structuralElement, BoundingBoxXYZ bbox, XYZ center)> FindDamperIntersections(
            Element damper,
            List<(Element, Transform?)> structuralElements,
            Transform? damperTransform,
            Action<string>? log = null)
        {
            return FindIntersections(damper, structuralElements, damperTransform, log);
        }
        
        public static List<Element> CollectStructuralElementsForDirectIntersectionVisibleOnly(
            Document doc,
            Action<string>? log = null)
        {
            var results = new List<Element>();
            try
            {
                var collector = new FilteredElementCollector(doc);
                foreach (var cat in STRUCTURAL_CATEGORY_WHITELIST)
                {
                    var elems = collector.OfCategory(cat).WhereElementIsNotElementType().ToElements();
                    results.AddRange(elems);
                }
            }
            catch (Exception ex)
            {
                log?.Invoke($"[R24] CollectStructural failed: {ex.Message}");
            }
            return results;
        }
    }
}
#endif // REVIT2024_OR_GREATER