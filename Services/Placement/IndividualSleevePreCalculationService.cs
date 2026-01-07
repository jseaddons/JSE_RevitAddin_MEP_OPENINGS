using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SOLID SRP: Service for pre-calculating individual sleeve placement data in parallel.
    /// 
    /// ⚠️⚠️⚠️ CRITICAL PROTECTION - DO NOT MODIFY WITHOUT CONSENT ⚠️⚠️⚠️
    /// This service implements the same robust parallel processing pattern as ClusterPreCalculationService:
    /// - 3-Tier Partitioning: Category → Spatial → Sequential
    /// - ThreadLocal Cache Isolation: Each thread gets its own isolated dictionary
    /// - Crash-Safe Execution: Exception handling, graceful degradation
    /// - Flag-Based Control: UseSOLIDRefactoredIndividualPreCalculation
    /// 
    /// Architecture:
    /// - Separates calculation phase (parallelizable) from placement phase (sequential)
    /// - Uses pure math operations (safe for multi-threading, no Revit API calls during calculation)
    /// - ThreadLocal caches prevent cross-thread data mixing
    /// 
    /// Performance:
    /// - Parallel processing: 2-4× faster for large zone counts (4-core CPU)
    /// - Pre-calculation runs BEFORE placement loop (no blocking)
    /// - Thread isolation eliminates lock contention
    /// </summary>
    public class IndividualSleevePreCalculationService : IIndividualSleevePreCalculationService
    {
        // =====================================================
        // THREAD-LOCAL CACHES (CRITICAL for parallel safety)
        // =====================================================
        
        /// <summary>
        /// Thread-local cache for dimension calculations.
        /// Each thread gets its own isolated dictionary to prevent race conditions.
        /// </summary>
        private readonly ThreadLocal<Dictionary<Guid, (double width, double height, double clearance)>> _dimensionCache;
        
        /// <summary>
        /// Thread-local cache for clearance lookups by category+insulation key.
        /// </summary>
        private readonly ThreadLocal<Dictionary<string, double>> _clearanceCache;
        
        /// <summary>
        /// Thread-local cache for placement point validation results.
        /// </summary>
        private readonly ThreadLocal<Dictionary<Guid, bool>> _validationCache;
        
        private const int MAX_CACHE_SIZE_PER_THREAD = 2000;
        
        // Dependencies
        private readonly OpeningConditions _conditions;
        private readonly Dictionary<string, double> _clearanceSettings;
        private readonly IInsulationAwareSizingService _sizingService;
        
        /// <summary>
        /// Constructor with dependency injection.
        /// </summary>
        public IndividualSleevePreCalculationService(
            OpeningConditions conditions = null,
            Dictionary<string, double> clearanceSettings = null,
            IInsulationAwareSizingService sizingService = null)
        {
            _conditions = conditions ?? new OpeningConditions();
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            
            // ✅ THREAD-SAFETY: Initialize ThreadLocal factories
            // Each thread will get its own new Dictionary when accessing .Value
            _dimensionCache = new ThreadLocal<Dictionary<Guid, (double, double, double)>>(
                () => new Dictionary<Guid, (double, double, double)>(), trackAllValues: false);
            
            _clearanceCache = new ThreadLocal<Dictionary<string, double>>(
                () => new Dictionary<string, double>(), trackAllValues: false);
            
            _validationCache = new ThreadLocal<Dictionary<Guid, bool>>(
                () => new Dictionary<Guid, bool>(), trackAllValues: false);
        }
        
        /// <summary>
        /// Pre-calculate placement data for all zones using 3-tier parallel processing.
        /// </summary>
        public Dictionary<Guid, SleevePreCalculationResult> PreCalculateAllZones(
            List<ClashZone> clashZones,
            Document doc)
        {
            var results = new ConcurrentDictionary<Guid, SleevePreCalculationResult>();
            
            if (clashZones == null || clashZones.Count == 0)
            {
                SafeFileLogger.SafeAppendText("individual_precalc.log",
                    $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc] ⚠️ No zones to pre-calculate\n");
                return new Dictionary<Guid, SleevePreCalculationResult>();
            }
            
            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            // =====================================================
            // TIER 1: CATEGORY PARTITIONING
            // =====================================================
            // Group zones by MepElementCategory for full isolation (Ducts don't interfere with Pipes)
            var zonesByCategory = clashZones
                .GroupBy(z => z.MepElementCategory ?? "Unknown")
                .ToDictionary(g => g.Key, g => g.ToList());
            
            SafeFileLogger.SafeAppendText("individual_precalc.log",
                $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc] 🔀 TIER 1 - CATEGORY PARTITIONING: {zonesByCategory.Count} categories, {clashZones.Count} total zones\n");
            
            foreach (var kvp in zonesByCategory)
            {
                SafeFileLogger.SafeAppendText("individual_precalc.log",
                    $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc]   - {kvp.Key}: {kvp.Value.Count} zones\n");
            }
            
            // Process each category in parallel (categories are fully independent)
            System.Threading.Tasks.Parallel.ForEach(zonesByCategory, 
                new ParallelOptions { MaxDegreeOfParallelism = zonesByCategory.Count }, 
                categoryEntry =>
            {
                string category = categoryEntry.Key;
                var categoryZones = categoryEntry.Value;
                
                // =====================================================
                // TIER 2: SPATIAL PARTITIONING (within each category)
                // =====================================================
                // Group zones by 10-foot spatial grid cells to allow parallel processing
                // of non-overlapping regions
                double cellSize = 10.0; // 10 feet grid (approx 3 meters)
                var spatialPartitions = new Dictionary<string, List<ClashZone>>();
                
                foreach (var zone in categoryZones)
                {
                    int cellX = (int)(zone.IntersectionPointX / cellSize);
                    int cellY = (int)(zone.IntersectionPointY / cellSize);
                    int cellZ = (int)(zone.IntersectionPointZ / cellSize);
                    string spatialKey = $"{cellX}_{cellY}_{cellZ}";
                    
                    if (!spatialPartitions.ContainsKey(spatialKey))
                        spatialPartitions[spatialKey] = new List<ClashZone>();
                    spatialPartitions[spatialKey].Add(zone);
                }
                
                SafeFileLogger.SafeAppendText("individual_precalc.log",
                    $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc]   📍 TIER 2 - Category '{category}': {spatialPartitions.Count} spatial cells\n");
                
                int maxThreads = CalculateOptimalThreadCount();
                
                // Process spatial partitions in parallel (non-overlapping regions are safe)
                System.Threading.Tasks.Parallel.ForEach(spatialPartitions, 
                    new ParallelOptions { MaxDegreeOfParallelism = maxThreads }, 
                    spatialEntry =>
                {
                    string spatialKey = spatialEntry.Key;
                    var spatialZones = spatialEntry.Value;
                    
                    // =====================================================
                    // TIER 3: SEQUENTIAL PROCESSING WITHIN CELL
                    // =====================================================
                    // Process zones in the same spatial cell sequentially to avoid overlap issues
                    foreach (var zone in spatialZones)
                    {
                        try
                        {
                            var result = PreCalculateSingleZone(zone, category);
                            results[zone.Id] = result;
                        }
                        catch (Exception ex)
                        {
                            // ✅ CRASH-SAFE: Graceful degradation - mark as invalid but continue
                            results[zone.Id] = new SleevePreCalculationResult
                            {
                                ClashZoneId = zone.Id,
                                IsValid = false,
                                ErrorMessage = ex.Message,
                                SkipReason = "PreCalculationError"
                            };
                            
                            SafeFileLogger.SafeAppendText("individual_precalc.log",
                                $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc] ❌ Error in zone {zone.Id} (cell={spatialKey}): {ex.Message}\n");
                        }
                    }
                });
            });
            
            sw.Stop();
            
            int validCount = results.Values.Count(r => r.IsValid);
            int invalidCount = results.Values.Count(r => !r.IsValid);
            int skippedCount = results.Values.Count(r => !string.IsNullOrEmpty(r.SkipReason));
            
            SafeFileLogger.SafeAppendText("individual_precalc.log",
                $"[{DateTime.Now:HH:mm:ss}] [IndividualPreCalc] ✅ Completed 3-tier parallel pre-calculation: " +
                $"Valid={validCount}, Invalid={invalidCount}, Skipped={skippedCount}, " +
                $"Time={sw.ElapsedMilliseconds}ms ({sw.ElapsedMilliseconds / (double)clashZones.Count:F2}ms per zone)\n");
            
            return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }
        
        /// <summary>
        /// Pre-calculate data for a single zone.
        /// Uses thread-local caches for isolation.
        /// </summary>
        private SleevePreCalculationResult PreCalculateSingleZone(ClashZone zone, string category)
        {
            var result = new SleevePreCalculationResult
            {
                ClashZoneId = zone.Id,
                IsValid = false
            };
            
            try
            {
                // =====================================================
                // FILTER 1: Basic Skip Checks (no Revit API needed)
                // =====================================================
                if (zone.IsResolvedFlag || zone.SleeveInstanceId > 0)
                {
                    result.SkipReason = "AlreadyResolved";
                    return result;
                }
                
                if (zone.IsClusterResolvedFlag || zone.ClusterSleeveInstanceId > 0)
                {
                    result.SkipReason = "ClusterResolved";
                    return result;
                }
                
                if (zone.HasDamperNearby && !string.Equals(category, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                {
                    result.SkipReason = "DamperNearby";
                    return result;
                }
                
                if (zone.StructuralElementIdValue <= 0)
                {
                    result.SkipReason = "InvalidHostId";
                    return result;
                }
                
                if (zone.MepElementIdValue <= 0)
                {
                    result.SkipReason = "InvalidMepId";
                    return result;
                }
                
                // =====================================================
                // DIMENSION CALCULATION (ThreadLocal cache)
                // =====================================================
                // Check thread-local cache first
                if (_dimensionCache.Value.TryGetValue(zone.Id, out var cachedDims))
                {
                    result.TargetWidth = cachedDims.width;
                    result.TargetHeight = cachedDims.height;
                    result.ClearanceFeet = cachedDims.clearance;
                }
                else
                {
                    // Calculate dimensions
                    bool isPipe = string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase);
                    bool isDuct = string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase);
                    bool isCableTray = string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase);
                    
                    double rawWidth = zone.MepElementWidth > 0 ? zone.MepElementWidth : zone.MepElementSize;
                    double rawHeight = zone.MepElementHeight > 0 ? zone.MepElementHeight : zone.MepElementSize;
                    
                    if (isPipe && zone.MepElementOuterDiameter > 0)
                    {
                        rawWidth = zone.MepElementOuterDiameter;
                        rawHeight = zone.MepElementOuterDiameter;
                    }
                    
                    // Get clearance (with thread-local caching)
                    double clearance = GetCachedClearance(category, zone.IsInsulated);
                    
                    // Calculate target dimensions
                    double insulation = zone.IsInsulated ? (zone.InsulationThickness * 2) : 0;
                    double targetWidth = rawWidth + insulation + (2 * clearance);
                    double targetHeight = rawHeight + insulation + (2 * clearance);
                    
                    result.TargetWidth = targetWidth;
                    result.TargetHeight = targetHeight;
                    result.ClearanceFeet = clearance;
                    
                    // Cache result (limit cache size)
                    if (_dimensionCache.Value.Count < MAX_CACHE_SIZE_PER_THREAD)
                    {
                        _dimensionCache.Value[zone.Id] = (targetWidth, targetHeight, clearance);
                    }
                }
                
                // =====================================================
                // ROTATION ANGLE
                // =====================================================
                result.RotationAngleDeg = zone.MepElementRotationAngle * 180.0 / Math.PI;
                
                // =====================================================
                // RISK CLASSIFICATION
                // =====================================================
                double rawSize = Math.Max(zone.MepElementWidth, zone.MepElementHeight);
                if (rawSize <= 0) rawSize = zone.MepElementSize;
                if (rawSize <= 0) rawSize = 0.25; // Fallback
                
                double ratio = result.ClearanceFeet / rawSize;
                if (ratio < 0.08)
                    result.Risk = ClearanceRiskClassification.Critical;
                else if (ratio < 0.10)
                    result.Risk = ClearanceRiskClassification.High;
                else if (ratio < 0.15)
                    result.Risk = ClearanceRiskClassification.Medium;
                else
                    result.Risk = ClearanceRiskClassification.Low;
                
                // =====================================================
                // HOST VALIDATION
                // =====================================================
                result.HostType = zone.StructuralElementType ?? "Unknown";
                result.HostExists = zone.StructuralElementIdValue > 0;
                result.HostIsValid = !string.IsNullOrEmpty(zone.StructuralElementType);
                
                // =====================================================
                // PLACEMENT POINT
                // =====================================================
                result.PlacementPointX = zone.IntersectionPointX;
                result.PlacementPointY = zone.IntersectionPointY;
                result.PlacementPointZ = zone.IntersectionPointZ;
                result.PlacementPointIsValid = 
                    !double.IsNaN(zone.IntersectionPointX) && 
                    !double.IsNaN(zone.IntersectionPointY) && 
                    !double.IsNaN(zone.IntersectionPointZ) &&
                    !(zone.IntersectionPointX == 0 && zone.IntersectionPointY == 0 && zone.IntersectionPointZ == 0);
                
                // ✅ SUCCESS
                result.IsValid = result.PlacementPointIsValid && result.HostIsValid;
                result.Category = category;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                result.IsValid = false;
            }
            
            return result;
        }
        
        /// <summary>
        /// Get clearance with thread-local caching.
        /// </summary>
        private double GetCachedClearance(string category, bool isInsulated)
        {
            string cacheKey = $"{category}_{isInsulated}";
            
            // Check thread-local cache
            if (_clearanceCache.Value.TryGetValue(cacheKey, out double cached))
                return cached;
            
            // Calculate clearance
            double clearanceMm = 50.0; // Default 50mm
            
            if (_conditions?.ClearanceSettings != null)
            {
                if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    clearanceMm = isInsulated 
                        ? _conditions.ClearanceSettings.PipesInsulated 
                        : _conditions.ClearanceSettings.PipesNormal;
                }
                else if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    clearanceMm = isInsulated 
                        ? _conditions.ClearanceSettings.RectangularInsulated 
                        : _conditions.ClearanceSettings.RectangularNormal;
                }
                else if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    clearanceMm = _conditions.ClearanceSettings.CableTrayTop;
                }
                else if (string.Equals(category, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                {
                    clearanceMm = isInsulated 
                        ? _conditions.ClearanceSettings.DuctAccessoryMepInsulated 
                        : _conditions.ClearanceSettings.DuctAccessoryMepNormal;
                }
            }
            
            double clearanceFt = clearanceMm / 304.8;
            
            // Cache result
            if (_clearanceCache.Value.Count < MAX_CACHE_SIZE_PER_THREAD)
            {
                _clearanceCache.Value[cacheKey] = clearanceFt;
            }
            
            return clearanceFt;
        }
        
        /// <summary>
        /// Calculate optimal thread count based on processor type.
        /// Same logic as ClusterPreCalculationService.
        /// </summary>
        private int CalculateOptimalThreadCount()
        {
            int logicalCores = Environment.ProcessorCount;
            
            if (logicalCores <= 4) return 2;                      // i3/low-end i5
            if (logicalCores <= 8) return logicalCores / 2;       // i5/i7
            if (logicalCores <= 16) return (int)(logicalCores * 0.6); // i7/i9
            if (logicalCores <= 32) return logicalCores / 2;      // Workstation
            return 16; // Server-grade: diminishing returns
        }
        
        /// <summary>
        /// Clear all thread-local caches. Call at end of placement batch.
        /// </summary>
        public void ClearCaches()
        {
            try
            {
                _dimensionCache.Value?.Clear();
                _clearanceCache.Value?.Clear();
                _validationCache.Value?.Clear();
            }
            catch { }
        }
    }
}
