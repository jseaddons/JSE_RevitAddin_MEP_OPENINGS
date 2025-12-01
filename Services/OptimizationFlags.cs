using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Feature flags for safe rollout of optimization features
    /// Provides risk mitigation and gradual deployment capability
    /// </summary>
    public static class OptimizationFlags
    {
        // Use minimal per-category global index during Refresh to avoid loading full XMLs
        public static bool UseGlobalCategoryIndexForRefresh { get; set; } = true;
        #region Phase 1 Foundation Flags (40% gain)
        
        /// <summary>
        /// Enable geometry caching for performance improvement
        /// Default: true (safe to enable)
        /// </summary>
        public static bool UseGeometryCache { get; set; } = true;
        
        /// <summary>
        /// Enable memory management with LRU eviction
        /// Default: true (safe to enable)
        /// </summary>
        public static bool UseMemoryManagement { get; set; } = true;
        
        /// <summary>
        /// Enable smart tolerance handling based on element size
        /// Default: true (safe to enable)
        /// </summary>
        public static bool UseSmartTolerance { get; set; } = true;
        
        /// <summary>
        /// Enable cache invalidation strategy
        /// Default: true (safe to enable)
        /// </summary>
        public static bool UseCacheInvalidation { get; set; } = true;
        
        #endregion
        
        #region Phase 2 Advanced Flags (45% gain)
        
        /// <summary>
        /// Enable R-tree spatial filtering
        /// Default: true (already implemented)
        /// </summary>
        public static bool UseRTreeFilter { get; set; } = true;
        
        /// <summary>
        /// Enable parallel processing for intersection testing
        /// Default: false (experimental - requires testing)
        /// </summary>
        public static bool UseParallelProcessing { get; set; } = true; // ✅ Enabled by default for cluster processing
        
        /// <summary>
        /// Enable two-tier spatial index (grid + R-tree)
        /// Default: true (enabled for testing - fixed intersection point issue)
        /// </summary>
        public static bool UseSpatialGrid { get; set; } = true;
        
        /// <summary>
        /// Enable R-tree spatial indexing in SQLite database for section box filtering
        /// When true: Uses R-tree virtual table for O(log n) spatial queries
        /// When false: Falls back to B-tree indexes with in-memory filtering
        /// Default: true (enabled - R-tree is supported in SQLite 3.42.0+)
        /// Expected gain: 10x faster section box filtering, 80-90% reduction in data transfer
        /// </summary>
        public static bool UseRTreeDatabaseIndex { get; set; } = true;
        
        #endregion
        
        #region Section Box & Filtering Optimizations (NEW - Priority 1)
        
        /// <summary>
        /// Use BoundingBoxIntersectsFilter instead of ElementIntersectsSolidFilter for section box filtering
        /// When true: Uses fast bounding box filter (20-30% faster)
        /// When false: Uses existing solid filter (slower but more precise)
        /// Default: false (disabled initially - enable after validation)
        /// Location: Helpers/SectionBoxHelper.cs
        /// </summary>
        public static bool UseBoundingBoxSectionBoxFilter { get; set; } = true;
        
        /// <summary>
        /// Re-enable TestCurveInBoundingBox filter for cheap rejection before solid intersection
        /// When true: Skips expensive solid extraction for non-intersecting curves (10-15% faster)
        /// When false: Always performs solid intersection (slower but more reliable)
        /// Default: false (disabled - causes inconsistent results with linked files)
        /// ⚠️ KEEP DISABLED: Edge cases with coordinate transforms cause false rejections
        /// Location: Services/MepIntersectionService.cs (line 757)
        /// </summary>
        public static bool UseCurveInBoundingBoxFilter { get; set; } = false;
        
        /// <summary>
        /// Use WhereElementIsViewIndependent() in FilteredElementCollector to skip view-dependent filtering
        /// When true: Faster element collection (5-10% faster) if view visibility not needed
        /// When false: Standard collector behavior (view-dependent filtering)
        /// Default: false (disabled initially - enable if view visibility not required)
        /// Location: Services/MepIntersectionService.cs, Helpers/MepElementCollectorHelper.cs
        /// </summary>
        public static bool UseViewIndependentCollector { get; set; } = true;
        
        #endregion
        
        #region Spatial & Geometry Optimizations (NEW - Priority 2)
        
        /// <summary>
        /// Use level-based spatial grid instead of 3D grid for better locality
        /// When true: Separate spatial grids per level (15-20% faster for multi-level projects)
        /// When false: Uses existing 3D spatial grid
        /// Default: false (disabled initially - enable after validation)
        /// Location: Services/SpatialPartitioningService.cs
        /// </summary>
        public static bool UseLevelBasedSpatialGrid { get; set; } = true;
        
        /// <summary>
        /// Migrate geometry cache to support List&lt;Solid&gt; for compound walls
        /// When true: Caches all solids for compound walls (eliminates recomputation)
        /// When false: Uses existing single Solid cache (may recompute for compound walls)
        /// Default: true (enabled - already partially implemented in R2024 path)
        /// Location: Services/MepIntersectionService.cs
        /// </summary>
        public static bool UseMultiSolidCache { get; set; } = true;
        
        #endregion
        
        #region Advanced Optimizations (NEW - Priority 3, Optional)
        
        /// <summary>
        /// Use progressive LOD (Level of Detail) pipeline: Outline → Curve → Solid
        /// When true: Tiered filtering (LOD0 outline, LOD1 curve, LOD2 solid) - 20-30% faster for large datasets
        /// When false: Single-pass filtering (current behavior)
        /// Default: false (disabled - high complexity, low priority)
        /// Location: Services/MepIntersectionService.cs
        /// </summary>
        public static bool UseProgressiveLOD { get; set; } = false;
        
        /// <summary>
        /// Use hybrid spatial index: R-tree for linear elements, grid for volumes
        /// When true: Selects optimal index per element type (10-15% faster for mixed types)
        /// When false: Uses existing spatial grid for all elements
        /// Default: false (disabled - high complexity, low priority)
        /// Location: Services/MepIntersectionService.cs
        /// </summary>
        public static bool UseHybridSpatialIndex { get; set; } = false;
        
        /// <summary>
        /// Enable performance metrics logging for diagnostics
        /// When true: Logs filter reduction ratios, geometry extraction times, transaction durations
        /// When false: No performance logging (deployment mode)
        /// Default: false (disabled for deployment - enable for diagnostics)
        /// </summary>
        public static bool LogPerformanceMetrics { get; set; } = false;
        
        #endregion
        
        #region Phase 3 Intelligence Flags (10% + 90% incremental)
        
        /// <summary>
        /// Enable incremental detection for real-time updates
        /// Default: false (experimental - requires testing)
        /// </summary>
        public static bool UseIncrementalDetection { get; set; } = false;
        
        /// <summary>
        /// Enable diagnostic mode for performance monitoring
        /// Default: false (disabled for deployment)
        /// </summary>
        public static bool UseDiagnosticMode { get; set; } = true; // ✅ DEBUG: Diagnostic mode ON - detailed performance logging enabled
        
        #endregion
        
        #region Sleeve Placement Safety Flags (NEW)
        
        /// <summary>
        /// Enable optimized XML batch saves (single write instead of per-sleeve)
        /// Default: true (safe to enable - already implemented)
        /// </summary>
        public static bool UseOptimizedXmlSaves { get; set; } = true;
        
        /// <summary>
        /// Enable XML save validation with backup/restore
        /// Default: true (safe to enable - critical for data integrity)
        /// </summary>
        public static bool UseXmlValidation { get; set; } = true;
        
        /// <summary>
        /// Enable pre-cached family symbols (load once, reuse many times)
        /// Default: true (safe to enable - significant performance gain)
        /// </summary>
        public static bool UseFamilySymbolCache { get; set; } = true;
        
        /// <summary>
        /// Enable incremental cache updates instead of full rebuilds
        /// Default: true (safe to enable - already partially implemented)
        /// </summary>
        public static bool UseIncrementalCache { get; set; } = true;
        
        /// <summary>
        /// Enable parallel processing for pre-filtering eligible clash zones (XML-only validation)
        /// Default: true (safe to enable - multi-threading optimization for flag checking)
        /// </summary>
        public static bool UseParallelPreFiltering { get; set; } = true;
        
        /// <summary>
        /// Enable multi-threading for non-Revit API operations in UniversalClusterService
        /// Operations: XML loading, parameter extraction, cache population
        /// Default: true (safe to enable - file I/O and data processing only, no Revit API calls)
        /// </summary>
        public static bool UseClusterServiceMultiThreading { get; set; } = true;
        
        /// <summary>
        /// Enable parallel clearance calculation for individual sleeve placement (non-Revit operation)
        /// Pre-calculates clearance values before placement loop to save time
        /// Default: true (safe to enable - pure math, no Revit API calls)
        /// </summary>
        public static bool UseParallelClearanceCalculation { get; set; } = true;

        /// <summary>
        /// Enable optimized single-write geometry and per-parameter timing instrumentation during sleeve placement.
        /// When true: Uses SetSleeveParametersOptimized path with timing logs to param_timing.log.
        /// Default: false (safe off; turn on for diagnostics only).
        /// </summary>
        public static bool EnableParameterTimingInstrumentation { get; set; } = true;
        
        /// <summary>
        /// Defer non-critical metadata writes during sleeve placement (Phase 2: Medium Risk).
        /// When true: Only writes critical parameters (MEP_Category, MEP_ElementId, ClashZone_GUID, Sleeve Instance ID, Filter Name) during placement.
        /// Non-critical parameters (MEP_UniqueId, MEP_Size, System_Abbreviation, MEP_Count, Bottom of Opening, Host Parameters) are deferred to batch write.
        /// Critical parameters are required for flag reset logic during refresh.
        /// Default: true (enabled - safe with batch writer integration).
        /// Expected gain: 150-180ms per sleeve (70-80% of metadata time).
        /// </summary>
        public static bool DeferNonCriticalMetadata { get; set; } = true;
        
        /// <summary>
        /// Batch parameter writes until after document regeneration (Step 5: High Impact Optimization).
        /// When true: Accumulates all parameter values during placement loop, regenerates once, then writes all parameters.
        /// When false: Writes parameters immediately during placement (current behavior).
        /// Default: true (enabled - safe with fallback to immediate writes on error).
        /// Expected gain: 4-6× faster individual placement (143-203ms → <30ms per sleeve).
        /// Location: Services/UniversalSleevePlacerService.cs, Services/OpeningCommandOrchestrator.cs
        /// ✅ VERIFIED: Set to true (2025-11-24) - Individual sleeve parameter batching enabled
        /// </summary>
        public static bool UseBatchedParameterWrites { get; set; } = true;

        /// <summary>
        /// Enable snapshot parameter transfer from SQLite SleeveSnapshots into placed sleeve instances after pipeline placement.
        /// When true: After placement pipeline finishes, retrieves snapshot MEP parameters and defers (or immediately writes) them.
        /// When false: Skips snapshot transfer entirely.
        /// Default: true (safe - transfer only uses existing parameters).
        /// Location: Services/Placement/SleevePlacementOrchestrator.cs (post-pipeline section)
        /// </summary>
        public static bool EnableSnapshotParameterTransfer { get; set; } = true;
        
        #region Refactoring Flags (Phase 1 - Safe Rollout)
        
        /// <summary>
        /// Enable new SleeveRepository for data persistence (XML/DB)
        /// When true: Uses extracted SleeveRepository service
        /// When false: Uses legacy private methods in UniversalSleevePlacerService
        /// Default: false (disabled initially)
        /// </summary>
        public static bool UseNewSleeveRepository { get; set; } = false;

        /// <summary>
        /// Enable new ZoneFilterService for clash zone filtering
        /// When true: Uses extracted ZoneFilterService
        /// When false: Uses legacy private methods in UniversalSleevePlacerService
        /// Default: false (disabled initially)
        /// </summary>
        public static bool UseNewZoneFilter { get; set; } = false;

        /// <summary>
        /// Enable new FamilyManager for family loading and caching
        /// When true: Uses extracted FamilyManager service
        /// When false: Uses legacy private methods in UniversalSleevePlacerService
        /// Default: false (disabled initially)
        /// </summary>
        public static bool UseNewFamilyManager { get; set; } = false;

        /// <summary>
        /// Enable new NewSleevePlacerService as the main entry point
        /// When true: Uses NewSleevePlacerService
        /// When false: Uses legacy UniversalSleevePlacerService
        /// Default: false (disabled initially)
        /// </summary>
        public static bool UseNewSleevePlacerService { get; set; } = false;

        #endregion

        /// <summary>
        /// Enable parallel planning for pipes (experimental)
        /// When true: Pipes can use parallel planning optimization for pre-computing dimensions
        /// When false: Pipes use normal sequential processing (safe default)
        /// Default: false (disabled initially - enable after validation)
        /// Location: Services/UniversalSleevePlacerService.cs
        /// </summary>
        public static bool EnableParallelPlanningForPipes { get; set; } = false;

        #endregion
        
        #region Configuration Methods
        
        /// <summary>
        /// Load optimization flags from configuration file
        /// Falls back to safe defaults if configuration is missing
        /// </summary>
        public static void LoadFromConfiguration()
        {
            try
            {
                // For now, use safe defaults since ConfigurationManager is not available
                // In a full implementation, this would load from app.config or user settings
                
                // Load Phase 1 flags (safe defaults)
                UseGeometryCache = GetConfigValue("UseGeometryCache", true);
                UseMemoryManagement = GetConfigValue("UseMemoryManagement", true);
                UseSmartTolerance = GetConfigValue("UseSmartTolerance", true);
                UseCacheInvalidation = GetConfigValue("UseCacheInvalidation", true);
                
                // Load Phase 2 flags (experimental defaults)
                UseRTreeFilter = GetConfigValue("UseRTreeFilter", true);
                UseParallelProcessing = GetConfigValue("UseParallelProcessing", false);
                UseSpatialGrid = GetConfigValue("UseSpatialGrid", true); // ✅ PERFORMANCE FIX: Enable spatial grid by default (70-90% reduction in intersection tests)
                UseRTreeDatabaseIndex = GetConfigValue("UseRTreeDatabaseIndex", true); // ✅ R-TREE: Enable database R-tree by default
                
                // Load Phase 3 flags (experimental defaults)
                UseIncrementalDetection = GetConfigValue("UseIncrementalDetection", false);
                UseDiagnosticMode = GetConfigValue("UseDiagnosticMode", true);
                
                // Load Sleeve Placement flags (safe defaults)
                UseOptimizedXmlSaves = GetConfigValue("UseOptimizedXmlSaves", true);
                UseXmlValidation = GetConfigValue("UseXmlValidation", true);
                UseFamilySymbolCache = GetConfigValue("UseFamilySymbolCache", true);
                UseIncrementalCache = GetConfigValue("UseIncrementalCache", true);
                UseParallelPreFiltering = GetConfigValue("UseParallelPreFiltering", true);
                UseClusterServiceMultiThreading = GetConfigValue("UseClusterServiceMultiThreading", true);
                UseParallelClearanceCalculation = GetConfigValue("UseParallelClearanceCalculation", true);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[OptimizationFlags] Loaded configuration successfully");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[OptimizationFlags] Failed to load configuration, using safe defaults: {ex.Message}");

                // Use safe defaults defined above
            }
        }
        
        /// <summary>
        /// Save current optimization flags to configuration
        /// </summary>
        public static void SaveToConfiguration()
        {
            try
            {
                // This would save to user settings or config file
                // For now, just log the current state
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[OptimizationFlags] Current flags - GeometryCache: {UseGeometryCache}, MemoryManagement: {UseMemoryManagement}, SmartTolerance: {UseSmartTolerance}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[OptimizationFlags] Failed to save configuration: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get configuration value with fallback to default
        /// </summary>
        private static bool GetConfigValue(string key, bool defaultValue)
        {
            try
            {
                // For now, always return default values since ConfigurationManager is not available
                // In a full implementation, this would read from app.config or user settings
                return defaultValue;
            }
            catch
            {
                // Ignore configuration errors
                return defaultValue;
            }
        }
        
        /// <summary>
        /// Reset all flags to safe defaults
        /// </summary>
        public static void ResetToSafeDefaults()
        {
            // Phase 1: Safe to enable
            UseGeometryCache = true;
            UseMemoryManagement = true;
            UseSmartTolerance = true;
            UseCacheInvalidation = true;
            
            // Phase 2: Conservative defaults
            UseRTreeFilter = true;
            UseParallelProcessing = false;
            UseSpatialGrid = false;
            UseRTreeDatabaseIndex = false; // ✅ SAFETY: Disable R-tree by default in safe mode (fallback to B-tree)
            
            // Phase 3: Conservative defaults
            UseIncrementalDetection = false;
            UseDiagnosticMode = false;
            
            // Sleeve Placement: Safe defaults (all enabled)
            UseOptimizedXmlSaves = true;
            UseXmlValidation = true;
            UseFamilySymbolCache = true;
            UseIncrementalCache = true;
            UseParallelPreFiltering = true;
            UseClusterServiceMultiThreading = true;
            
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[OptimizationFlags] Reset to safe defaults");
        }
        
        /// <summary>
        /// Enable all Phase 1 optimizations (40% gain)
        /// </summary>
        public static void EnablePhase1Optimizations()
        {
            UseGeometryCache = true;
            UseMemoryManagement = true;
            UseSmartTolerance = true;
            UseCacheInvalidation = true;
            
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[OptimizationFlags] Enabled Phase 1 optimizations (40% gain)");
        }
        
        /// <summary>
        /// Enable all Phase 2 optimizations (45% gain)
        /// </summary>
        public static void EnablePhase2Optimizations()
        {
            UseRTreeFilter = true;
            UseParallelProcessing = true;
            UseSpatialGrid = true;
            UseRTreeDatabaseIndex = true;
            
                        if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[OptimizationFlags] Enabled Phase 2 optimizations (45% gain)");
        }
        
        /// <summary>
        /// Get current optimization status for logging
        /// </summary>
        public static string GetOptimizationStatus()
        {
            return $@"Optimization Flags Status:
Phase 1 (40% gain): GeometryCache={UseGeometryCache}, MemoryManagement={UseMemoryManagement}, SmartTolerance={UseSmartTolerance}, CacheInvalidation={UseCacheInvalidation}
Phase 2 (45% gain): RTreeFilter={UseRTreeFilter}, ParallelProcessing={UseParallelProcessing}, SpatialGrid={UseSpatialGrid}, RTreeDatabaseIndex={UseRTreeDatabaseIndex}
Phase 3 (10% + 90% incremental): IncrementalDetection={UseIncrementalDetection}, DiagnosticMode={UseDiagnosticMode}
Section Box Optimizations: BoundingBoxSectionBoxFilter={UseBoundingBoxSectionBoxFilter}, CurveInBoundingBoxFilter={UseCurveInBoundingBoxFilter}, ViewIndependentCollector={UseViewIndependentCollector}
Spatial Optimizations: LevelBasedSpatialGrid={UseLevelBasedSpatialGrid}, MultiSolidCache={UseMultiSolidCache}
Advanced Optimizations: ProgressiveLOD={UseProgressiveLOD}, HybridSpatialIndex={UseHybridSpatialIndex}, LogPerformanceMetrics={LogPerformanceMetrics}
Refactoring Flags: SleeveRepository={UseNewSleeveRepository}, ZoneFilter={UseNewZoneFilter}, FamilyManager={UseNewFamilyManager}";
        }
        
        #endregion
    }
}
