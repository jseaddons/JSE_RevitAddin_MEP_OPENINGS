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
        /// <summary>
        /// Enable sequential workflow orchestration (Placement → Global Sync → Clustering).
        /// When true: Processes all filters for individual placement first, then global corner extraction, then all clustering.
        /// When false: Uses interleaved processing per filter (legacy behavior).
        /// Default: false (safe rollout)
        /// </summary>
        public static bool UseSequentialOrchestrator { get; set; } = true;

        /// <summary>
        /// Enable refactored parameter transfer optimizations (buffered logging, lazy cache, unified lookup, struct-based deferred params).
        /// When true: Uses new optimized code paths in ParameterTransferService and SleeveParameterService.
        /// When false: Uses legacy code paths.
        /// Default: false (safe rollout, enable for testing).
        /// Location: Services/ParameterTransferService.cs, Services/Placement/SleeveParameterService.cs
        /// </summary>
        public static bool UseParameterTransferRefactor { get; set; } = true;
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
        /// Default: false (disabled - caused missing linked-host clashes due to coordinate mismatch)
        /// Note: R-tree in link coordinates rejected valid clashes in host coordinates
        /// When disabled: Falls back to spatial grid results which correctly handles transforms
        /// </summary>
        public static bool UseRTreeFilter { get; set; } = false;
        
        /// <summary>
        /// Enable parallel processing for intersection testing
        /// Default: false (experimental - requires testing)
        /// </summary>
        public static bool UseParallelProcessing { get; set; } = true; // General parallel flag

        /// <summary>
        /// Enable Parallel Broad Phase Search for Clash Detection.
        /// When true: Uses multi-threading for BoundingBox overlap checks (Safe).
        /// When false: Uses sequential loop.
        /// Default: true (Safe to enable - no API calls in parallel part).
        /// </summary>
        public static bool UseParallelClashSearch { get; set; } = true;
        
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
        public static bool UseRTreeDatabaseIndex { get; set; } = true; // ✅ FIXED: R-tree synchronization now handled in bulk saves
        
        #endregion
        
        #region Section Box & Filtering Optimizations (NEW - Priority 1)

        /// <summary>
        /// Use cached section box from database instead of querying Revit API every time.
        /// When true: Uses SectionBoxService.GetSectionBoxBounds(dbConnection) for filtering.
        /// When false: Uses live SectionBoxHelper.GetSectionBoxBounds(view3D).
        /// Default: false (safe rollout)
        /// </summary>
        public static bool UseSectionBoxCache { get; set; } = true;
        
        /// <summary>
        /// Use BoundingBoxIntersectsFilter instead of ElementIntersectsSolidFilter for section box filtering
        /// When true: Uses fast bounding box filter (20-30% faster)
        /// When false: Uses existing solid filter (slower but more precise)
        /// Default: false (disabled initially - enable after validation)
        /// Location: Helpers/SectionBoxHelper.cs
        /// </summary>
        public static bool UseBoundingBoxSectionBoxFilter { get; set; } = false; // ⚠️ DISABLED: Causes inverted section box filtering (places sleeves outside instead of inside)
        
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
        /// Use ReferenceIntersector (Ray Casting) for linear elements (Ducts, Pipes, etc.)
        /// When true: Uses Revit's fast ray casting engine (Main Thread) instead of slow solid intersection
        /// When false: Uses parallel solid intersection (Legacy)
        /// Default: true (Safe, native method)
        /// </summary>
        public static bool UseRayCastingForLinearElements { get; set; } = true;

        /// <summary>
        /// Use Hybrid detection: Only use Ray Casting for DIAGONAL elements.
        /// When true: Logic checks if element is diagonal. If yes -> Ray Cast. If no -> Parallel Solid (or whatever is default).
        /// When false: Uses Ray Casting for ALL linear elements if UseRayCastingForLinearElements is true.
        /// Default: false (Optimized: Use Ray Cast for ALL linear elements per user request)
        /// </summary>
        public static bool UseHybridDiagonalDetection { get; set; } = false;
        
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

        /// <summary>
        /// Use stable geometry cache keys that do NOT depend on Transform.GetHashCode()
        /// When true: Cache keys derived from element id + linked/active tag (prevents hash churn)
        /// When false: Falls back to legacy cache key (may cause repeated geometry extraction)
        /// Default: true (safe – deterministic, no functional impact)
        /// Location: Services/MepIntersectionService.cs (FindIntersectionsBatchInternal)
        /// </summary>
        public static bool UseStableGeometryCacheKeys { get; set; } = true;

        /// <summary>
        /// Precompute and cache ALL host (structural) solids once per intersection run
        /// When true: Loads & transforms solids up-front (best for small host sets <= 200)
        /// When false: Lazy loads per candidate (current behavior)
        /// Default: false (DISABLED - causes massive slowdown on large linked projects)
        /// Location: Services/MepIntersectionService.cs (FindIntersectionsBatchInternal)
        /// </summary>
        public static bool PrecomputeHostSolids { get; set; } = false;

        /// <summary>
        /// Log aggregated geometry extraction metrics (total ms, cache hits/misses) at end of batch
        /// When true: Adds diagnostic summary lines (negligible overhead)
        /// When false: Suppresses extra performance logging
        /// Default: true (useful during optimization phase; disable for production noise reduction)
        /// Location: Services/MepIntersectionService.cs (end of FindIntersectionsBatchInternal)
        /// </summary>
        public static bool LogGeometryExtractionMetrics { get; set; } =false;

        /// <summary>
        /// Enable detailed timing diagnostics for Flag Reset operation during refresh
        /// When true: Logs breakdown of DB query, Revit API checks, and batch update times
        /// When false: Standard logging only
        /// Default: false (no extra logging; set true only when debugging flag reset timing)
        /// Location: Services/FlagManager_Legacy.cs (ResetFlagsForDeletedSleeves) — not currently referenced
        /// </summary>
        public static bool LogFlagResetDiagnostics { get; set; } = false;

        public static bool EnablePlacementDebugLog { get; set; } = UseDiagnosticMode;

        /// <summary>
        /// Enable detailed Bulk Placement diagnostics (individual sleeves).
        /// When true: Logs placement details to bulk_placement_debug.log.
        /// Default: false (skipped for performance).
        /// </summary>
        public static bool EnableBulkPlacementDebugLog { get; set; } = true;

        /// <summary>
        /// Enable detailed Batch Clustering diagnostics (Step 4 & 5).
        /// When true: Logs batch placement details to batch_v2.log.
        /// Default: false (skipped for performance).
        /// </summary>
        public static bool EnableBatchClusteringDebugLog { get; set; } = true;

        /// <summary>
        /// Enable detailed Cluster Sizing diagnostics (RCS dimensions, swaps).
        /// When true: Logs sizing details to cluster_debug.log and placement_sizing_debug.log.
        /// Default: false (skipped for performance).
        /// </summary>
        public static bool EnableClusterSizingDebugLog { get; set; } = true;

        /// <summary>
        /// Enable detailed Proximity Check diagnostics (Clustering workflow).
        /// When true: Logs proximity analysis details to flag_workflow.log.
        /// Default: false (skipped for performance).
        /// </summary>
        public static bool EnableProximityDebugLog { get; set; } = true;

        /// <summary>
        /// Enable detailed placement diagnostics (e.g. corner comparisons).
        /// When true: Logs calculated vs Revit-extracted corners for clusters.
        /// Default: false (skipped for performance).
        /// </summary>
        public static bool EnablePlacementDiagnostics { get; set; } = true;

        /// <summary>
        /// Use streamlined clash zone creation path when intersections come from optimized MepIntersectionService
        /// When true: Skips redundant validation, bounding box queries, and complex penetration calculations (10-20x faster clash zone creation)
        /// When false: Uses legacy full validation path
        /// Default: true (safe - intersections from MepIntersectionService are already validated)
        /// Impact: Reduces clash zone creation from 189ms to ~10ms per zone (19x faster)
        /// Location: Services/ClashZoneService_Legacy.cs (DetectNewClashZones)
        /// </summary>
        public static bool UseStreamlinedClashZoneCreation { get; set; } = true;

        /// <summary>
        /// ✅ PERFORMANCE: Bulk orientation pre-calculation for clash zones
        /// When true: Pre-calculates ALL MEP element orientations before zone creation loop, caches in dictionary
        /// When false: Calculates orientation for each zone individually (redundant for same MEP crossing multiple walls)
        /// Default: true (enabled - 26 zones is correct count)
        /// Impact: Reduces orientation calculation from O(N zones) to O(N unique MEPs), ~50% savings
        /// Location: Services/ClashZoneService.cs (DetectNewClashZonesStreamlined)
        /// </summary>
        public static bool UseBulkOrientationCaching { get; set; } = true;

        /// <summary>
        /// 🔍 DIAGNOSTIC: Enable detailed line-by-line profiling in CreateClashZone()
        /// When true: Logs timing for each operation in CreateClashZone to identify bottlenecks
        /// When false: No profiling overhead
        /// Default: true (enabled for performance analysis)
        /// Impact: Adds logging overhead but reveals where time is spent
        /// Location: Services/ClashZoneService.cs (CreateClashZone)
        /// </summary>
        public static bool EnableDetailedClashZoneProfiler { get; set; } = false;

        /// <summary>
        /// Skip XML processing during Flag Reset operation (database-only mode optimization)
        /// When true: Only updates SQLite database, skips reading/writing Global XML files (5-10x faster flag reset)
        /// When false: Updates both database AND XML files (legacy behavior)
        /// Default: true (safe - database is primary source, XML is deprecated backup)
        /// Impact: Reduces flag reset from ~450ms to ~50ms overhead (9x faster)
        /// Location: Services/FlagManager_Legacy.cs (ResetFlagsForDeletedSleeves)
        /// </summary>
        public static bool SkipXmlDuringFlagReset { get; set; } = true;

        /// <summary>
        /// Skip XML file writes during Save operation (database-only mode optimization)
        /// When true: Only writes to SQLite database, skips creating XML files (3-5x faster save)
        /// When false: Writes to both database AND XML files (legacy behavior)
        /// Default: true (safe - database is primary source, XML is deprecated backup)
        /// Impact: Reduces save time from ~494ms to ~100ms (5x faster)
        /// Location: Services/ClashZonePersistenceService.cs (SaveClashZones, SaveCategory)
        /// </summary>
        public static bool SkipXmlDuringSave { get; set; } = true;

        /// <summary>
        /// Use bulk UPDATE statement for SQLite clash zone saves instead of individual row updates
        /// When true: Batches all clash zone updates into single SQL transaction (10-20x faster database writes)
        /// When false: Updates each clash zone row individually with separate queries (legacy behavior)
        /// Default: true (safe - uses parameterized queries with transaction, maintains data integrity)
        /// Impact: Reduces SQLite write time from ~517ms to ~30ms for 9 zones (17x faster)
        /// Location: Data/ClashZoneRepository.cs (InsertOrUpdateClashZones)
        /// </summary>
        public static bool UseBulkSqliteUpdates { get; set; } = true;

        /// <summary>
        /// Perform SQLite schema/R-tree verification only once per session.
        /// When true: Skips repeated DB verification/logging on subsequent refreshes (safe if DB path is stable).
        /// When false: Verifies on every refresh (legacy behavior).
        /// Default: true (safe for deployed environments).
        /// Location: Data/ClashZoneRepository.cs (session guard around verification)
        /// </summary>
        public static bool UseOneTimeDbVerificationDuringSession { get; set; } = true;

        /// <summary>
        /// Use single-query batch UPDATE for SetReadyForPlacementFlag instead of memory filtering
        /// When true: Uses single SQL UPDATE with WHERE clause for section box + unresolved check (18x faster)
        /// When false: Loads zones into memory, filters in C#, then batch updates per filter/category (legacy behavior)
        /// Default: true (safe - uses same logic but in SQL instead of C#)
        /// Impact: Reduces SetReadyForPlacementFlag from ~367ms to ~20ms (18x faster)
        /// Location: Data/ClashZoneRepository.cs (SetReadyForPlacementForUnresolvedZonesInSectionBox)
        /// </summary>
        public static bool UseBatchReadyForPlacementUpdate { get; set; } = true;

        /// <summary>
        /// Use batch parameter capture for dampers instead of capturing per damper
        /// When true: Captures parameters for ALL dampers and walls BEFORE creating ClashZones (7x faster)
        /// When false: Captures parameters individually for each damper during ClashZone creation (legacy behavior)
        /// Default: true (safe - same parameters captured, just batched)
        /// Impact: Reduces damper ClashZone creation from ~687ms to ~100ms (7x faster)
        /// Location: Services/Refresh/DamperProcessingService.cs (ProcessDampers)
        /// </summary>
        public static bool UseBatchDamperParameterCapture { get; set; } = true;

        /// <summary>
        /// Skip XML reads during refresh (database-only mode).
        /// When true: Bypasses XML-CACHE loading and uses DB as the single source.
        /// When false: Loads XML cache for compatibility with older flows.
        /// Default: true (XML deprecated as backup).
        /// Location: Services/RefreshServiceRefactored.cs (PrepareExistingZones/XML-CACHE)
        /// </summary>
        public static bool SkipXmlLoadingDuringRefresh { get; set; } = true;

        /// <summary>
        /// Reuse a single SQLite DB context per refresh to avoid multiple opens.
        /// When true: Uses a shared SleeveDbContext for the refresh lifecycle, reducing connection and setup overhead.
        /// When false: Legacy behavior with multiple contexts created per phase.
        /// Default: true (safe, falls back instantly if disabled).
        /// Location: Services/FlagManager_Legacy.cs (context acquisition)
        /// </summary>
        public static bool ReuseDbContextDuringRefresh { get; set; } = true;

        /// <summary>
        /// Disable all verbose logging (keep only performance timing logs).
        /// When true: Silences all [SQLite], [INTERSECTION-PROCESSOR], [FLAG-MANAGER], etc. logs.
        /// When false: Full diagnostic logging enabled.
        /// Default: false (full logging for debugging).
        /// Use in production/deployment mode for maximum performance.
        /// Location: DebugLogger.cs (all Log calls check this flag)
        /// </summary>
        public static bool DisableVerboseLogging { get; set; } = false;// ✅ ENABLED for performance - reduces overhead

        /// <summary>
        /// Skip synchronous parameter capture for existing zones in ClashZoneService.
        /// When true: Parameters are only captured in the parallel Phase 8.
        /// When false: Sync capture during detection (legacy/slow).
        /// Default: true (Performance optimization)
        /// </summary>
        public static bool UseLazyParameterCapture { get; set; } = true;

        /// <summary>
        /// Enable high-performance Bulk Placement using NewFamilyInstances2
        /// Strictly Revit-only placement, uses legacy persistence architecture (SleeveParameterService).
        /// Location: Services/BulkPlacementService.cs
        /// Default: false (safe rollout)
        /// </summary>
        public static bool UseBulkIndividualSleevePlacement { get; set; } = true;

        /// <summary>
        /// Bulk placement transaction flow: two transactions (Monday-style) vs single transaction.
        /// When true: Tx1 = placement only → Commit; Tx2 = Flush parameters → Commit; then Regenerate (faster, ~25 sleeves/s).
        /// When false: Single tx = placement + Flush parameters → Commit (rollback path if two-tx has issues).
        /// Default: false (Single transaction avoids wasted regeneration of default-sized sleeves).
        /// Location: Services/OpeningCommandOrchestrator.cs (bulk placement block)
        /// </summary>
        public static bool UseTwoTransactionBulkPlacement { get; set; } = false;

        /// <summary>
        /// Skip forced garbage collection at end of refresh.
        /// When true: Lets CLR manage memory naturally (saves ~289ms).
        /// When false: Explicit GC.Collect(2) after refresh.
        /// Default: true (skip GC - modern .NET handles this efficiently).
        /// Tradeoff: Memory stays allocated slightly longer, but CLR optimizes collection timing.
        /// Location: refresh_service_refactored.cs FinalCleanup()
        /// </summary>
        public static bool SkipForcedGarbageCollection { get; set; } = true;

        /// <summary>
        /// Use single batched DB query to check sleeve existence instead of individual GetElement() calls.
        /// When true: Loads all zones for category from DB, checks existence in one pass (saves ~100-150ms).
        /// When false: Iterates through Global XML entries, checking each sleeve individually.
        /// Default: true (batched approach is more efficient).
        /// Location: FlagManager_Legacy.cs ProcessFlagResetForDeletedSleeves()
        /// </summary>
        public static bool UseBatchedSleeveExistenceCheck { get; set; } = true;

        /// <summary>
        /// Enable SOLID-compliant damper filter during refresh clash zone detection.
        /// When true: Uses ZoneFilterService to skip ducts with dampers nearby (two-tier flag-based caching).
        /// When false: Skips damper check entirely (legacy behavior, no filtering).
        /// Default: true (enables priority filtering for damper sleeves).
        /// Architecture: Implements OCP by delegating filtering to ZoneFilterService interface.
        /// Location: ClashZoneService_Legacy.cs (lines 520-580)
        /// Note: Uses cached HasDamperNearby flag + proximity check fallback.
        /// </summary>
        public static bool UseSOLIDCompliantDamperFilter { get; set; } = true;
        
        /// <summary>
        /// Enable SOLID-refactored cluster pre-calculation with parallel processing.
        /// When true: Uses IClusterPreCalculationService to pre-calculate rotation angles and bounding boxes
        ///            in parallel BEFORE the placement loop (2-4× faster for large clusters).
        /// When false: Uses legacy sequential calculation inside placement loop (current behavior).
        /// Default: true (enabled - SOLID refactoring with all 28 features preserved).
        /// Architecture: Implements SRP by separating calculation phase from placement phase.
        /// Performance: Pre-calculation runs in parallel (pure math, no Revit API calls).
        /// Location: Services/Clustering/RefactoredClusterService.cs
        /// Note: Falls back to legacy code automatically if pre-calculation fails (crash-safe).
        /// </summary>
        public static bool UseSOLIDRefactoredClusterPreCalculation { get; set; } = true;

        /// <summary>
        /// Enable "Bottom of Opening" parameter calculation for RectangularOpeningOnWall sleeves.
        /// When true: Calculates and sets "Bottom of Opening" = Schedule of Level - (Height / 2) for all rectangular wall sleeves.
        /// When false: Skips "Bottom of Opening" calculation (legacy behavior).
        /// Default: true (enabled - safe calculation with validation).
        /// Formula: Bottom of Opening = Schedule of Level - (Height / 2)
        /// Applies to: Individual sleeves and cluster sleeves (RectangularOpeningOnWall family only).
        /// Location: Services/Placement/SleeveParameterService.cs, Services/Clustering/Placement/ClusterPlacementService.cs
        /// Architecture: Uses BottomOfOpeningCalculationService helper (SRP-compliant, pure calculation).
        /// Features: Supports parameter batching, performance monitoring, safe validation, diagnostic logging.
        /// </summary>
        public static bool UseBottomOfOpeningCalculation { get; set; } = true;
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
        /// When falsimage.pnge: Uses existing spatial grid for all elements
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
        public static bool LogPerformanceMetrics { get; set; } = true;
        
        #endregion
        
        #region Phase 3 Intelligence Flags (10% + 90% incremental)
        
        /// <summary>
        /// Enable incremental detection for real-time updates
        /// Default: false (experimental - requires testing)
        /// </summary>
        public static bool UseIncrementalDetection { get; set; } = false;
        
        /// <summary>
        /// Enable diagnostic mode for performance monitoring
        /// Default: false (disabled - causes 2.6x slowdown due to logging overhead)
        /// </summary>
        public static bool UseDiagnosticMode { get; set; } = true; // false = no per-zone debug logging overhead
        
        /// <summary>
        /// Enable batch clash zone creation (pre-calculate common data once)
        /// Default: true (CRITICAL for performance - reduces 10.3s creation to ~2s)
        /// Location: Services/ClashZoneService.cs (DetectNewClashZones)
        /// </summary>
        public static bool UseBatchClashZoneCreation { get; set; } = true;

        #endregion
        
        #region Refresh Optimization Phase 3 (Batch GUID & Parameter Cache)

        /// <summary>
        /// Enable batch pre-fetching of deterministic GUIDs from database.
        /// When true: Fetches all required GUIDs in a single query before the loop (O(1) lookups).
        /// When false: Performs individual database lookups for each intersection (legacy behavior, O(n) lookups).
        /// Default: true (enables 90% faster GUID resolution).
        /// </summary>
        public static bool UseBatchGuidLookup { get; set; } = true;

        /// <summary>
        /// Enable pre-cached parameter whitelist for clash zone creation.
        /// When true: Builds the parameter whitelist once per refresh instead of per-zone.
        /// When false: Rebuilds whitelist for every single clash zone creation (legacy behavior).
        /// Default: true (reduces CPU overhead by redundant whitelist building).
        /// </summary>
        public static bool UsePreCachedWhitelist { get; set; } = true;

        /// <summary>
        /// Enable using a temporary table for bulk database updates.
        /// When true: Uploads changes to a temp table and performs a single JOIN-based UPDATE/INSERT.
        /// When false: Uses a single massive UPDATE statement with CASE WHEN clauses (higher memory/SQL overhead).
        /// Default: true (improves save performance for 1000+ zones).
        /// </summary>
        public static bool UseTempTableForBulkUpdates { get; set; } = true;

        #endregion

        #region Bulk Processing & DB Optimizations (NEW - Priority 4)

        /// <summary>
        /// Enable bulk intersection processing for Ducts, Pipes, and Cable Trays.
        /// When true: Batches parameter capture and creation for all intersection types (O(1) parameter access).
        /// When false: Legacy per-element processing (O(n) API calls).
        /// Default: true (enabled for performance testing).
        /// Target: 40-50 zones/sec.
        /// </summary>
        public static bool UseBulkIntersectionProcessing { get; set; } = true;

        /// <summary>
        /// Enable optimized database save strategy with pre-filtering and reduced grouping.
        /// When true: Loads existing IDs once, filters insert vs update, and streamlines category grouping (2x faster save).
        /// When false: Legacy save with repeated index loads and manual grouping.
        /// Default: true (enabled for performance testing).
        /// </summary>
        public static bool UseOptimizedDbSaveStrategy { get; set; } = true;

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
        public static bool EnableParameterTimingInstrumentation { get; set; } = false;
        
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
        /// Default: true (enabled - provides 4-6× performance improvement).
        /// Expected gain: 4-6× faster individual placement (143-203ms → <30ms per sleeve).
        /// Location: Services/UniversalSleevePlacerService.cs, Services/OpeningCommandOrchestrator.cs
        /// ✅ VERIFIED: Set to true (2025-11-24) - Individual sleeve parameter batching enabled
        /// ✅ BUG FIXED (2025-01-XX): Added GetParameterValueWithBatchingSupport() to read from deferred cache
        ///    before flushing, preventing stale reads of Width/Height/Depth during corner placement calculations.
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
        /// [OBSOLETE] Enable new SleeveRepository for data persistence (XML/DB)
        /// This flag is obsolete - the refactoring was never completed. The service exists as a stub but is not used.
        /// UseNewSleevePlacerService is the active refactored path instead.
        /// </summary>
        [Obsolete("This flag is obsolete - refactoring was never completed. UseNewSleevePlacerService is the active refactored path.")]
        public static bool UseNewSleeveRepository { get; set; } = false;

        /// <summary>
        /// [OBSOLETE] Enable new ZoneFilterService for clash zone filtering
        /// This flag is obsolete - the refactoring was never completed. The service exists but is not used.
        /// UseNewSleevePlacerService is the active refactored path instead.
        /// </summary>
        [Obsolete("This flag is obsolete - refactoring was never completed. UseNewSleevePlacerService is the active refactored path.")]
        public static bool UseNewZoneFilter { get; set; } = false;

        /// <summary>
        /// [OBSOLETE] Enable new FamilyManager for family loading and caching
        /// This flag is obsolete - the refactoring was never completed. The service was never implemented.
        /// UseNewSleevePlacerService is the active refactored path instead.
        /// </summary>
        [Obsolete("This flag is obsolete - refactoring was never completed. UseNewSleevePlacerService is the active refactored path.")]
        public static bool UseNewFamilyManager { get; set; } = false;

        /// <summary>
        /// Enable new NewSleevePlacerService as the main entry point
        /// When true: Uses NewSleevePlacerService
        /// When false: Uses legacy UniversalSleevePlacerService
        /// Default: true (✅ ENABLED FOR TESTING - SRP-compliant refactored service)
        /// </summary>
        public static bool UseNewSleevePlacerService { get; set; } = true;

        #region Cluster Sleeve Optimization Flags (NEW - High Impact)

        /// <summary>
        /// Use cluster family symbol caching to eliminate repeated family loading
        /// When true: Caches family symbols across all cluster placements
        /// When false: Loads family for each cluster (current behavior)
        /// Default: true (high impact optimization for cluster-heavy projects)
        /// Location: Services/Clustering/Placement/ClusterPlacementService.cs
        /// </summary>
        public static bool UseClusterFamilySymbolCaching { get; set; } = true;

        /// <summary>
        /// Use batch parameter operations for cluster sleeves
        /// When true: Sets all cluster parameters in batch operations
        /// When false: Sets parameters individually (current behavior)
        /// Default: true (significant performance improvement)
        /// Location: Services/Clustering/Placement/ClusterPlacementService.cs
        /// </summary>
        public static bool UseClusterBatchParameterOperations { get; set; } = true;

        /// <summary>
        /// Use parameter validation caching for cluster sleeves
        /// When true: Caches parameter validation results per family type
        /// When false: Validates parameters individually (current behavior)
        /// Default: true (moderate performance improvement)
        /// Location: Services/Clustering/Placement/ClusterPlacementService.cs
        /// </summary>
        public static bool UseClusterParameterValidationCaching { get; set; } = true;

        /// <summary>
        /// Use cluster family pre-loading for known family types
        /// When true: Pre-loads all required families before cluster placement
        /// When false: Loads families on-demand (current behavior)
        /// Default: true (high impact for projects with many clusters)
        /// Location: Services/Clustering/Placement/ClusterPlacementService.cs
        /// </summary>
        public static bool UseClusterFamilyPreLoading { get; set; } = true;

        /// <summary>
        /// Enable high-performance Bulk Placement for Cluster Sleeves using NewFamilyInstances2.
        /// When true: Uses doc.Create.NewFamilyInstances2 for clusters (similar to individual sleeves)
        /// When false: Places cluster sleeves one by one (current behavior)
        /// Default: true (high impact for cluster-heavy projects)
        /// Location: Services/Clustering/BatchClusterPlacementService.cs
        /// </summary>
        public static bool UseBulkClusterSleevePlacement { get; set; } = true;
        
        /// <summary>
        /// Enable optimized sleeve category filtering in RevitSleeveHelper.
        /// When true: Uses LINQ-based filtering over collected sleeve instances.
        /// When false: Uses the legacy foreach loop (identical behavior, slightly less efficient).
        /// Default: true (safe to enable; rollback by setting to false).
        /// Location: Services/Helpers/RevitSleeveHelper.cs
        /// </summary>
        public static bool UseOptimizedSleeveCategoryFilter { get; set; } = true;

        #endregion

        /// <summary>
        /// Enable SOLID-refactored command and orchestrator services
        /// When true: Uses extracted services (IConditionsLoader, IPathDeterminer, IStrategyFactory, etc.)
        /// When false: Uses legacy inline implementations
        /// Default: true (✅ ENABLED FOR TESTING - SOLID-compliant refactored services)
        /// Location: Commands/UniversalSleevePlacementCommand.cs, Services/OpeningCommandOrchestrator.cs
        /// </summary>
        public static bool UseRefactoredCommandServices { get; set; } = true;

        #endregion

        #region Crash-Safe Features (NewSleevePlacerService)

        /// <summary>
        /// Enable crash-safe execution with timeout protection for NewSleevePlacerService
        /// When true: Uses CrashSafeExecutor to prevent infinite hangs and ensure graceful failures
        /// When false: Uses normal execution (faster but less safe)
        /// Default: true (enabled for safety)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseCrashSafeExecution { get; set; } = true;

        #region Phase 1 Performance Optimizations (High Impact - Low Complexity)

        /// <summary>
        /// Skip redundant element validation during sleeve placement
        /// When true: Removes validation for elements just placed in same transaction (90% reduction in GetElement calls)
        /// When false: Validates every element (current behavior)
        /// Default: true (safe - elements just placed should be valid)
        /// Location: Services/NewSleevePlacerService.cs, Services/Placement/SleeveParameterService.cs
        /// </summary>
        public static bool SkipRedundantValidation { get; set; } = true;

        /// <summary>
        /// Use element location caching to avoid repeated LocationPoint/LocationCurve queries
        /// When true: Caches placement points during batch placement (70-80% reduction in location queries)
        /// When false: Queries location for every sleeve (current behavior)
        /// Default: true (safe optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseElementLocationCaching { get; set; } = true;

        /// <summary>
        /// Use level reference caching to avoid repeated level lookups
        /// When true: Caches level references and batches level parameter setting (80-90% reduction in level lookups)
        /// When false: Looks up level for every sleeve (current behavior)
        /// Default: true (safe optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseLevelReferenceCaching { get; set; } = true;

        /// <summary>
        /// Pre-calculate all dimensions before sleeve placement
        /// When true: Calculates all dimensions upfront before placement loop (eliminates repeated calculations)
        /// When false: Calculates dimensions during placement (current behavior)
        /// Default: true (safe optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UsePreCalculatedDimensions { get; set; } = true;

        /// <summary>
        /// Use batch parameter operations instead of individual parameter reads/writes
        /// When true: Batches parameter operations for better performance (50-60% reduction in parameter operations)
        /// When false: Individual parameter operations (current behavior)
        /// Default: true (safe optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseBatchParameterOperations { get; set; } = true;

        /// <summary>
        /// Use pre-caching of family symbols before placement (eliminates loading overhead)
        /// When true: Pre-loads and validates all required family symbols before placement loop
        /// When false: Loads symbols on-demand (current behavior with overhead)
        /// Default: true (high impact optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UsePreCachedFamilySymbols { get; set; } = true;

        /// <summary>
        /// Use memory leak detection and automatic garbage collection
        /// When true: Monitors memory usage and forces garbage collection to prevent leaks
        /// When false: Standard memory management (may accumulate memory)
        /// Default: true (stability optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseMemoryLeakDetection { get; set; } = true;

        /// <summary>
        /// Use operation variance reduction (warm-up and consistent data structures)
        /// When true: Pre-warms operations and uses consistent data structures to reduce variance
        /// When false: Standard operation execution (may have high variance)
        /// Default: true (consistency optimization)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseVarianceReduction { get; set; } = true;

        #endregion

        /// <summary>
        /// Enable safe element validation (avoids document mismatch bugs)
        /// When true: Validates elements by ElementId instead of Document reference (prevents false positives)
        /// When false: Uses direct element access (faster but less safe)
        /// Default: true (enabled - critical to avoid parameter transfer bug)
        /// Location: Services/NewSleevePlacerService.cs
        /// ⚠️ CRITICAL: This flag prevents the document mismatch bug we experienced in parameter transfer
        /// </summary>

        /// <summary>
        /// Enable safe transaction management (checks IsModifiable, handles rollback properly)
        /// When true: Validates document state before transactions, handles rollback safely
        /// When false: Assumes document is always modifiable (faster but less safe)
        /// Default: true (enabled for safety)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseSafeTransactionManagement { get; set; } = true;

        /// <summary>
        /// Enable performance monitoring for NewSleevePlacerService
        /// When true: Tracks timing, memory, and item counts for all operations
        /// When false: No performance tracking (faster but no metrics)
        /// Default: true (enabled for diagnostics)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UsePerformanceMonitoring { get; set; } = true;

        /// <summary>
        /// 🚀 MULTI-FLOOR OPTIMIZATION: Enable 3-phase global placement flow
        /// When true: 
        ///   Phase 1: Place ALL individual sleeves (all floors) → 1 Regenerate
        ///   Phase 2: Cluster ALL floors together → 1 Regenerate
        ///   Phase 3: Delete ALL old sleeves → 1 Regenerate
        ///   Result: Only 3 Regenerates total (vs 3N for N floors)
        /// When false: Legacy per-floor processing (3 regenerates per floor)
        /// Default: false (safe - use legacy until tested)
        /// ⚠️ ROLLBACK: Set to false to revert to working code instantly
        /// </summary>
        public static bool UseOptimizedMultiFloorFlow { get; set; } = false;

        /// <summary>
        /// Enable timeout protection for long-running operations
        /// When true: Operations timeout after 5 minutes to prevent infinite hangs
        /// When false: Operations run indefinitely (faster but can hang)
        /// Default: true (enabled for safety)
        /// Location: Services/NewSleevePlacerService.cs
        /// </summary>
        public static bool UseTimeoutProtection { get; set; } = true;
        
        #region Flag Management Refactoring
        
        /// <summary>
        /// Enable refactored SOLID-compliant flag management services
        /// When true: Uses FlagManagerService (bug-free, SOLID-compliant)
        /// When false: Uses legacy FlagManager (has bugs, but stable)
        /// Default: true (✅ ENABLED FOR TESTING - SOLID refactored services)
        /// Expected: Fixes sleeve duplication bugs, maintains all optimizations
        /// </summary>
        public static bool UseRefactoredClashZoneFlagServices { get; set; } = true; // ✅ ENABLED: Batch update now implemented (line 329 in FlagManagerService.cs) // \u26a0\ufe0f DISABLED: Batch update not implemented (line 296-297 in FlagManagerService.cs)
        
        #endregion

        #endregion

        /// <summary>
        /// [OBSOLETE] Enable parallel planning for pipes (experimental)
        /// This flag is obsolete and not useful - pipes are already included in parallel planning phase.
        /// The flag only controls whether pipes can USE pre-computed planning data, but all categories
        /// (including pipes) are already processed in parallel during planning. The flag adds unnecessary
        /// complexity without benefit. Parallel planning works for all categories regardless of this flag.
        /// </summary>
        [Obsolete("This flag is obsolete and not useful - pipes are already included in parallel planning. All categories use pre-computed planning data regardless of this flag.")]
        public static bool EnableParallelPlanningForPipes { get; set; } = false;

        #endregion

        #region Parameter Service Performance Optimizations (NEW - High Impact)

        /// <summary>
        /// Enable section box filtering for sleeves in parameter transfer dialog.
        /// When true: Only processes sleeves within the active 3D view's section box (80-95% reduction in processing time).
        /// When false: Processes all sleeves in the document (current behavior).
        /// Default: true (enabled - safe with fallback to all sleeves if section box unavailable).
        /// Location: Views/ParameterServiceDialogV2.cs
        /// Expected gain: 80-95% reduction in processing time for large projects.
        /// </summary>
        public static bool UseSectionBoxFilterForParameterTransfer { get; set; } = true;

        /// <summary>
        /// Enable skip logic for already-transferred parameters in configuration-based transfer.
        /// When true: Skips sleeves where target parameter already matches snapshot value (50-90% reduction on re-runs).
        /// When false: Processes every sleeve even if parameters already match (current behavior).
        /// Default: true (enabled - safe with value comparison).
        /// Location: Services/ParameterTransferService.cs
        /// Expected gain: 50-90% reduction in processing time on re-runs.
        /// </summary>
        public static bool SkipAlreadyTransferredParameters { get; set; } = true;

        /// <summary>
        /// Enable batch parameter lookups in ParameterTransferService.
        /// When true: Pre-caches all parameters for all sleeves before transfer loop (20-30% reduction in lookup overhead).
        /// When false: Uses individual LookupParameter() calls per sleeve (current behavior).
        /// Default: true (enabled - safe optimization).
        /// Location: Services/ParameterTransferService.cs
        /// Expected gain: 20-30% reduction in parameter lookup overhead.
        /// </summary>
        public static bool UseBatchParameterLookups { get; set; } = true;

        /// <summary>
        /// Enable batch parameter capture during refresh (ParameterCaptureService).
        /// When true: Collects all unique element IDs first, gets elements once and caches them, then processes parameters (30-50% faster).
        /// When false: Gets elements one-by-one per clash zone (current behavior).
        /// Default: true (enabled - validated and safe, preserves all parameters).
        /// Location: refresh refactor/parameter_capture_service.cs
        /// Expected gain: 30-50% reduction in parameter capture time (701ms → ~350-500ms).
        /// ⚠️ CRITICAL: Must preserve all parameters - validation ensures no data loss.
        /// </summary>
        public static bool UseBatchParameterCapture { get; set; } = true; // ✅ ENABLED: Batch parameter lookup optimization (re-enabled after fixing document validation bug)

        /// <summary>
        /// Enable bulk database operations for cluster save (ClusterSleeveRepository).
        /// When true: Uses bulk check query + bulk INSERT/UPDATE operations (90%+ faster).
        /// When false: Uses individual SELECT + INSERT/UPDATE per cluster (current behavior).
        /// Default: true (enabled - validated for data integrity).
        /// Location: Data/Repositories/ClusterSleeveRepository.cs (BatchSaveClusterSleeves)
        /// Expected gain: 90%+ reduction in database save time (7616ms → ~500ms).
        /// ✅ ROLLBACK: If issues occur, set this to false.
        /// </summary>
        public static bool UseBulkClusterSave { get; set; } = true;


        /// <summary>
        /// Enable database-only pre-calculation (avoids Revit API calls during parallel processing).
        /// When true: Uses database data (corners, placement points) instead of Revit API calls (50-70% faster).
        /// When false: Uses Revit API to get elements (current behavior, causes serialization).
        /// Default: false (disabled initially - enable after validation).
        /// Location: Services/Clustering/PreCalculation/ClusterPreCalculationService.cs
        /// Expected gain: 50-70% reduction in pre-calculation time (14803ms → ~6000-7500ms).
        /// ⚠️ CRITICAL: Requires all sleeve data (corners, placement points) to be in database.
        /// </summary>
        public static bool UseDatabaseOnlyPreCalculation { get; set; } = true; // ✅ ENABLED: Use database data instead of Revit API (50-70% faster)

        /// <summary>
        /// Limit parallel threads for cluster pre-calculation (reduces Revit API contention).
        /// When set: Limits MaxDegreeOfParallelism to this value (reduces thread contention).
        /// When 0: Auto-detects optimal thread count based on processor type.
        /// When -1: Uses Environment.ProcessorCount (all cores, may cause contention with Revit API).
        /// Default: 0 (auto-detect) - Recommended manual values:
        ///   - i5 (4-6 cores): 2-3 threads
        ///   - i7 (6-8 cores): 4-6 threads (or use all if database-only mode)
        ///   - i9 (8-16 cores): 6-8 threads (or use all if database-only mode)
        ///   - Xeon (16+ cores): 8-12 threads (or use all if database-only mode)
        /// Location: Services/Clustering/PreCalculation/ClusterPreCalculationService.cs
        /// Expected gain: 20-30% improvement on i5 processors by reducing thread contention.
        /// ⚠️ NOTE: With UseDatabaseOnlyPreCalculation=true, thread limiting is less critical (no Revit API contention).
        /// </summary>
        public static int ClusterPreCalculationMaxThreads { get; set; } = 0; // 0 = auto-detect, -1 = all cores, >0 = manual limit

        #region Combined Cluster Rollout (Phase 1-2)

        /// <summary>
        /// Master toggle for multi-category combined clustering rollout.
        /// When false: All combined clustering phases are bypassed.
        /// Default: false (opt-in while feature stabilizes).
        /// </summary>
        public static bool UseCombinedClustering { get; set; } = true;

        /// <summary>
        /// Enable Phase 1-2 (discovery and grouping) for combined clustering (DB/CPU-only).
        /// Default: false (isolated rollout without impacting placement).
        /// </summary>
        public static bool UseCombinedClusteringPhase1And2 { get; set; } = true;

        /// <summary>
        /// Placeholder toggle for Phase 3-4 (creation and persistence) of combined clustering.
        /// Default: false (kept off until downstream implementation is ready).
        /// </summary>
        public static bool UseCombinedClusteringPhase3And4 { get; set; } = true;

        /// <summary>
        /// Proximity tolerance (Revit internal units) used to merge cluster sleeves into combined candidates.
        /// Default: 0.5 (adjust per project scale; applied symmetrically to all axes).
        /// </summary>
        public static double CombinedClusterProximityTolerance { get; set; } = 0.5;

        /// <summary>
        /// Max degree of parallelism for combined clustering discovery/grouping.
        /// 0 = auto-detect (Environment.ProcessorCount - 1), -1 = all cores, >0 = explicit.
        /// </summary>
        public static int CombinedClusterMaxDegreeOfParallelism { get; set; } = 0;

        #endregion

        #endregion
        
        #region SOLID Refactoring Flags (NEW - Safe Rollback Strategy)
        
        /// <summary>
        /// Enable interface-based architecture for refresh services.
        /// When true: Uses IRefreshDataCacheManager, IRefreshPathDeterminer, IClashZoneRepository interfaces
        /// When false: Uses concrete classes directly (legacy behavior)
        /// Default: false (disabled - enable after Phase 1 testing)
        /// Location: refresh refactor/refresh_service_refactored.cs
        /// Expected benefit: Better testability, extensibility, SOLID compliance
        /// </summary>
        public static bool UseRefreshServiceInterfaces { get; set; } = true;

        /// <summary>
        /// Enable using specialized SOLID-refactored repositories (Combined, Cluster, Snapshot)
        /// via ClashZoneRepository delegation.
        /// When true: ClashZoneRepository delegates logic to specialized repos.
        /// When false: ClashZoneRepository performs operations directly (legacy).
        /// Default: true (enabled).
        /// </summary>
        public static bool UseSolidRefactoredRepositories { get; set; } = true;
        
        /// <summary>
        /// Enable dependency injection for refresh services.
        /// When true: Supports constructor injection of refresh service dependencies
        /// When false: Uses direct instantiation (legacy behavior)
        /// Default: false (disabled - enable after Phase 2 testing)
        /// Location: refresh refactor/refresh_service_refactored.cs
        /// Expected benefit: Better testability, loose coupling, easier mocking
        /// </summary>
        public static bool UseRefreshDependencyInjection { get; set; } = true;
        
        /// <summary>
        /// Enable split context interfaces for refresh operations.
        /// When true: Uses focused interfaces (IRefreshDocumentContext, IRefreshCacheContext, etc.)
        /// When false: Uses monolithic RefreshContext class (legacy behavior)
        /// Default: false (disabled - enable after Phase 3 testing)
        /// Location: refresh refactor/refresh_context.cs
        /// Expected benefit: Better Interface Segregation Principle compliance, focused dependencies
        /// </summary>
        public static bool UseRefreshSplitContext { get; set; } = true;
        
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
                UseParallelProcessing = GetConfigValue("UseParallelProcessing", true);
                UseSpatialGrid = GetConfigValue("UseSpatialGrid", true); // ✅ PERFORMANCE FIX: Enable spatial grid by default (70-90% reduction in intersection tests)
                UseRTreeDatabaseIndex = GetConfigValue("UseRTreeDatabaseIndex", true); // ✅ R-TREE: Enable database R-tree by default
                
                // Load Phase 3 flags (experimental defaults)
                UseIncrementalDetection = GetConfigValue("UseIncrementalDetection", false);
                UseDiagnosticMode = GetConfigValue("UseDiagnosticMode", false);
                UseBatchClashZoneCreation = GetConfigValue("UseBatchClashZoneCreation", true);
                
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
            UseRTreeFilter = false; // ✅ DISABLED: Coordinate mismatch with linked docs (fallback to spatial grid)
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
            UseRTreeFilter = false; // ✅ DISABLED: Coordinate mismatch with linked docs (keep spatial grid)
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
#pragma warning disable CS0618 // FIX: CS0618 - obsolete flags kept for backward compatibility logging
Refactoring Flags: SleeveRepository={UseNewSleeveRepository}, ZoneFilter={UseNewZoneFilter}, FamilyManager={UseNewFamilyManager}";
#pragma warning restore CS0618
        }
        
        #endregion
        /// <summary>
        /// Enable logic for V3 Unified Batch Mode (Calculate -> Persist -> Place).
        /// Replaces the legacy placement + clustering sequence with a unified workflow.
        /// </summary>
        public static bool EnableUnifiedBatchMode { get; set; } = true;

        /// <summary>
        /// Enable Geometric Center calculation for clustering (replaces Centroid/Average corners).
        /// Also enables use of native Revit BoundingBox for cluster cleanup.
        /// Fixes lateral shifts in rotated clusters and walls when sleeves are unbalanced.
        /// Default: false (safety flag per user request)
        /// </summary>
        public static bool UseGeometricCenterForClustering { get; set; } = false;

        /// <summary>When false: use intersection centroid for wall cluster placement (e.g. 0.72). When true: use geometric center of corners (e.g. 1.05).</summary>
        public static bool UseGeometricCenterForWalls { get; set; } = false;
        public static bool UseGeometricCenterForFloors { get; set; } = false;

        /// <summary>
        /// Enable the post-placement clustering workflow (Phase 2-4).
        /// When false: Skips corner extraction and clustering, only places individual sleeves.
        /// Default: true
        /// </summary>
        public static bool EnableClusteringWorkflow { get; set; } = false;

        /// <summary>
        /// Sleeve InstanceIds to trace in cluster_debug.log and batch_v2.log (e.g. 1436672, 1436679, 1436686).
        /// When a zone or cluster contains one of these IDs, a dedicated [TRACE SLEEVE] line is written.
        /// </summary>
        public static readonly System.Collections.Generic.HashSet<int> TraceSleeveIds = new System.Collections.Generic.HashSet<int> { 1436672, 1436679, 1436686 };

        public static bool ShouldTraceSleeve(int sleeveInstanceId) => sleeveInstanceId > 0 && TraceSleeveIds.Contains(sleeveInstanceId);

    }
}
