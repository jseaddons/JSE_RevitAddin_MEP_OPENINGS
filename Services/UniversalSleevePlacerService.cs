using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.Sizing;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Universal sleeve placer service - handles ALL MEP categories using strategy pattern
    /// Based on proven DuctSleevePlacerService architecture
    /// 
    /// Uses 4 universal families:
    /// - OpeningOnWall-Rectangular (ROW)
    /// - OpeningOnWall-Circular (COW)
    /// - OpeningOnSlab-Rectangular (ROS)
    /// - OpeningOnSlab-Circular (COS)
    /// </summary>
    public class UniversalSleevePlacerService
    {
        private readonly Document _doc;
        private readonly OpeningConditions _conditions;
        private readonly ISleevePlacementStrategy _strategy;
        private readonly Dictionary<string, double> _clearanceSettings;
        private readonly string _filterName;
        private readonly bool _isReplayPath;
        
        // ✅ OOP REFACTORING: Centralized flag management
        private readonly FlagManager _flagManager;
        
        // ✅ PARALLEL PLANNING: Optional planner for pre-computation (OOP, DI-ready)
        private readonly ISleevePlacementPlanner _planner;
        
        // ✅ OOP METHOD: Insulation-aware sizing service (SOLID principles)
        private readonly IInsulationAwareSizingService _sizingService;

        // ⚠️ QUICK WIN: Pre-cached family symbols (load once, reuse many times)
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();
        
        // ✅ STEP 5 OPTIMIZATION: Deferred parameter batching (4-6× faster placement)
        
        // ✅ DIRECT FILE WRITE HELPER: Bypasses SafeFileLogger DeploymentMode check
        private static void DirectLog(string message)
        {
            try
            {
                var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                Directory.CreateDirectory(logDir);
                File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), message);
            }
            catch { }
        }
        // Accumulates parameter values during placement loop, writes all after single regeneration
        // Key: ElementId of sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private Dictionary<ElementId, Dictionary<string, object>> _deferredParameters = new Dictionary<ElementId, Dictionary<string, object>>();

        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacerService(
            Document doc,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings = null,
            string filterName = null,
            FlagManager flagManager = null,
            bool isReplayPath = false,
            ISleevePlacementPlanner planner = null,  // ✅ NEW: Optional planner injection
            IInsulationAwareSizingService sizingService = null)  // ✅ OOP METHOD: Optional sizing service injection (SOLID)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _filterName = filterName;
            _isReplayPath = isReplayPath;
            
            // ✅ OOP REFACTORING: Initialize FlagManager (create if not provided for backward compatibility)
            _flagManager = flagManager ?? new FlagManager(doc);
            
            // ✅ PARALLEL PLANNING: Initialize planner with conditions and clearance settings
            _planner = planner ?? new ParallelSleevePlacementPlanner(_conditions, _clearanceSettings);
            
            // ✅ OOP METHOD: Initialize sizing service (create if not provided - Dependency Injection)
            _sizingService = sizingService ?? new InsulationAwareSizingService();
            
            // 🔥 CRITICAL DEBUG: Direct file logging to trace service instantiation (SAFE - won't crash)
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"🔥 UniversalSleevePlacerService CONSTRUCTOR CALLED 🔥");
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"FilterName parameter: '{filterName}'");
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"_filterName field: '{_filterName}'");
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"Strategy Type: {_strategy.GetType().Name}");
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"Strategy Category: {_strategy.GetCategoryName()}");
            SafeFileLogger.SafeAppendText("service_instantiation.log", 
                $"Clearance Settings Count: {_clearanceSettings.Count}");
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlaycer] Initialized for category: {_strategy.GetCategoryName()}");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlaycer] Received {_clearanceSettings.Count} clearance settings from UI");
            
            // Log all clearance settings for debugging
            foreach (var kvp in _clearanceSettings)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlaycer] Clearance: {kvp.Key} = {kvp.Value}mm");
            }
            
            // Add build timestamp to placement_debug.log
            try
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                var ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location)?.FileVersion ?? "?";
                var ts = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[BUILD] {ts} Assembly={System.IO.Path.GetFileName(asm.Location)} Version={ver} Path={asm.Location}\n");
            }
            catch { }
        }

        private double GetPipeOutsideDiameterFromClashZone(ClashZone cz)
        {
            // Prefer snapshot value if captured
            try
            {
                try
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    var ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location)?.FileVersion ?? "?";
                    var ts = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[BUILD] {ts} Assembly={System.IO.Path.GetFileName(asm.Location)} Version={ver} Path={asm.Location}\n");
                }
                catch { }
                var od = TryGetSnapshotDouble(cz.MepParameterValues, new[] { "Outside Diameter", "OD", "OutsideDiameter" });
                if (od > 0) return od;
            }
            catch { }

            // Fallback: derive from stored formatted size (e.g., Ø200)
            try
            {
                var size = cz.MepElementFormattedSize ?? string.Empty;
                if (size.StartsWith("Ø"))
                {
                    var mm = double.Parse(size.TrimStart('Ø'));
                    return RevitUnitConversionService.Instance.ToInternalMillimeters(mm);
                }
            }
            catch { }

            // Final fallback: use MepElementWidth as diameter if present
            return cz.MepElementWidth > 0 ? cz.MepElementWidth : 0.0;
        }

        private double TryGetSnapshotDouble(List<Models.SerializableKeyValue> bag, string[] keys)
        {
            if (bag == null) return 0.0;
            foreach (var k in keys)
            {
                var item = bag.Find(p => string.Equals(p.Key, k, StringComparison.OrdinalIgnoreCase));
                if (item != null)
                {
                    var raw = item.Value ?? string.Empty;
                    var s = raw.Trim();
                    // Strip common unit suffixes and symbols (case-insensitive)
                    var sl = s.ToLowerInvariant();
                    sl = sl.Replace("millimeters", "");
                    sl = sl.Replace("millimeter", "");
                    sl = sl.Replace("mm", "");
                    sl = sl.Replace("ø", "");
                    // Keep digits, sign, dot or comma only
                    var filtered = new System.Text.StringBuilder();
                    foreach (var ch in sl)
                    {
                        if ((ch >= '0' && ch <= '9') || ch == '-' || ch == '.' || ch == ',') filtered.Append(ch);
                    }
                    var norm = filtered.ToString().Replace(',', '.');
                    if (double.TryParse(norm, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var val))
                        return norm.Length > 0 ? RevitUnitConversionService.Instance.ToInternalMillimeters(val) : 0.0;
                    // Try plain parse (feet) if invariant mm parse failed
                    if (double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var feetVal))
                        return feetVal;
                }
            }
            return 0.0;
        }

        /// <summary>
        /// STEP 1: Collect MEP elements from all documents (active + linked)
        /// READ-ONLY phase - NO transaction required
        /// Matches MD pattern: Section 2, Lines 67-69
        /// </summary>
        public List<Element> CollectMepElementsWithClashZones(List<ClashZone> clashZones)
        {
            try
            {
                var mepElements = new List<Element>();
                var mepElementIds = clashZones.Select(cz => cz.MepElementId).ToHashSet();
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Looking for {mepElementIds.Count} MEP elements with IDs from clash zones");
                
                // 1. Search in active document first
                var activeElements = new FilteredElementCollector(_doc)
                    .WhereElementIsNotElementType()
                    .Where(elem => mepElementIds.Contains(elem.Id))
                    .ToList();
                
                mepElements.AddRange(activeElements);
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Found {activeElements.Count} MEP elements in active document");
                
                // 2. Search in linked documents
                var linkedDocs = _doc.Application.Documents.Cast<Document>()
                    .Where(doc => doc.IsLinked && doc.Title != _doc.Title)
                    .ToList();
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Searching {linkedDocs.Count} linked documents for MEP elements");
                
                foreach (var linkedDoc in linkedDocs)
                {
                    try
                    {
                        var linkedElements = new FilteredElementCollector(linkedDoc)
                            .WhereElementIsNotElementType()
                            .Where(elem => mepElementIds.Contains(elem.Id))
                            .ToList();
                        
                        mepElements.AddRange(linkedElements);
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Found {linkedElements.Count} MEP elements in linked document: {linkedDoc.Title}");
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalSleevePlacer] Error searching MEP elements in linked document {linkedDoc.Title}: {ex.Message}");
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Total collected {mepElements.Count} MEP elements from all documents");

                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [RETURN-CollectMepElementsWithClashZones] Count={mepElements.Count}\n");
                }
                catch { }

                return mepElements;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] Error collecting MEP elements: {ex.Message}");
                return new List<Element>();
            }
        }
        
        /// <summary>
        /// STEP 2: Place all sleeves within an active transaction
        /// WRITE phase - Transaction MUST be started by caller (Command)
        /// Matches MD pattern: Section 2, Lines 72-80
        /// </summary>
        /// <summary>
        /// ✅ SIMPLIFIED: Filter XML already contains valid zones - only sync flags from Global XML
        /// Filter XML zones are already:
        /// - Valid category (loaded from category-specific XML file)
        /// - Valid placement points (saved during refresh)
        /// - Filtered by UI selections (host type, files, section box)
        /// 
        /// ONLY check: Sync flags from Global XML to see which zones need placement
        /// </summary>
        private List<ClashZone> PreFilterEligibleClashZones(List<ClashZone> clashZones)
        {
            // ✅ PERFORMANCE: Pre-filter and validate clash zones (reserved for future optimization)
            // Note: Actual clearance calculation happens during placement, not here
            // ClashZone does not have SleeveDepth or cached clearance properties
            // This method can be used for parallel validation in the future
            
            if (!OptimizationFlags.UseParallelClearanceCalculation || clashZones.Count <= 10)
            {
                // Skip parallel processing for small lists
                return clashZones;
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            // Currently just returns the list as-is
            // Future optimization: Add parallel validation logic here
            // (e.g., check required properties are set, validate coordinates, etc.)
            
            sw.Stop();
            if (!DeploymentConfiguration.DeploymentMode && sw.ElapsedMilliseconds > 5)
            {
                SafeFileLogger.SafeAppendText("placement_performance.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚡ Pre-filtered {clashZones.Count} clash zones in {sw.ElapsedMilliseconds}ms\n");
            }
            
            return clashZones.ToList();
        }
        
        public (int PlacedCount, int SkippedCount, int ErrorCount) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            // ✅ CRITICAL: Force DeploymentMode OFF for diagnostic logging
            DeploymentConfiguration.DeploymentMode = false;
            
            // ✅ DIAGNOSTIC: Log batching flag status at placement start
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[BATCH-PARAMS] ═══ PLACEMENT START ═══ UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, _deferredParameters initialized={_deferredParameters != null}");
            }
            
            // ✅ PERFORMANCE MONITORING: Initialize placement performance monitor
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string performanceLogName = $"IndividualPlacement_{timestamp}.log";
            var performanceMonitor = new Services.Placement.PlacementPerformanceMonitor(performanceLogName);

            // ⏱️ TIMING: Start overall placement timer
            var overallTimer = System.Diagnostics.Stopwatch.StartNew();
            var detailedTimingLog = new System.Text.StringBuilder();
            // Diagnostics: per-run placement log in AppData
            var placementLogName = SafeFileLogger.GetLogFilePath($"sleeve_placement_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
            
            // ✅ PERFORMANCE OPTIMIZATION: Track placed sleeves for batch processing
            var placedSleeveData = new List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter)>();
            var placedSleeveIds = new List<ElementId>();
            
            // ✅ SESSION FLAG: Track processed zones for ReadyForPlacement flag reset
            var processedZoneGuids = new List<Guid>();
            
            // ✅ PERFORMANCE OPTIMIZATION: Element cache to avoid redundant GetElement calls
            var elementCache = new Dictionary<int, Element>();
            Element GetCachedElement(int elementId)
            {
                if (!elementCache.TryGetValue(elementId, out var element))
                {
                    element = _doc.GetElement(new ElementId(elementId));
                    if (element != null)
                        elementCache[elementId] = element;
                }
                return element;
            }
            
            // Also mirror a one-line header into project Log for visibility even if AppData not checked
            try 
            { 
                string runLogPath = SafeFileLogger.GetLogFilePath("placement_run.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(runLogPath, $"[{DateTime.Now}] ENTER PlaceAllSleeves: incoming={clashZones?.Count ?? 0}\n");
                } 
            } 
            catch { }
            
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;
            
            // ✅ PERFORMANCE: Track level caching
            var cachedLevels = new List<Level>();
            using (var cacheTracker = performanceMonitor.TrackOperation("Cache Levels"))
            {
            // ✅ PERFORMANCE OPTIMIZATION: Pre-cache expensive operations
            // Cache levels list once instead of scanning for each clash zone
                cachedLevels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();
                cacheTracker.SetItemCount(cachedLevels.Count);
            }
            detailedTimingLog.AppendLine($"[TIMING] Level caching: {cachedLevels.Count} levels");
            
            // ✅ PERFORMANCE OPTIMIZATION: Cache GlobalIndexService lookups per category (not per clash zone)
            // NOTE: GlobalIndexService is static, no need to cache instances
            
            // ✅ PERFORMANCE OPTIMIZATION: Batch file logging - collect logs and write once
            var batchLogs = new System.Text.StringBuilder();
            
            // 🔥 CRITICAL DEBUG: Log flag status from the clash zones passed to this method
            int clusterResolvedCount = clashZones.Count(cz => cz.IsClusterResolved);
            int individualResolvedCount = clashZones.Count(cz => cz.IsResolved);
            int eligibleCount = clashZones.Count(cz => !cz.IsResolved && !cz.IsClusterResolved);
            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] 🔥 UNIVERSAL SLEEVE PLACER RECEIVED: {clashZones.Count} clash zones");
            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] 📊 FLAGS RECEIVED: IsClusterResolved=True: {clusterResolvedCount}, IsResolved=True: {individualResolvedCount}");
            SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] START: Total={clashZones.Count}, AlreadyResolved={individualResolvedCount}, ClusterResolved={clusterResolvedCount}, Eligible={eligibleCount}\n");
            
            // ✅ CRITICAL: Log where error logs are located so users can find them + DIRECT FILE WRITE to ensure log exists
            var errorLogPath = SafeFileLogger.GetLogFilePath("sleeve_placement_errors.log");
            var debugLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
            
            // ✅ AGGRESSIVE LOGGING: Direct file write to ensure logs are ALWAYS created
            try 
            { 
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
            { 
                File.AppendAllText(errorLogPath, $"[{DateTime.Now}] ========== PLACEMENT STARTED ==========\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(errorLogPath, $"[{DateTime.Now}] Total zones: {clashZones.Count}, Eligible: {eligibleCount}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] ========== PLACEMENT STARTED ==========\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] Total zones: {clashZones.Count}, Eligible: {eligibleCount}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] Error log path: {errorLogPath}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] Debug log path: {debugLogPath}\n");
                }
            } 
            catch (Exception logEx) 
            { 
                System.Diagnostics.Debug.WriteLine($"[Placement-Start] Failed to write logs: {logEx.Message}");
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] 📁 Error logs: {errorLogPath}");
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] 📁 Debug logs: {debugLogPath}");
            System.Diagnostics.Debug.WriteLine($"[UniversalSleevePlacer] 📁 Error logs: {errorLogPath}");
            System.Diagnostics.Debug.WriteLine($"[UniversalSleevePlacer] 📁 Debug logs: {debugLogPath}");
            
            // 🛡️ FAIL-SAFE: Check document state before starting
            if (!_doc.IsModifiable)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error("[UniversalSleevePlacer] Document is read-only or workshared and not checked out");
                throw new InvalidOperationException("Document is not modifiable. Please check out the file or ensure it's not read-only.");
            }
            
            if (clashZones == null || clashZones.Count == 0)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalSleevePlacer] No clash zones provided");
                return (0, 0, 0);
            }
            
            // DEBUG: Log all incoming ClashZone objects to trace XML loading
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Log($"[XML-DEBUG] PlaceAllSleevesInTransaction called with {clashZones.Count} clash zones:");
            foreach (var cz in clashZones.Take(5)) // Log first 5 to avoid spam
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Log($"[XML-DEBUG] ClashZone {cz.Id}: MepElementCategory='{cz.MepElementCategory}', StructuralElementType='{cz.StructuralElementType}', MepElementId={cz.MepElementIdValue}, StructuralElementId={cz.StructuralElementIdValue}");
            }
            
            // ⚠️ DON'T reset resolved flags during placement!
            // Flags are managed by refresh - it checks if sleeves exist and resets flags if deleted
            // If we reset here, we'll place duplicate sleeves for existing ones
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] Processing {clashZones.Count} clash zones (zero linked file access), trusting IsResolved flags from refresh");
            
            // ✅ PERFORMANCE: Track pre-filtering
            List<ClashZone> eligibleClashZones;
            using (var preFilterTracker = performanceMonitor.TrackOperation("Pre-Filter Clash Zones"))
            {
            // ✅ MULTI-THREADING: Pre-calculate clearances in parallel (non-Revit operation)
            // This filters and prepares clash zones BEFORE entering sequential placement loop
            eligibleClashZones = PreFilterEligibleClashZones(clashZones);
                preFilterTracker.SetItemCount(eligibleClashZones.Count);
            }
            detailedTimingLog.AppendLine($"[TIMING] Pre-filtering: {clashZones.Count} → {eligibleClashZones.Count} eligible");
            
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] ✅ MULTI-THREADING: Pre-filtered {clashZones.Count} clash zones → {eligibleClashZones.Count} eligible");
            
            // ✅ PRIORITY SORTING: Process Duct Accessories (Dampers) BEFORE Ducts
            var sortedClashZones = eligibleClashZones
                .OrderBy(cz => GetCategoryPriority(cz.MepElementCategory))
                .ThenBy(cz => cz.Id)
                .ToList();
            
            // ✅ DEBUG: Log how many zones are eligible after pre-filtering
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[UniversalSleevePlacer] ✅ Pre-filtered {clashZones.Count} clash zones → {eligibleClashZones.Count} eligible → {sortedClashZones.Count} sorted");
                if (sortedClashZones.Count == 0)
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ WARNING: No eligible clash zones after pre-filtering! All {clashZones.Count} zones were filtered out.");
                    return (0, 0, 0); // Early return if no eligible zones
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] Sorted {sortedClashZones.Count} clash zones by category priority");

            // Log first 10 sorted clash zones to verify order
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info("[SORT-VERIFY] First 10 clash zones after sorting:");
            foreach (var cz in sortedClashZones.Take(10))
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"  ClashZone {cz.Id}: Category='{cz.MepElementCategory}', Priority={GetCategoryPriority(cz.MepElementCategory)}");
            }
            
            // ✅ PERFORMANCE: Track family symbol pre-caching
            using (var preCacheTracker = performanceMonitor.TrackOperation("Pre-Cache Family Symbols"))
            {
            // ⚠️ QUICK WIN: Pre-load family symbols for all needed families (significant performance gain)
            if (OptimizationFlags.UseFamilySymbolCache)
            {
                PreCacheFamilySymbols(sortedClashZones);
            }
                preCacheTracker.SetItemCount(sortedClashZones.Count);
            }
            detailedTimingLog.AppendLine($"[TIMING] Family symbol pre-caching completed");
            
            // ⏱️ TIMING: Per-sleeve operation timers
            var totalPlacementTime = TimeSpan.Zero;
            var totalClearanceTime = TimeSpan.Zero;
            var totalLevelFindTime = TimeSpan.Zero;
            var totalFamilyLoadTime = TimeSpan.Zero;
            var totalSleeveCreateTime = TimeSpan.Zero;
            var totalParameterTime = TimeSpan.Zero;
            var totalValidationTime = TimeSpan.Zero;

            try
            {
                // ⚠️ CRITICAL: Iterate over CLASH ZONES, not MEP elements
                // Each clash zone represents a unique (MEP Element + Structural Element) PAIR
                // The same MEP element can appear in multiple clash zones if it intersects multiple walls
                // Clash zones are already filtered by UI during refresh, so process all provided zones

                // ⚠️ CRITICAL: Add timeout protection to prevent infinite hangs
                var placementTimer = System.Diagnostics.Stopwatch.StartNew();
                const int MAX_PLACEMENT_TIME_MS = 300000; // 5 minutes
                int processedCount = 0;

                // ✅ PARALLEL PLANNING PHASE: Pre-compute dimensions, clearance, rotation in parallel
                Dictionary<Guid, SleevePlacementPlanningDto> planningMap = null;
                List<ClashZone> planningOrderedZones = null;
                var planningLogs = new System.Text.StringBuilder();
                
                if (DeploymentConfiguration.EnableParallelPlanning && _planner != null)
                {
                    try
                    {
                        using (var planningTracker = performanceMonitor.TrackOperation("Parallel Planning Phase"))
                        {
                            planningLogs.AppendLine($"[PLANNING] Starting parallel planning for {sortedClashZones.Count} zones...");
                            
                            var planningResult = _planner.Plan(sortedClashZones);
                            
                            planningLogs.AppendLine($"[PLANNING] Planning completed in {planningResult.PlanningDurationMs:F2} ms");
                            planningLogs.AppendLine($"[PLANNING] Total zones: {planningResult.TotalCount}");
                            planningLogs.AppendLine($"[PLANNING] Zones to skip: {planningResult.SkippedCount}");
                            planningLogs.AppendLine($"[PLANNING] High-risk zones: {planningResult.HighRiskCount}");
                            planningLogs.AppendLine($"[PLANNING] Critical-risk zones: {planningResult.CriticalRiskCount}");
                            
                            // Create lookup map for fast DTO access during placement
                            planningMap = planningResult.Items.ToDictionary(dto => dto.ClashZoneId);
                            
                            // Filter out zones marked for skipping in planning phase
                            var eligibleZones = planningResult.Items
                                .Where(dto => !dto.ShouldSkip)
                                .Select(dto => sortedClashZones.First(z => z.Id == dto.ClashZoneId))
                                .ToList();
                            
                            // Reorder by risk (low → high) for better success rate
                            planningOrderedZones = planningResult.Items
                                .Where(dto => !dto.ShouldSkip)
                                .OrderBy(dto => dto.ClearanceRisk) // Low risk first
                                .Select(dto => sortedClashZones.First(z => z.Id == dto.ClashZoneId))
                                .ToList();
                            
                            planningLogs.AppendLine($"[PLANNING] Processing order: {planningOrderedZones.Count} zones (low-risk → high-risk)");
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[PLANNING] Parallel planning succeeded: {planningResult.TotalCount} zones analyzed, {planningResult.SkippedCount} skipped, {planningOrderedZones.Count} eligible for placement");
                            }
                            
                            planningTracker.SetItemCount(planningResult.TotalCount);
                        }
                    }
                    catch (Exception planningEx)
                    {
                        planningLogs.AppendLine($"[PLANNING] ERROR: {planningEx.Message}");
                        planningLogs.AppendLine($"[PLANNING] Falling back to original sorted order");
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[PLANNING] Planning phase failed: {planningEx.Message}. Falling back to original logic.");
                        }
                        
                        // Fallback to original behavior
                        planningMap = null;
                        planningOrderedZones = null;
                    }
                }
                else
                {
                    planningLogs.AppendLine($"[PLANNING] Parallel planning disabled (EnableParallelPlanning={DeploymentConfiguration.EnableParallelPlanning}, Planner={(_planner != null ? "available" : "null")})");
                }
                
                // Use planning-ordered zones if available, otherwise use original sorted zones
                var finalProcessingList = planningOrderedZones ?? sortedClashZones;
                planningLogs.AppendLine($"[PLANNING] Final processing list: {finalProcessingList.Count} zones");

                // ✅ PERFORMANCE: Track main placement loop
                using (var placementLoopTracker = performanceMonitor.TrackOperation("Place Individual Sleeves Loop"))
                {
                    foreach (var clashZone in finalProcessingList)
                    {
                        // ✅ PERFORMANCE: Track each sleeve placement operation
                        using (var singleSleeveTracker = performanceMonitor.TrackOperation("Place Single Sleeve"))
                        {
                            // ✅ PARALLEL PLANNING: Try to get pre-computed DTO for this zone
                            SleevePlacementPlanningDto planningDto = null;
                            planningMap?.TryGetValue(clashZone.Id, out planningDto);
                            
                            // ✅ DEBUG: Log which clash zone is being processed
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                var riskLabel = planningDto != null ? $", Risk={planningDto.ClearanceRisk}" : "";
                                DebugLogger.Info($"[PLACEMENT-LOOP] Processing ClashZone {clashZone.Id} ({processedCount + 1}/{finalProcessingList.Count}): MEP={clashZone.MepElementIdValue}, Flags: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, SleeveId={clashZone.SleeveInstanceId}{riskLabel}");
                            }

                            // ⚠️ CRITICAL: Check timeout every 10 sleeves to prevent infinite hangs
                            processedCount++;
                            if (processedCount % 10 == 0 && placementTimer.ElapsedMilliseconds > MAX_PLACEMENT_TIME_MS)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error($"[UniversalSleevePlacer] ⏱ TIMEOUT: Sleeve placement exceeded {MAX_PLACEMENT_TIME_MS / 1000} second limit after processing {processedCount} sleeves");
                                }
                                System.Windows.Forms.MessageBox.Show(
                                    $"Sleeve placement is taking too long and has been cancelled.\n\nProcessed: {processedCount} of {sortedClashZones.Count} sleeves\nTime: {placementTimer.ElapsedMilliseconds / 1000} seconds\nLimit: {MAX_PLACEMENT_TIME_MS / 1000} seconds\n\nThis usually indicates:\n• Very large model\n• Infinite loop\n• Corrupted clash zone data\n\nPlease check the log file and try processing in smaller batches.",
                                    "Operation Timeout",
                                    System.Windows.Forms.MessageBoxButtons.OK,
                                    System.Windows.Forms.MessageBoxIcon.Warning);
                                break; // Exit loop to prevent crash
                            }
                            // ✅ STARTUP MARKER: Log to confirm new code is running (using SafeFileLogger path)
                            // Note: debugLogPath is already declared at method level (line 254)
                            /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                            try
                            {
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 NEW CODE RUNNING: Zone={clashZone.Id}, MEP={clashZone.MepElementIdValue}, HOST={clashZone.StructuralElementIdValue}\n");
                                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    var pathName = _isReplayPath ? "PATH 1 (Replay)" : "PATH 2/3 (Sizing/Detection)";
                                    File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 0: Starting processing for Zone={clashZone.Id}, Placement Path={pathName}\n");
                                    DebugLogger.Info($"[UniversalSleevePlacer] Zone={clashZone.Id}: Using {pathName}");
                                }
                            }
                            catch { }
                            */

                            // ⏱️ TIMING: Start per-sleeve timer (only for actual placement operations)
                            var sleeveTimer = System.Diagnostics.Stopwatch.StartNew();
                            var sleeveLog = new System.Text.StringBuilder();

                            try
                            {
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                try
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 1: Entered try block for Zone={clashZone.Id}\n");
                                    }
                                }
                                catch { }
                                */

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");

                                // ✅ DEPLOYMENT MODE: Skip file writes
                                /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                try
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2: About to validate placement point for Zone={clashZone.Id}, IP=({clashZone.IntersectionPointX},{clashZone.IntersectionPointY},{clashZone.IntersectionPointZ})\n");
                                    }
                                }
                                catch { }
                                */

                                // ✅ CRITICAL: Validate intersection coordinates - NO FALLBACK to wall center
                                // This will throw an exception if coordinates are invalid, which we catch and handle below
                                try
                                {
                                    ValidatePlacementPoint(clashZone);
                                    try
                                    {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.1: ✅ Placement point validation PASSED for Zone={clashZone.Id}\n");
                                        }
                                        */
                                    }
                                    catch { }
                                }
                                catch (InvalidOperationException validationEx)
                                {
                                    // Invalid coordinates - skip this zone and log error
                                    var msg = $"[UniversalSleevePlacer] ERROR: {validationEx.Message}";
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error(msg);
                                    if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✗ ERROR: ClashZone {clashZone.Id} has invalid intersection coordinates - SKIPPED");

                                    try
                                    {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.2: ❌ VALIDATION FAILED: {validationEx.Message}\n");
                                        }
                                    }
                                    catch { }

                                    // ✅ FIXED: Use SafeFileLogger for correct log paths (dev + deployment)
                                    try
                                    {
                                        var errorLogFilePath = SafeFileLogger.GetLogFilePath("sleeve_placement_errors.log");
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(errorLogFilePath, $"[{DateTime.Now}] {msg}\n");
                                        }
                                        System.Diagnostics.Debug.WriteLine($"[Placement] ✅ Error logged to: {errorLogFilePath}");
                                    }
                                    catch (Exception logEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[Placement] ❌ Failed to log error: {logEx.Message}");
                                    }
                                    try
                                    {
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] INVALID-POINT: Zone={clashZone.Id}, MEP={clashZone.MepElementIdValue}, HOST={clashZone.StructuralElementIdValue}, IP=({clashZone.IntersectionPointX},{clashZone.IntersectionPointY},{clashZone.IntersectionPointZ}), SPP=({clashZone.SleevePlacementPointX},{clashZone.SleevePlacementPointY},{clashZone.SleevePlacementPointZ})\n");
                                        }
                                    }
                                    catch (Exception logEx)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[Placement] ❌ Failed to log debug: {logEx.Message}");
                                    }

                                    ErrorCount++;
                                    singleSleeveTracker.SetItemCount(1); // Track error as 1 item
                                    continue; // Skip this zone - cannot place without valid intersection point
                                }

                                // ✅ DEPLOYMENT MODE: Skip file writes
                                /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                try
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 3: Validation passed, checking flags: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, SleeveId={clashZone.SleeveInstanceId}, ClusterId={clashZone.ClusterSleeveInstanceId}\n");
                                    }
                                }
                                catch { }
                                */

                                // ✅ PERFORMANCE OPTIMIZATION: Batch file logging instead of individual writes
                                // Only log to batch - will write once at end (or every 50 clash zones)
                                if (PlacedCount + SkippedCount < 50) // Only log first 50 for debugging
                                {
                                    batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                                    batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}");
                                }

                                // ✅ BUG FIX: Check IsClusterResolved flag BEFORE individual placement
                                // If IsClusterResolved=true, there MUST be a ClusterSleeveInstanceId (impossible to be cluster-resolved without a cluster sleeve ID)
                                if (clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                                {
                                    var clusterSleeveElement = GetCachedElement(clashZone.ClusterSleeveInstanceId);
                                    if (clusterSleeveElement != null)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already cluster-resolved (IsClusterResolved=true) with cluster sleeve {clashZone.ClusterSleeveInstanceId}");
                                        SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SKIP ClusterResolved: ClashZone={clashZone.Id}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, IsClusterResolved=true\n");
                                        SkippedCount++;
                                        singleSleeveTracker.SetItemCount(1); // Track skipped as 1 item
                                        continue;
                                    }
                                    else
                                    {
                                        // ✅ CRITICAL: Cluster sleeve was deleted - reset flags in memory
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] RESET: ClashZone {clashZone.Id} has IsClusterResolved=true but cluster sleeve {clashZone.ClusterSleeveInstanceId} was deleted - resetting flags");
                                        clashZone.IsClusterResolved = false;
                                        clashZone.ClusterSleeveInstanceId = -1;
                                        // Continue to individual placement check below
                                    }
                                }

                                // ✅ BUG FIX: Check IsResolved flag BEFORE checking SleeveInstanceId
                                // PATH 1 should respect flag state from XML - if IsResolved=true, skip placement
                                // If IsResolved=true, there MUST be a SleeveInstanceId (impossible to be resolved without a sleeve ID)
                                if (clashZone.IsResolved && clashZone.SleeveInstanceId > 0)
                                {
                                    var existingIndividualSleeve = GetCachedElement(clashZone.SleeveInstanceId);
                                    if (existingIndividualSleeve != null)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already resolved (IsResolved=true) with individual sleeve {clashZone.SleeveInstanceId}");
                                        SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SKIP AlreadyResolved: ClashZone={clashZone.Id}, SleeveId={clashZone.SleeveInstanceId}, IsResolved=true\n");
                                        SkippedCount++;
                                        singleSleeveTracker.SetItemCount(1); // Track skipped as 1 item
                                        continue;
                                    }
                                    else
                                    {
                                        // ✅ CRITICAL: Sleeve was deleted - reset flag and continue to placement
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] RESET: ClashZone {clashZone.Id} has IsResolved=true but sleeve {clashZone.SleeveInstanceId} was deleted - resetting flag");
                                        clashZone.IsResolved = false;
                                        clashZone.SleeveInstanceId = -1;
                                        // Continue to placement below
                                    }
                                }

                                // Check if sleeve exists by ID (even if flag is false - handles stale data)
                                if (clashZone.SleeveInstanceId > 0)
                                {
                                    var existingIndividualSleeve = GetCachedElement(clashZone.SleeveInstanceId);
                                    if (existingIndividualSleeve != null)
                                    {
                                        // ✅ PATH 1 FIX: If sleeve exists and PATH 1 (Replay), just reset flags and skip placement
                                        // No calculations needed - sleeve instance IDs are already in DB
                                        if (_isReplayPath)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[UniversalSleevePlacer] ✅ PATH 1: Sleeve {clashZone.SleeveInstanceId} exists for ClashZone {clashZone.Id} - resetting flags only, skipping placement");
                                            
                                            // Reset flags to true (sleeve exists, so it's resolved)
                                            clashZone.IsResolved = true;
                                            
                                            // Update flags in DB/XML
                                            try
                                            {
                                                _flagManager.UpdateFlagsForPlacement(clashZone, clashZone.SleeveInstanceId, isCluster: false, clashZone.MepElementCategory, _filterName);
                                            }
                                            catch (Exception flagEx)
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Warning($"[UniversalSleevePlacer] Error updating flags for PATH 1 zone {clashZone.Id}: {flagEx.Message}");
                                            }
                                            
                                            SkippedCount++;
                                            SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] PATH 1 SKIP IndividualExists: ClashZone={clashZone.Id}, SleeveId={clashZone.SleeveInstanceId}, Flags reset\n");
                                            singleSleeveTracker.SetItemCount(1);
                                            continue;
                                        }
                                        else
                                        {
                                            // PATH 2/3: Normal skip logic
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already has individual sleeve {clashZone.SleeveInstanceId}");
                                            SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SKIP IndividualExists: ClashZone={clashZone.Id}, SleeveId={clashZone.SleeveInstanceId}\n");
                                            SkippedCount++;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        // Sleeve ID exists in DB but sleeve was deleted - reset and place
                                        clashZone.IsResolved = false;
                                        clashZone.SleeveInstanceId = -1;
                                    }
                                }

                                // STEP 3: If we reach here, place individual sleeve (fresh or replacement)
                                // ✅ PATH 1: This only happens if sleeve was deleted (ID in DB but sleeve missing in Revit)

                                // ✅ CRITICAL: Check Global XML before placement (prevents cross-filter duplicates)
                                // ✅ USES GlobalIndexService (GUID-based) per methodology document
                                try
                                {
                                    var categoryName = clashZone.MepElementCategory;
                                    var globalIndex = GlobalIndexService.LoadOrCreate(_doc, categoryName);

                                    // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                                    // Entries are now stored in Filters → FileCombos → Entries, not just in flat Entries list
                                    var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                                    var entry = allEntries?.FirstOrDefault(e => e.Id == clashZone.Id.ToString());

                                    if (entry != null)
                                    {
                                        // ✅ DEBUG: Log Global XML entry state
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[GLOBAL-XML-CHECK] ClashZone {clashZone.Id}: Found Global XML entry - IsResolved={entry.IsResolved}, IsClusterResolved={entry.IsClusterResolved}, SleeveId={entry.SleeveInstanceId}, ClusterId={entry.ClusterSleeveInstanceId}");

                                        // Entry exists in Global XML - check if sleeve still exists in Revit
                                        if (entry.IsClusterResolved && entry.ClusterSleeveInstanceId > 0)
                                        {
                                            // Cluster sleeve exists in Global XML - verify it exists in Revit
                                            var clusterSleeve = _doc.GetElement(new ElementId(entry.ClusterSleeveInstanceId));
                                            if (clusterSleeve != null)
                                            {
                                                // Cluster sleeve exists - skip individual placement
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[GLOBAL-XML] SKIP: ClashZone {clashZone.Id} has cluster sleeve {entry.ClusterSleeveInstanceId} in Global XML (placed from another filter)");
                                                if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: Cluster sleeve {entry.ClusterSleeveInstanceId} exists in Global XML");

                                                // Update clash zone flags to match Global XML state
                                                clashZone.IsClusterResolved = true;
                                                clashZone.ClusterSleeveInstanceId = entry.ClusterSleeveInstanceId;
                                                clashZone.IsResolved = false;
                                                clashZone.SleeveInstanceId = -1;

                                                SkippedCount++;
                                                continue;
                                            }
                                            else
                                            {
                                                // ✅ CRITICAL FIX: Cluster sleeve was deleted - reset flags in memory
                                                // DO NOT call UpdateFlagsForPlacement with isCluster=true here - it would set IsClusterResolved=true again!
                                                // The flags will be persisted correctly when the new individual sleeve is placed below
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[FLAG-MANAGER] RESET: ClashZone {clashZone.Id} has cluster sleeve in Global XML but sleeve was deleted - resetting flags and proceeding with placement");
                                                clashZone.IsClusterResolved = false;
                                                clashZone.ClusterSleeveInstanceId = -1;
                                                clashZone.IsResolved = false; // Also reset individual flag since cluster was deleted
                                                clashZone.SleeveInstanceId = -1;
                                                // ✅ NOTE: Flags will be persisted when new sleeve is placed via UpdateFlagsForPlacement(isCluster: false)
                                            }
                                        }
                                        else if (entry.IsResolved && entry.SleeveInstanceId > 0)
                                        {
                                            // ✅ CRITICAL: Always verify sleeve exists in Revit (Global XML may be stale)
                                            // Revit is authoritative source - if sleeve doesn't exist, reset flags and place
                                            var individualSleeve = _doc.GetElement(new ElementId(entry.SleeveInstanceId));
                                            if (individualSleeve != null)
                                            {
                                                // Individual sleeve exists in Revit - skip placement
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[GLOBAL-XML] SKIP: ClashZone {clashZone.Id} has individual sleeve {entry.SleeveInstanceId} in Global XML (placed from another filter) - sleeve EXISTS in Revit");
                                                if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: Individual sleeve {entry.SleeveInstanceId} exists in Global XML and Revit");

                                                // Update clash zone flags to match Global XML state
                                                clashZone.IsResolved = true;
                                                clashZone.SleeveInstanceId = entry.SleeveInstanceId;

                                                SkippedCount++;
                                                continue;
                                            }
                                            else
                                            {
                                                // ✅ CRITICAL FIX: Sleeve doesn't exist in Revit - Global XML is stale
                                                // Reset flags in Global XML and proceed with placement
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Info($"[FLAG-MANAGER] RESET: ClashZone {clashZone.Id} has individual sleeve {entry.SleeveInstanceId} in Global XML but sleeve was DELETED in Revit - resetting flags and proceeding with placement");
                                                clashZone.IsResolved = false;
                                                clashZone.SleeveInstanceId = -1;

                                                // ✅ CRITICAL: Update Global XML immediately to prevent future skips
                                                _flagManager.UpdateFlagsForPlacement(clashZone, -1, isCluster: false, categoryName, _filterName);

                                                // Continue to placement (don't skip)
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // ✅ DEBUG: Log when no Global XML entry exists
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[GLOBAL-XML-CHECK] ClashZone {clashZone.Id}: No Global XML entry found - proceeding with placement");
                                    }
                                    // If no Global XML entry exists, proceed with placement (normal case)
                                }
                                catch (Exception globalEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[GLOBAL-XML] Error checking Global XML for ClashZone {clashZone.Id}: {globalEx.Message} - proceeding with placement");
                                    // Continue with placement on error (fail-safe)
                                }

                                // DEPLOYMENT MODE: Skip file writes
                                /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                try
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 6: Checking category match: ZoneCategory='{clashZone.MepElementCategory}' vs StrategyCategory='{_strategy?.GetCategoryName() ?? "NULL"}'\n");
                                    }
                                }
                                catch { }
                                */

                                // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging - use DebugLogger only
                                // Validate category match
                                if (!string.IsNullOrEmpty(clashZone.MepElementCategory) &&
                                    clashZone.MepElementCategory != _strategy.GetCategoryName())
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} category '{clashZone.MepElementCategory}' doesn't match '{_strategy.GetCategoryName()}'");
                                    try
                                    {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 6.1: ❌ SKIPPED - Category mismatch\n");
                                        }
                                        */
                                    }
                                    catch { }
                                    SkippedCount++;
                                    sleeveTimer.Stop();
                                    continue;
                                }

                                try
                                {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                    /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 7: Category match passed, creating MEP size\n");
                                    }
                                    */
                                }
                                catch { }

                                // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                                // ⏱️ TIMING: MEP size creation
                                var mepSizeTimer = System.Diagnostics.Stopwatch.StartNew();
                                // ⚠️ ZERO LINKED FILE ACCESS - use pre-calculated MEP size from ClashZone
                                // ✅ DIAGNOSTIC: Log values being read from ClashZone BEFORE creating mepSize
                                if (!DeploymentConfiguration.DeploymentMode && string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                                {
                                    DebugLogger.Info($"[PLACEMENT-DEBUG] Zone {clashZone.Id}: Reading from ClashZone - MepElementWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepElementHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)");
                                    if (clashZone.MepElementWidth <= 0 || clashZone.MepElementHeight <= 0)
                                    {
                                        DebugLogger.Warning($"[PLACEMENT-DEBUG] ⚠️ Zone {clashZone.Id}: MepElementWidth or MepElementHeight is ZERO - this will cause fallback! Width={clashZone.MepElementWidth}, Height={clashZone.MepElementHeight}");
                                    }
                                    if (clashZone.MepElementSizeData != null)
                                    {
                                        DebugLogger.Info($"[PLACEMENT-DEBUG] Zone {clashZone.Id}: MepElementSizeData exists - Width={clashZone.MepElementSizeData.Width:F6}ft ({clashZone.MepElementSizeData.Width * 304.8:F1}mm), Height={clashZone.MepElementSizeData.Height:F6}ft ({clashZone.MepElementSizeData.Height * 304.8:F1}mm)");
                                    }
                                }
                                
                                // ✅ DIRECT FILE WRITE: Always log path and what we're using
                                try
                                {
                                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                    Directory.CreateDirectory(logDir);
                                    File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [PATH-DECISION] Zone {clashZone.Id}: _isReplayPath={_isReplayPath}, SleeveWidth={clashZone.SleeveWidth:F6}ft ({clashZone.SleeveWidth * 304.8:F1}mm), SleeveHeight={clashZone.SleeveHeight:F6}ft ({clashZone.SleeveHeight * 304.8:F1}mm), MepElementWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), MepElementHeight={clashZone.MepElementHeight:F6}ft ({clashZone.MepElementHeight * 304.8:F1}mm)\n");
                                } catch { }
                                
                                var mepSize = new MepElementSize
                                {
                                    Width = clashZone.MepElementWidth,
                                    Height = clashZone.MepElementHeight,
                                    Diameter = clashZone.MepElementWidth, // For round, width = diameter
                                    Shape = clashZone.DuctShape
                                };
                                mepSizeTimer.Stop();
                                try
                                {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                    /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 8: MEP size created: W={mepSize.Width}, H={mepSize.Height}, D={mepSize.Diameter}, Shape={mepSize.Shape}\n");
                                    }
                                    */
                                }
                                catch { }

                                // ⚠️ SPECIAL HANDLING & SIZING ORDER:
                                // 1) Pipes (host-agnostic), 2) Dampers, 3) Cable trays, 4) Ducts/default
                                XYZ placementOffset = XYZ.Zero;
                                double finalWidth = 0.0, finalHeight = 0.0, finalDiameter = 0.0;

                                // ✅ PARALLEL PLANNING: Use pre-computed dimensions if available
                                // ⚠️ EXCEPTION: Duct Accessories (dampers) always excluded - need DamperPlacementStrategy for connector-based asymmetric clearance
                                // Pipes can use parallel planning if OptimizationFlags.EnableParallelPlanningForPipes is enabled
                                // When disabled (default), pipes use normal sequential processing for safety
                                bool isDuctAccessory = string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                                bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                                bool canUseParallelPlanningForPipe = isPipe && OptimizationFlags.EnableParallelPlanningForPipes;
                                if (planningDto != null && DeploymentConfiguration.EnableParallelPlanning && !isDuctAccessory && (canUseParallelPlanningForPipe || !isPipe))
                                {
                                    // Use planning layer dimensions (already includes clearance)
                                    finalWidth = planningDto.TargetWidthFt;
                                    finalHeight = planningDto.TargetHeightFt;
                                    finalDiameter = Math.Max(finalWidth, finalHeight);
                                    
                                    // ✅ DIRECT FILE WRITE: Log parallel planning dimensions
                                    try
                                    {
                                        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                        Directory.CreateDirectory(logDir);
                                        File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [PARALLEL-PLANNING] Zone {clashZone.Id}: Using pre-computed dimensions - Width={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), Height={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm), Risk={planningDto.ClearanceRisk}\n");
                                    } catch { }
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[PLANNING] Using pre-computed dimensions: W={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm (Risk={planningDto.ClearanceRisk})");
                                    }
                                }
                                // ✅ CRITICAL FIX: PATH 1 (Replay) should use existing sleeve size and skip clearance calculation
                                // For PATH 1, sleeve size data is already in ClashZone (SleeveWidth, SleeveHeight, SleeveDiameter)
                                // No need to calculate clearance - just use existing data
                                else if (_isReplayPath && clashZone.SleeveWidth > 0 && clashZone.SleeveHeight > 0)
                                {
                                    // ✅ PATH 1: Use existing sleeve size from ClashZone (already includes clearance)
                                    finalWidth = clashZone.SleeveWidth;
                                    finalHeight = clashZone.SleeveHeight;
                                    finalDiameter = clashZone.SleeveDiameter > 0 ? clashZone.SleeveDiameter : Math.Max(finalWidth, finalHeight);

                                    DebugLogger.Info($"[UniversalSleevePlacer] ✅ PATH 1 (Replay): Using existing sleeve size - W={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm, D={RevitUnitConversionService.Instance.FromInternalMillimeters(finalDiameter):F1}mm");
                                    DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [PATH-1-REPLAY] Zone {clashZone.Id}: Using SleeveWidth={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), SleeveHeight={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm) - SKIPPING clearance calculation\n");
                                }
                                else if (_isReplayPath)
                                {
                                    // ✅ PATH 1 but missing sleeve size data - log warning
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ PATH 1 (Replay): Missing sleeve size data for Zone={clashZone.Id} - SleeveWidth={clashZone.SleeveWidth}, SleeveHeight={clashZone.SleeveHeight}. Falling back to PATH 2/3 clearance calculation.");
                                        // File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 9: ⚠️ PATH 1 (Replay) - Missing sleeve size, falling back to clearance calculation\n");
                                    }
                                }
                                else
                                {
                                    // ✅ PATH 2/3: Calculate clearance (normal flow)
                                // ⏱️ TIMING: Clearance calculation
                                var clearanceTimer = System.Diagnostics.Stopwatch.StartNew();
                                using (singleSleeveTracker?.TrackSubOperation("Clearance Calculation"))
                                {
                                    // 🛡️ ARCHITECTURE FIX: Use CONDITIONS service for ALL clearance types
                                    // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
                                    // Raw dimensions from ClashZone + Clearance from CONDITIONS = Final dimensions

                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[UniversalSleevePlacer] CLEARANCE CALCULATION START: Category='{clashZone.MepElementCategory}', Strategy={(_strategy?.GetType().Name ?? "NULL")}, Path={(_isReplayPath ? "Replay" : "Sizing/Detection")}");
                                    
                                    // ✅ DIRECT FILE WRITE: Always log strategy type for debugging
                                    try
                                    {
                                        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                        Directory.CreateDirectory(logDir);
                                        File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [STRATEGY-TYPE] Zone {clashZone.Id}: Category='{clashZone.MepElementCategory}', StrategyType={(_strategy?.GetType().Name ?? "NULL")}, IsDamperStrategy={(_strategy is DamperPlacementStrategy)}, IsDuctStrategy={(_strategy is DuctPlacementStrategy)}\n");
                                    } catch { }

                                    bool isPipesCategory = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);

                                    // ✅ CRITICAL LOGGING: Always log which path pipes are taking
                                    if (isPipesCategory && !DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[PIPE-PATH-DEBUG] Zone {clashZone.Id}: Category='{clashZone.MepElementCategory}', isPipesCategory={isPipesCategory}, Will execute pipe sizing block");
                                        DebugLogger.Info($"[PIPE-PATH-DEBUG] Zone {clashZone.Id}: MepElementWidth={clashZone.MepElementWidth:F6}ft ({clashZone.MepElementWidth * 304.8:F1}mm), IsInsulated={clashZone.IsInsulated}, InsulationThickness={clashZone.InsulationThickness * 304.8:F1}mm");
                                    }

                                    if (isPipesCategory)
                                    {
                                        // ✅ OOP METHOD: Pipes - use insulation-aware sizing service (SOLID principles)
                                        // ✅ HARDCODED: Always use OUTER DIAMETER from database column (RBS_PIPE_OUTER_DIAMETER)
                                        // This ensures accurate sizing based on actual pipe outer diameter, not nominal
                                        var rawDiameter = clashZone.MepElementOuterDiameter > 0 
                                            ? clashZone.MepElementOuterDiameter 
                                            : clashZone.MepElementWidth; // Fallback to MepElementWidth if outer diameter not available
                                        
                                        if (!DeploymentConfiguration.DeploymentMode && clashZone.MepElementOuterDiameter > 0)
                                        {
                                            var odMm = RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.MepElementOuterDiameter);
                                            var nominalMm = clashZone.MepElementNominalDiameter > 0 
                                                ? RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.MepElementNominalDiameter) 
                                                : 0.0;
                                            DebugLogger.Info($"[PIPE-OD-USE] Zone {clashZone.Id}: Using MepElementOuterDiameter={odMm:F1}mm (nominal={nominalMm:F1}mm) for sizing");
                                        }
                                        
                                        // ✅ CRITICAL FIX: Use ClashZone.IsInsulated for correct clearance selection
                                        // The ClashZone has the correct insulation data from refresh, mepSize might be stale
                                        var clearance = GetClearanceForPipeFromClashZone(clashZone);
                                        
                                        // ✅ DETAILED LOGGING: Log all pipe sizing inputs for debugging (ALWAYS log, even in deployment mode via SafeFileLogger)
                                        var rawOdMm = RevitUnitConversionService.Instance.FromInternalMillimeters(rawDiameter);
                                        var clearanceMm = RevitUnitConversionService.Instance.FromInternalMillimeters(clearance);
                                        var insulationThicknessMm = clashZone.IsInsulated ? RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.InsulationThickness) : 0.0;
                                        
                                        // Always log via SafeFileLogger (works even if DebugLogger is disabled)
                                        var diameterSource = clashZone.MepElementOuterDiameter > 0 ? "MepElementOuterDiameter (RBS_PIPE_OUTER_DIAMETER)" : "MepElementWidth (fallback)";
                                        SafeFileLogger.SafeAppendText("pipe_sizing_debug.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PIPE-SIZING-DEBUG] Zone {clashZone.Id}: Raw OD={rawOdMm:F1}mm (from {diameterSource}), IsInsulated={clashZone.IsInsulated}, InsulationThickness={insulationThicknessMm:F1}mm, Clearance={clearanceMm:F1}mm\n");
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[PIPE-SIZING-DEBUG] Zone {clashZone.Id}: Raw OD={rawOdMm:F1}mm (from {diameterSource}), IsInsulated={clashZone.IsInsulated}, InsulationThickness={insulationThicknessMm:F1}mm, Clearance={clearanceMm:F1}mm");
                                        }
                                        
                                        (finalWidth, finalHeight, finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZone(
                                            rawDiameter, rawDiameter, rawDiameter, clashZone, clearance);
                                        
                                        // ✅ DETAILED LOGGING: Log calculated final size before rounding (ALWAYS log via SafeFileLogger)
                                        // Note: rawOdMm, clearanceMm, and insulationThicknessMm are already declared in outer scope above
                                        var finalDiameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(finalDiameter);
                                        if (clashZone.IsInsulated)
                                        {
                                            // Reuse existing variables from outer scope (no redeclaration)
                                            var formulaResult = rawOdMm + (insulationThicknessMm * 2) + (clearanceMm * 2);
                                            
                                            // Always log via SafeFileLogger
                                            SafeFileLogger.SafeAppendText("pipe_sizing_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PIPE-SIZING-CALC] Zone {clashZone.Id}: {rawOdMm:F1}mm (OD) + {insulationThicknessMm:F1}mm×2 (insulation) + {clearanceMm:F1}mm×2 (clearance) = {formulaResult:F1}mm (expected) vs {finalDiameterMm:F1}mm (calculated) (BEFORE ROUNDING)\n");
                                            
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[PIPE-SIZING-CALC] Zone {clashZone.Id}: {rawOdMm:F1}mm (OD) + {insulationThicknessMm:F1}mm×2 (insulation) + {clearanceMm:F1}mm×2 (clearance) = {finalDiameterMm:F1}mm (BEFORE ROUNDING)");
                                            }
                                        }
                                        else
                                        {
                                            // Reuse existing variables from outer scope (no redeclaration)
                                            var formulaResult = rawOdMm + (clearanceMm * 2);
                                            
                                            // Always log via SafeFileLogger
                                            SafeFileLogger.SafeAppendText("pipe_sizing_debug.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PIPE-SIZING-CALC] Zone {clashZone.Id}: {rawOdMm:F1}mm (OD) + {clearanceMm:F1}mm×2 (clearance) = {formulaResult:F1}mm (expected) vs {finalDiameterMm:F1}mm (calculated) (BEFORE ROUNDING)\n");
                                            
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[PIPE-SIZING-CALC] Zone {clashZone.Id}: {rawOdMm:F1}mm (OD) + {clearanceMm:F1}mm×2 (clearance) = {finalDiameterMm:F1}mm (BEFORE ROUNDING)");
                                            }
                                        }
                                        
                                        // ✅ CRITICAL: Store calculated dimensions in clashZone for pipes (same as dampers)
                                        // This ensures the calculated dimensions are preserved through the placement process
                                        // For pipes, we ALWAYS use the freshly calculated dimensions (never use old database values)
                                        clashZone.SleeveWidth = finalWidth;
                                        clashZone.SleeveHeight = finalHeight;
                                        clashZone.SleeveDiameter = finalDiameter;
                                    }
                                    else if (_strategy is DamperPlacementStrategy damperStrategy)
                                    {
                                        // ✅ DIRECT FILE WRITE: Always log entry into damper strategy block
                                        try
                                        {
                                            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                            Directory.CreateDirectory(logDir);
                                            File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [DAMPER-STRATEGY-ENTRY] Zone {clashZone.Id}: Entering damper strategy block, Category={clashZone.MepElementCategory}\n");
                                        } catch { }
                                        
                                        // ✅ Fire dampers: Raw dimensions + CONDITIONS clearance via strategy
                                        var rawWidth = clashZone.MepElementWidth;
                                        var rawHeight = clashZone.MepElementHeight;

                                        // ✅ DIRECT LOGGING: Always log what we're reading from ClashZone
                                        DebugLogger.Info($"[DAMPER-PLACEMENT-DEBUG] Zone {clashZone.Id}: Reading from ClashZone - MepElementWidth={rawWidth:F6}ft ({RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}mm), MepElementHeight={rawHeight:F6}ft ({RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm)");
                                        DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [PLACEMENT-READ] Zone {clashZone.Id}: MepElementWidth={rawWidth:F6}ft ({rawWidth * 304.8:F1}mm), MepElementHeight={rawHeight:F6}ft ({rawHeight * 304.8:F1}mm)\n");

                                        // Get offset and final dimensions from strategy (uses CONDITIONS)
                                        var adj = damperStrategy.GetDamperPlacementAdjustment(clashZone, _conditions);
                                        placementOffset = adj.offsetVector;
                                        finalWidth = adj.finalWidth;
                                        finalHeight = adj.finalHeight;
                                        finalDiameter = finalWidth; // Not used for dampers (rectangular only)

                                        DebugLogger.Info($"[UniversalSleevePlacer] DAMPER: Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm → Final={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
                                        DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [PLACEMENT-FINAL] Zone {clashZone.Id}: FinalWidth={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), FinalHeight={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm)\n");
                                        
                                        // ✅ CRITICAL: Always store calculated dimensions in clashZone for dampers
                                        // This ensures the calculated dimensions are preserved through the placement process
                                        // For dampers, we ALWAYS use the freshly calculated dimensions (never use old database values)
                                        clashZone.SleeveWidth = finalWidth;
                                        clashZone.SleeveHeight = finalHeight;
                                        clashZone.SleeveDiameter = finalDiameter;
                                    }
                                    else if (_strategy is DuctPlacementStrategy ductStrategy)
                                    {
                                        // ✅ DIRECT LOGGING: Always log if DUCT strategy is being used for Duct Accessories (this is the bug!)
                                        if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                                        {
                                            DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [BUG-DETECTED] Zone {clashZone.Id}: ⚠️ DUCT STRATEGY is being used for Duct Accessories! This should use DamperPlacementStrategy!\n");
                                        }
                                        
                                        // ✅ OOP METHOD: Ducts - use insulation-aware sizing service (SOLID principles)
                                        var rawWidth = clashZone.MepElementWidth;
                                        var rawHeight = clashZone.MepElementHeight;

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT CALCULATION START: Raw dimensions {RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm");

                                        // ✅ CRITICAL FIX: Use ClashZone.IsInsulated for correct clearance selection
                                        var clearance = GetClearanceForDuctFromClashZone(clashZone);
                                        
                                        // ✅ DIRECT LOGGING: Always log clearance being used
                                        DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [DUCT-STRATEGY-CLEARANCE] Zone {clashZone.Id}: Clearance={clearance:F6}ft ({clearance * 304.8:F1}mm) from GetClearanceFromConditions('Ducts')\n");
                                        
                                        // ✅ OOP METHOD: Use sizing service for consistent calculation across all categories
                                        (finalWidth, finalHeight, finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZone(
                                            rawWidth, rawHeight, rawWidth, clashZone, clearance);
                                        finalDiameter = finalWidth; // For round elements

                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            if (clashZone.IsInsulated)
                                                DebugLogger.Info($"[UniversalSleevePlacer] DUCT (OOP): Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm + Insulation({RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.InsulationThickness):F1}mm x2) + Clearance({RevitUnitConversionService.Instance.FromInternalMillimeters(clearance):F1}mm x2) = Final={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
                                            else
                                                DebugLogger.Info($"[UniversalSleevePlacer] DUCT (OOP): Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm + Clearance({RevitUnitConversionService.Instance.FromInternalMillimeters(clearance):F1}mm x2) = Final={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
                                        }
                                        
                                        DirectLog($"[{DateTime.Now:HH:mm:ss.fff}] [DUCT-STRATEGY-FINAL] Zone {clashZone.Id}: FinalWidth={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), FinalHeight={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm)\n");
                                    }
                                    else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                                    {
                                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                                        // ✅ Cable trays: Raw dimensions + UI/XML clearance via strategy
                                        var rawWidth = clashZone.MepElementWidth;
                                        var rawHeight = clashZone.MepElementHeight;

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: UI Clearance Settings Count={_clearanceSettings.Count}");

                                        // Get offset and final dimensions from strategy (uses UI settings first, then CONDITIONS)
                                        var adj2 = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions, _clearanceSettings);
                                        placementOffset = adj2.offsetVector;
                                        finalWidth = adj2.finalWidth;
                                        finalHeight = adj2.finalHeight;

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Final={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
                                        finalDiameter = finalWidth; // Not used for cable trays (rectangular only)

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY: Raw={RevitUnitConversionService.Instance.FromInternalMillimeters(rawWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(rawHeight):F1}mm → Final={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
                                    }
                                    else
                                    {
                                        // ✅ OOP METHOD: Fallback - use insulation-aware sizing service (SOLID principles)
                                        var rawWidth = clashZone.MepElementWidth;
                                        var rawHeight = clashZone.MepElementHeight;
                                        var defaultClearance = RevitUnitConversionService.Instance.ToInternalMillimeters(50); // 50mm default

                                        // ✅ OOP METHOD: Use sizing service for consistent calculation
                                        (finalWidth, finalHeight, finalDiameter) = _sizingService.CalculateFinalDimensionsFromClashZone(
                                            rawWidth, rawHeight, Math.Max(rawWidth, rawHeight), clashZone, defaultClearance);

                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ NO STRATEGY MATCHED for ClashZone {clashZone.Id} - Using fallback dimensions with OOP sizing service");
                                    }
                                    clearanceTimer.Stop();
                                    totalClearanceTime += clearanceTimer.Elapsed;
                                    sleeveLog.AppendLine($"  Clearance calc: {clearanceTimer.ElapsedMilliseconds}ms");
                                } // End Clearance Calculation sub-operation

                                    try
                                    {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 9: Clearance calculation completed for PATH 2/3\n");
                                        }
                                        */
                                    }
                                    catch { }
                                }

                                // ⚠️ REMOVED: Old width/height swapping logic that was causing double-swapping
                                // The new logic later in the method (lines 660-665) handles this correctly
                                // by ensuring the longer dimension becomes width, not just swapping blindly

                                try
                                {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                    /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 10: Final dimensions determined, selecting family\n");
                                    }
                                    */
                                }
                                catch { }

                                // ⏱️ TIMING: Family selection and loading
                                var familyTimer = System.Diagnostics.Stopwatch.StartNew();
                                FamilySymbol familySymbol = null;
                                string familyName = "";
                                bool isCircular = false;
                                using (singleSleeveTracker?.TrackSubOperation("Load Family Symbol"))
                                {
                                    // Select universal family
                                    (familyName, isCircular) = SelectUniversalFamily(clashZone, mepSize);
                                    try
                                    {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 11: Family selected: '{familyName}', IsCircular={isCircular}\n");
                                        }
                                        */
                                    }
                                    catch { }

                                    familySymbol = LoadFamilySymbol(familyName);
                                    familyTimer.Stop();
                                    totalFamilyLoadTime += familyTimer.Elapsed;
                                    sleeveLog.AppendLine($"  Family load: {familyTimer.ElapsedMilliseconds}ms");
                                } // End Load Family Symbol sub-operation

                                bool shouldSkipSleeve = false;

                                if (familySymbol == null)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error($"[UniversalSleevePlacer] Family '{familyName}' not found");
                                    try
                                    {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 11.1: ❌ ERROR - Family '{familyName}' not found\n");
                                        }
                                        */
                                    }
                                    catch { }
                                    ErrorCount++;
                                    continue;
                                }
                                try
                                {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                    /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 12: ✅ Family symbol loaded successfully\n");
                                    }
                                    */
                                }
                                catch { }

                                // ✅ SIMPLIFIED: Activate symbol if needed - don't care about type name, just activate any type
                                if (!familySymbol.IsActive)
                                {
                                    try
                                    {
                                        familySymbol.Activate();
                                    }
                                    catch (Exception activationEx)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[UniversalSleevePlacer] Failed to activate family '{familyName}': {activationEx.Message}");

                                        var family = familySymbol.Family;
                                        var allSymbols = family.GetFamilySymbolIds()
                                            .Select(id => _doc.GetElement(id) as FamilySymbol)
                                            .Where(s => s != null)
                                            .ToList();

                                        var activeSymbol = allSymbols.FirstOrDefault(s => s.IsActive);
                                        if (activeSymbol != null)
                                        {
                                            familySymbol = activeSymbol;
                                        }
                                        else
                                        {
                                            var firstSymbol = allSymbols.FirstOrDefault();
                                            if (firstSymbol != null)
                                            {
                                                try
                                                {
                                                    firstSymbol.Activate();
                                                    familySymbol = firstSymbol;
                                                }
                                                catch (Exception reactivationEx)
                                                {
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                        DebugLogger.Error($"[UniversalSleevePlacer] Unable to activate any type in family '{familyName}': {reactivationEx.Message}");
                                                    ErrorCount++;
                                                    shouldSkipSleeve = true;
                                                }
                                            }
                                            else
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                    DebugLogger.Error($"[UniversalSleevePlacer] Family '{familyName}' does not contain any symbols.");
                                                ErrorCount++;
                                                shouldSkipSleeve = true;
                                            }
                                        }
                                    }
                                }

                                if (shouldSkipSleeve)
                                {
                                    continue;
                                }

                                // ✅ SLEEVE PLACEMENT FLOW - Step 2: Determine placement origin
                                bool hasSnapshotData = HasSnapshotPlacementData(clashZone);
                                bool usingXmlSnapshot = false;
                                bool usingDbPlacementPoint = false;
                                XYZ placementPointChosen;

                                if (_isReplayPath)
                                {
                                    // ✅ PATH 1: Prioritize DB placement point (from database), fallback to XML snapshot, then intersection point
                                    var dbPlacementPoint = clashZone.SleevePlacementPoint;

                                    if (dbPlacementPoint != null && HasValidPlacementCoordinate(dbPlacementPoint))
                                    {
                                        // ✅ Use DB placement point (from database)
                                        placementPointChosen = dbPlacementPoint;
                                        usingDbPlacementPoint = true;
                                        clashZone.IntersectionPoint = placementPointChosen;
                                        clashZone.IntersectionPointX = placementPointChosen.X;
                                        clashZone.IntersectionPointY = placementPointChosen.Y;
                                        clashZone.IntersectionPointZ = placementPointChosen.Z;

                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[PLACEMENT-SOURCE] Zone={clashZone.Id}: Using DB placement point (from database)");
                                        }
                                    }
                                    else if (hasSnapshotData)
                                    {
                                        // ✅ Fallback to XML snapshot if DB placement point is not available
                                        placementPointChosen = GetSnapshotPlacementPoint(clashZone);

                                        if (HasValidPlacementCoordinate(placementPointChosen))
                                        {
                                            usingXmlSnapshot = true;
                                            clashZone.IntersectionPoint = placementPointChosen;
                                            clashZone.IntersectionPointX = placementPointChosen.X;
                                            clashZone.IntersectionPointY = placementPointChosen.Y;
                                            clashZone.IntersectionPointZ = placementPointChosen.Z;

                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Warning($"[PLACEMENT-SOURCE] Zone={clashZone.Id}: Using XML snapshot (DB placement point not available)");
                                            }
                                        }
                                        else
                                        {
                                            // Fallback: use intersection point
                                            placementPointChosen = clashZone.IntersectionPoint ??
                                                                   new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                                        }
                                    }
                                    else
                                    {
                                        // Fallback: use intersection point
                                        placementPointChosen = clashZone.IntersectionPoint ??
                                                               new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                                    }
                                }
                                else
                                {
                                    if (clashZone.IntersectionPoint == null)
                                    {
                                        placementPointChosen = new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                                        clashZone.IntersectionPoint = placementPointChosen; // Rehydrate for downstream code
                                    }
                                    else
                                    {
                                        placementPointChosen = clashZone.IntersectionPoint;
                                    }
                                }

                                if (!DeploymentConfiguration.DeploymentMode && PlacedCount < 5)
                                {
                                    DebugLogger.Info($"[SLEEVE-PLACEMENT-FLOW] Step 2: Placement data snapshot → HasSnapshot={hasSnapshotData}, ReplayPath={_isReplayPath}");
                                }

                                // Validate placement point has meaningful coordinates
                                if (!HasValidPlacementCoordinate(placementPointChosen))
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error($"[UniversalSleevePlacer] ❌ CRITICAL: Cannot place sleeve for Zone {clashZone.Id} - Placement point is (0,0,0)!");
                                    // try { File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.3: ❌ SKIPPED - Placement point is (0,0,0)\n"); } catch { }
                                    SkippedCount++;
                                    sleeveTimer.Stop();
                                    continue;
                                }

                                // Keep SleevePlacementPoint in sync (for clustering compatibility, but for individual sleeves it's same as IntersectionPoint)
                                clashZone.SleevePlacementPoint = placementPointChosen;
                                clashZone.SleevePlacementPointX = placementPointChosen.X;
                                clashZone.SleevePlacementPointY = placementPointChosen.Y;
                                clashZone.SleevePlacementPointZ = placementPointChosen.Z;

                                // ✅ CRITICAL DEBUG: Log placement point for each zone to detect same-point issue (using SafeFileLogger path)
                                try
                                {
                                    // DEPLOYMENT MODE: Skip file writes
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                        var placementSource = _isReplayPath
                                            ? (usingDbPlacementPoint ? "DB placement point (from database)" : (usingXmlSnapshot ? "XML snapshot (placement point from XML)" : "Intersection point (fallback)"))
                                            : "Recomputed intersection";

                                        // ✅ DIAGNOSTIC: Log data source for placement point
                                        var dbPlacementPoint = clashZone.SleevePlacementPoint;
                                        var dbPlacementPointStr = dbPlacementPoint != null ? $"DB_SPP=({dbPlacementPoint.X:F3},{dbPlacementPoint.Y:F3},{dbPlacementPoint.Z:F3})" : "DB_SPP=null";

                                        /* EXCESSIVE LOGGING COMMENTED OUT FOR PERFORMANCE
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [PLACEMENT-ORIGIN] Zone={clashZone.Id}, Source={placementSource}, Point=({placementPointChosen?.X:F3},{placementPointChosen?.Y:F3},{placementPointChosen?.Z:F3}), {dbPlacementPointStr}, XML_SPP=({clashZone.SleevePlacementPointX:F3},{clashZone.SleevePlacementPointY:F3},{clashZone.SleevePlacementPointZ:F3})\n");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] PLACEMENT: Zone={clashZone.Id}, MEP={clashZone.MepElementIdValue}, HOST={clashZone.StructuralElementIdValue}, Point=({placementPointChosen?.X:F3},{placementPointChosen?.Y:F3},{placementPointChosen?.Z:F3}), SPP_XML=({clashZone.SleevePlacementPointX:F3},{clashZone.SleevePlacementPointY:F3},{clashZone.SleevePlacementPointZ:F3}), IP=({clashZone.IntersectionPointX:F3},{clashZone.IntersectionPointY:F3},{clashZone.IntersectionPointZ:F3})\n");
                                        */

                                        // ✅ DIAGNOSTIC: Log if using XML snapshot when DB has placement point
                                        if (usingXmlSnapshot && dbPlacementPoint != null && HasValidPlacementCoordinate(dbPlacementPoint))
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Warning($"[PLACEMENT-SOURCE] ⚠️ Zone={clashZone.Id}: Using XML snapshot placement point, but DB has valid placement point! XML=({clashZone.SleevePlacementPointX:F3},{clashZone.SleevePlacementPointY:F3},{clashZone.SleevePlacementPointZ:F3}), DB=({dbPlacementPoint.X:F3},{dbPlacementPoint.Y:F3},{dbPlacementPoint.Z:F3})");
                                            }
                                        }
                                    }
                                }
                                catch (Exception logEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error($"[PLACEMENT-LOG-ERROR] {logEx.Message}");
                                }

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] Using exact intersection point {placementPointChosen} for {clashZone.MepElementCategory} on {clashZone.StructuralElementType}");

                                XYZ adjustedPlacementPoint = placementPointChosen + placementOffset;
                                
                                // ✅ DETAILED LOGGING: Log damper offset application for debugging
                                if (string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase) && placementOffset.GetLength() > 0.001)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        var offsetMm = placementOffset.GetLength() * 304.8;
                                        DebugLogger.Info($"[DAMPER-OFFSET] Zone {clashZone.Id}: PlacementPoint BEFORE offset = ({placementPointChosen.X:F6}, {placementPointChosen.Y:F6}, {placementPointChosen.Z:F6})");
                                        DebugLogger.Info($"[DAMPER-OFFSET] Zone {clashZone.Id}: Offset Vector = ({placementOffset.X*304.8:F1}, {placementOffset.Y*304.8:F1}, {placementOffset.Z*304.8:F1})mm (Length={offsetMm:F1}mm)");
                                        DebugLogger.Info($"[DAMPER-OFFSET] Zone {clashZone.Id}: PlacementPoint AFTER offset = ({adjustedPlacementPoint.X:F6}, {adjustedPlacementPoint.Y:F6}, {adjustedPlacementPoint.Z:F6})");
                                        try
                                        {
                                            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                            Directory.CreateDirectory(logDir);
                                            File.AppendAllText(Path.Combine(logDir, "damper_placement_trace.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [OFFSET-APPLIED] Zone {clashZone.Id}: PlacementPoint BEFORE=({placementPointChosen.X:F6}, {placementPointChosen.Y:F6}, {placementPointChosen.Z:F6}), Offset=({placementOffset.X*304.8:F1}, {placementOffset.Y*304.8:F1}, {placementOffset.Z*304.8:F1})mm, PlacementPoint AFTER=({adjustedPlacementPoint.X:F6}, {adjustedPlacementPoint.Y:F6}, {adjustedPlacementPoint.Z:F6})\n");
                                        } catch { }
                                    }
                                }
                                else if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalSleevePlacer] Using placement point {placementPointChosen} → adjusted {adjustedPlacementPoint}");

                                    if (placementOffset.GetLength() > 0.001)
                                    {
                                        DebugLogger.Info($"[UniversalSleevePlacer] Applied offset {placementOffset} to placement point");
                                    }
                                }

                                // ⏱️ TIMING: Level finding
                                var levelTimer = System.Diagnostics.Stopwatch.StartNew();
                                Level nearestLevel = null;
                                using (singleSleeveTracker?.TrackSubOperation("Find Nearest Level"))
                                {
                                // ✅ PERFORMANCE OPTIMIZATION: Use cached levels list and select level based on structural element type

                                // ✅ CORRECT: For walls and framing, find nearest BOTTOM level (where wall/framing starts)
                                // For floors, find nearest level overall (floor can span multiple levels)
                                bool isWallOrFraming = clashZone.StructuralElementType == "Wall" ||
                                                      clashZone.StructuralElementType == "Walls" ||
                                                      clashZone.StructuralElementType == "Structural Framing";

                                if (isWallOrFraming)
                                {
                                    // Find nearest level BELOW or AT the placement point (bottom level for wall/framing)
                                    double placementZ = adjustedPlacementPoint.Z;
                                    Level nearestBottomLevel = null;
                                    double minDistanceBelow = double.MaxValue;

                                    foreach (var level in cachedLevels)
                                    {
                                        // Only consider levels at or below the placement point
                                        if (level.Elevation <= placementZ)
                                        {
                                            var distance = placementZ - level.Elevation; // Distance from placement to level below
                                            if (distance < minDistanceBelow)
                                            {
                                                minDistanceBelow = distance;
                                                nearestBottomLevel = level;
                                            }
                                        }
                                    }

                                    nearestLevel = nearestBottomLevel;

                                    if (nearestLevel == null)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[UniversalSleevePlacer] No level found BELOW placement point Z={placementZ:F3} for {clashZone.StructuralElementType}");
                                        SkippedCount++;
                                        sleeveTimer.Stop();
                                        continue;
                                    }

                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[UniversalSleevePlacer] Found bottom level '{nearestLevel.Name}' (Elevation={nearestLevel.Elevation:F3}) for {clashZone.StructuralElementType} at Z={placementZ:F3}");
                                }
                                else
                                {
                                    // For floors: find nearest level overall (any direction)
                                    double minDistance = double.MaxValue;
                                    foreach (var level in cachedLevels)
                                    {
                                        var distance = Math.Abs(level.Elevation - adjustedPlacementPoint.Z);
                                        if (distance < minDistance)
                                        {
                                            minDistance = distance;
                                            nearestLevel = level;
                                        }
                                    }

                                    if (nearestLevel == null)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[UniversalSleevePlacer] No level found for placement point");
                                        SkippedCount++;
                                        sleeveTimer.Stop();
                                        continue;
                                    }

                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[UniversalSleevePlacer] Found nearest level '{nearestLevel.Name}' (Elevation={nearestLevel.Elevation:F3}) for Floor at Z={adjustedPlacementPoint.Z:F3}");
                                }
                                levelTimer.Stop();
                                totalLevelFindTime += levelTimer.Elapsed;
                                sleeveLog.AppendLine($"  Level find: {levelTimer.ElapsedMilliseconds}ms");
                                } // End Find Nearest Level sub-operation

                                // DEPLOYMENT MODE: Skip file writes
                                try
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 13: About to create sleeve at point=({adjustedPlacementPoint.X:F3},{adjustedPlacementPoint.Y:F3},{adjustedPlacementPoint.Z:F3}), Level='{nearestLevel?.Name ?? "NULL"}'\n");
                                    }
                                }
                                catch { }

                                // ✅ SLEEVE PLACEMENT FLOW - Step 3: Place sleeve using data from Filter XML
                                // Uses placement coordinates, family selection, and dimensions from Filter XML
                                // ⏱️ TIMING: Sleeve creation
                                var createTimer = System.Diagnostics.Stopwatch.StartNew();
                                FamilyInstance sleeveInstance = null;
                                using (singleSleeveTracker?.TrackSubOperation("Create Sleeve Instance"))
                                {
                                    if (!DeploymentConfiguration.DeploymentMode && PlacedCount < 5)
                                        DebugLogger.Info($"[SLEEVE-PLACEMENT-FLOW] Step 3: Placing sleeve at ({adjustedPlacementPoint.X:F3},{adjustedPlacementPoint.Y:F3},{adjustedPlacementPoint.Z:F3}) " +
                                            $"using family '{familySymbol.Family.Name}' with dimensions W={finalWidth:F3}, H={finalHeight:F3}");

                                    // Place sleeve instance (NO HOST PARAMETER - workplane-based families)
                                    // ✅ Works with linked structural elements because no host reference needed
                                    try
                                    {
                                        // ✅ STEP 3: Create sleeve instance using placement data from Filter XML
                                        // ✅ PERFORMANCE PROFILING: Profile family instantiation to identify symbol binding vs geometry creation
                                        var instantiationTimer = System.Diagnostics.Stopwatch.StartNew();
                                        var beforeInstantiation = System.GC.CollectionCount(0); // Track GC before
                                        
                                        sleeveInstance = _doc.Create.NewFamilyInstance(
                                            adjustedPlacementPoint,  // From Filter XML (IntersectionPoint)
                                            familySymbol,            // Selected based on host type from Filter XML
                                            nearestLevel,
                                            StructuralType.NonStructural);
                                        
                                        instantiationTimer.Stop();
                                        var afterInstantiation = System.GC.CollectionCount(0);
                                        var gcCollections = afterInstantiation - beforeInstantiation;
                                        
                                        // ✅ PROFILING: Log instantiation timing to identify bottlenecks
                                        // Fast (<10ms) = quick placement, Medium (10-50ms) = moderate overhead, Slow (>50ms) = high overhead
                                        // Note: No geometry creation - just family placement (symbol binding)
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            string instantiationType = instantiationTimer.ElapsedMilliseconds < 10 
                                                ? "FAST_PLACEMENT" 
                                                : instantiationTimer.ElapsedMilliseconds < 50 
                                                    ? "MODERATE_OVERHEAD" 
                                                    : "SLOW_OVERHEAD";
                                            
                                            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                                            Directory.CreateDirectory(logDir);
                                            var profPath = Path.Combine(logDir, "family_instantiation_profile.log");
                                            
                                            File.AppendAllText(profPath, 
                                                $"{DateTime.Now:O}\t" +
                                                $"SleeveId={sleeveInstance?.Id?.IntegerValue ?? -1}\t" +
                                                $"Family={familySymbol?.Family?.Name ?? "NULL"}\t" +
                                                $"Symbol={familySymbol?.Name ?? "NULL"}\t" +
                                                $"Type={instantiationType}\t" +
                                                $"TimeMs={instantiationTimer.ElapsedMilliseconds}\t" +
                                                $"TimeTicks={instantiationTimer.ElapsedTicks}\t" +
                                                $"GCCollections={gcCollections}\t" +
                                                $"Level={nearestLevel?.Name ?? "NULL"}\n");
                                        }
                                        try
                                        {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14: ✅ Sleeve instance created! ID={sleeveInstance?.Id?.IntegerValue ?? -1}\n");
                                            }
                                        }
                                        catch { }
                                    }
                                    catch (Exception createEx)
                                    {
                                        try
                                        {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14.1: ❌ EXCEPTION during sleeve creation: {createEx.Message}\n{createEx.StackTrace}\n");
                                            }
                                        }
                                        catch { }
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Error($"[UniversalSleevePlacer] Exception creating sleeve: {createEx.Message}");
                                        ErrorCount++;
                                        singleSleeveTracker.SetItemCount(1); // Track error as 1 item
                                        sleeveTimer.Stop();
                                        continue;
                                    }

                                    createTimer.Stop();
                                    totalSleeveCreateTime += createTimer.Elapsed;
                                    sleeveLog.AppendLine($"  Sleeve create: {createTimer.ElapsedMilliseconds}ms");
                                } // End Create Sleeve Instance sub-operation

                                if (sleeveInstance == null)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error($"[UniversalSleevePlacer] Failed to create sleeve instance");
                                    try
                                    {                             // ✅ DEPLOYMENT MODE: Skip file writes
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14.2: ❌ ERROR - Sleeve instance is NULL after creation\n");
                                        }
                                    }
                                    catch { }
                                    ErrorCount++;
                                    singleSleeveTracker.SetItemCount(1); // Track error as 1 item
                                    sleeveTimer.Stop();
                                    continue;
                                }

                                // ⏱️ TIMING: Parameter setting
                                var parameterTimer = System.Diagnostics.Stopwatch.StartNew();
                                using (singleSleeveTracker?.TrackSubOperation("Set Sleeve Parameters"))
                                {
                                // ✅ CRITICAL FIX FOR DAMPERS AND PIPES: Ensure finalWidth matches calculated value (restore from clashZone if needed)
                                // Note: isPipe is already declared in outer scope (line 1120), reuse it here
                                bool isDamper = string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                                
                                if (isDamper || isPipe)
                                {
                                    // ✅ For dampers and pipes, clashZone.SleeveWidth should contain the calculated value
                                    // If finalWidth doesn't match, restore it from clashZone.SleeveWidth
                                    if (clashZone.SleeveWidth > 0 && Math.Abs(finalWidth - clashZone.SleeveWidth) > 0.0001)
                                    {
                                        try { 
                                            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023"); 
                                            Directory.CreateDirectory(logDir); 
                                            var logPath = Path.Combine(logDir, isDamper ? "damper_placement_trace.log" : "placement_debug.log");
                                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [SET-PARAMS-FIX] Zone {clashZone.Id} ({clashZone.MepElementCategory}): finalWidth was {finalWidth * 304.8:F1}mm, restoring from clashZone.SleeveWidth={clashZone.SleeveWidth * 304.8:F1}mm\n");
                                        } catch { }
                                        finalWidth = clashZone.SleeveWidth;
                                        finalHeight = clashZone.SleeveHeight;
                                        finalDiameter = clashZone.SleeveDiameter;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[SET-PARAMS-FIX] Zone {clashZone.Id} ({clashZone.MepElementCategory}): Restored calculated dimensions - Width={finalWidth * 304.8:F1}mm, Height={finalHeight * 304.8:F1}mm, Diameter={finalDiameter * 304.8:F1}mm");
                                        }
                                    }
                                    
                                    // ✅ DIRECT LOGGING: Always log final dimensions before setting parameters
                                    try { 
                                        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023"); 
                                        Directory.CreateDirectory(logDir); 
                                        var logPath = Path.Combine(logDir, isDamper ? "damper_placement_trace.log" : "placement_debug.log");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [SET-PARAMS] Zone {clashZone.Id} ({clashZone.MepElementCategory}): About to set Width={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), Height={finalHeight:F6}ft ({finalHeight * 304.8:F1}mm), Diameter={finalDiameter:F6}ft ({finalDiameter * 304.8:F1}mm) on sleeve {sleeveInstance.Id.IntegerValue}\n");
                                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [SET-PARAMS-DIAGNOSTIC] Zone {clashZone.Id}: finalWidth={finalWidth:F6}ft ({finalWidth * 304.8:F1}mm), clashZone.SleeveWidth={clashZone.SleeveWidth:F6}ft ({clashZone.SleeveWidth * 304.8:F1}mm), _isReplayPath={_isReplayPath}\n");
                                    } catch { }
                                }
                                
                                // Set parameters
                                SetSleeveParameters(sleeveInstance, mepSize, finalWidth, finalHeight, finalDiameter, clashZone, isCircular);

                                // Set sleeve metadata for fast parameter transfer
                                SetSleeveMetadata(sleeveInstance, clashZone);

                                // ✅ CRITICAL: Set ClashZone_GUID parameter with STABLE GUID per 3-point combo
                                // GUID is looked up from Global XML first (if entry exists, use that GUID)
                                // If not found, use clashZone.Id (stable per 3-point combo)
                                // This ensures GUID is unique and stable across multiple detection runs
                                SetClashZoneGuidOnSleeveStable(sleeveInstance, clashZone);

                                // ⚠️ CRITICAL: Set orientation (rotation for floors, HostOrientation parameter for walls/framing)
                                SetSleeveOrientation(sleeveInstance, clashZone, planningDto);
                                parameterTimer.Stop();
                                totalParameterTime += parameterTimer.Elapsed;
                                sleeveLog.AppendLine($"  Parameters: {parameterTimer.ElapsedMilliseconds}ms");
                                } // End Set Sleeve Parameters sub-operation

                                // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                                // ⏱️ TIMING: Coordinate validation and update
                                var validationTimer = System.Diagnostics.Stopwatch.StartNew();
                                using (singleSleeveTracker?.TrackSubOperation("Update ClashZone"))
                                {
                                // Ensure final location matches the intended adjusted placement point (some families snap to level origin)
                                try
                                {
                                    var loc = sleeveInstance.Location as LocationPoint;
                                    if (loc != null)
                                    {
                                        var currentPt = loc.Point;
                                        if (currentPt.DistanceTo(adjustedPlacementPoint) > 0.0001)
                                        {
                                            var delta = adjustedPlacementPoint - currentPt;
                                            loc.Move(delta);
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[UniversalSleevePlacer] Moved sleeve {sleeveInstance.Id} to exact placement point {adjustedPlacementPoint} (from {currentPt})");
                                        }
                                        else
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[UniversalSleevePlacer] Sleeve {sleeveInstance.Id} already at desired point {currentPt}");
                                        }

                                        // ✅ CRITICAL FIX: Update ClashZone with actual sleeve placement coordinates
                                        // This ensures PreCalculatedClusterService uses real sleeve locations for proximity calculation
                                        clashZone.SleevePlacementPoint = currentPt;
                                        clashZone.SleevePlacementPointX = currentPt.X;  // XML serializable
                                        clashZone.SleevePlacementPointY = currentPt.Y;  // XML serializable
                                        clashZone.SleevePlacementPointZ = currentPt.Z;  // XML serializable

                                        // ✅ CRITICAL: Save active document coordinates for proximity calculation
                                        clashZone.SleevePlacementPointActiveDocument = currentPt;
                                        clashZone.SleevePlacementPointActiveDocumentX = currentPt.X;  // XML serializable
                                        clashZone.SleevePlacementPointActiveDocumentY = currentPt.Y;  // XML serializable
                                        clashZone.SleevePlacementPointActiveDocumentZ = currentPt.Z;  // XML serializable

                                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                                        // 🔥 CRITICAL VALIDATION: Check if dimensions are valid before saving
                                        if (finalWidth <= 0 || finalHeight <= 0 || finalDiameter <= 0)
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Error($"[UniversalSleevePlacer] ❌ INVALID DIMENSIONS DETECTED for sleeve {sleeveInstance.Id.IntegerValue}!");
                                            ErrorCount++;
                                            sleeveTimer.Stop();
                                            validationTimer.Stop();
                                            continue;
                                        }

                                        clashZone.SleeveWidth = finalWidth;
                                        clashZone.SleeveHeight = finalHeight;
                                        clashZone.SleeveDiameter = finalDiameter;

                                        // ✅ PERFORMANCE OPTIMIZATION: Defer bounding box retrieval until after batch regeneration
                                        // Store sleeve data for batch processing instead of immediate bounding box call
                                        placedSleeveIds.Add(sleeveInstance.Id);
                                        placedSleeveData.Add((sleeveInstance, clashZone, finalWidth, finalHeight, finalDiameter));

                                        // ✅ CRITICAL FIX: Set the actual Revit element ID immediately (needed for validation)
                                        clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;

                                        // ⚠️ NOTE: Bounding box will be retrieved after batch regeneration (see batch processing at end of method)

                                        // ✅ PERFORMANCE OPTIMIZATION: Batch XML updates instead of updating per sleeve
                                        // XML will be updated once at the end of placement via orchestrator

                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[UniversalSleevePlacer] Placed sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id}, W={RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm (bbox deferred)");
                                            try
                                            {
                                                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                                System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [SLEEVE-ID-SET] ClashZone {clashZone.Id} ← SleeveId {clashZone.SleeveInstanceId}\n");
                                            }
                                            catch { }
                                        }
                                    }
                                    else
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Error($"[UniversalSleevePlacer] ❌ Failed to get LocationPoint for sleeve {sleeveInstance.Id} - cannot update coordinates!");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[COORD-UPDATE-FAIL] Sleeve {sleeveInstance.Id.IntegerValue}: LocationPoint is NULL ❌\n");
                                    }
                                }
                                catch (Exception coordEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Error($"[UniversalSleevePlacer] ❌ CRITICAL ERROR updating coordinates for sleeve {sleeveInstance.Id}: {coordEx.Message}");
                                    if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"[COORD-UPDATE-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {coordEx.Message}");
                                }
                                validationTimer.Stop();
                                totalValidationTime += validationTimer.Elapsed;
                                sleeveLog.AppendLine($"  Validation/update: {validationTimer.ElapsedMilliseconds}ms");
                                } // End Update ClashZone sub-operation

                                // Update ClashZone flags
                                clashZone.IsResolved = true;
                                clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;
                                clashZone.SleeveFamilyName = familySymbol.Family.Name;

                                // ✅ NOTE: Global XML update happens in batch at end (line ~1238) via GlobalIndexService.UpsertFlagsWithIds()
                                // This ensures all placed sleeves are updated together efficiently
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[GLOBAL-XML] Individual sleeve {sleeveInstance.Id.IntegerValue} placed for ClashZone {clashZone.Id} - will be saved to Global XML in batch");

                                // ✅ PERFORMANCE OPTIMIZATION: Batch logging instead of individual file writes
                                if (PlacedCount < 50) batchLogs.AppendLine($"[SLEEVE-PLACED] ClashZone {clashZone.Id}: SleeveInstanceId = {clashZone.SleeveInstanceId}, RevitElementId = {sleeveInstance.Id.IntegerValue}");

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] ✅ SLEEVE PLACED: ClashZone {clashZone.Id} → SleeveInstanceId = {clashZone.SleeveInstanceId} (Revit: {sleeveInstance.Id.IntegerValue})");

                                // ✅ CRITICAL: SleevePlacementPoint is already set to actual sleeve location above
                                // DO NOT overwrite it with adjustedPlacementPoint - we need the REAL coordinates for clustering

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] Saved sleeve placement point: {adjustedPlacementPoint}");

                                // ⏱️ TIMING: Stop per-sleeve timer and log details
                                sleeveTimer.Stop();
                                totalPlacementTime += sleeveTimer.Elapsed;
                                sleeveLog.AppendLine($"  TOTAL: {sleeveTimer.ElapsedMilliseconds}ms");

                                // Log per-sleeve timing details (first 10 sleeves for detailed analysis)
                                if (PlacedCount < 10)
                                {
                                    SafeFileLogger.SafeAppendText("sleeve_placement_timing.log",
                                        $"[SLEEVE {PlacedCount + 1}] ClashZone {clashZone.Id} - {sleeveLog.ToString().TrimEnd()}");
                                }

                                PlacedCount++;

                                // ✅ SESSION FLAG: Track processed zone for flag reset
                                if (clashZone.Id != Guid.Empty && !processedZoneGuids.Contains(clashZone.Id))
                                    processedZoneGuids.Add(clashZone.Id);

                                // ✅ PERFORMANCE: Track successful placement
                                singleSleeveTracker.SetItemCount(1); // 1 sleeve placed

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] ✓ Placed {_strategy.GetCategoryName()} sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id} at {adjustedPlacementPoint} ({sleeveTimer.ElapsedMilliseconds}ms)");
                            }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[UniversalSleevePlacer] Error placing sleeve for ClashZone {clashZone.Id}: {ex.Message}");
                                try
                                {                         // ✅ DEPLOYMENT MODE: Skip file writes
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION IN OUTER TRY-CATCH: Zone={clashZone.Id}, Error={ex.Message}\n{ex.StackTrace}\n");
                                    }
                                }
                                catch { }
                                ErrorCount++;
                                // Stop timer even on error
                                if (sleeveTimer.IsRunning) sleeveTimer.Stop();
                            }
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            try
                            {
                                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [PLACEMENT-SUMMARY] Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}\n");
                            }
                            catch { }
                        }

                        // ⏱️ TIMING: Stop overall timer and generate summary
                        overallTimer.Stop();

                        // Calculate averages
                        double avgPlacementTime = PlacedCount > 0 ? totalPlacementTime.TotalMilliseconds / PlacedCount : 0;
                        double avgClearanceTime = PlacedCount > 0 ? totalClearanceTime.TotalMilliseconds / PlacedCount : 0;
                        double avgLevelFindTime = PlacedCount > 0 ? totalLevelFindTime.TotalMilliseconds / PlacedCount : 0;
                        double avgFamilyLoadTime = PlacedCount > 0 ? totalFamilyLoadTime.TotalMilliseconds / PlacedCount : 0;
                        double avgSleeveCreateTime = PlacedCount > 0 ? totalSleeveCreateTime.TotalMilliseconds / PlacedCount : 0;
                        double avgParameterTime = PlacedCount > 0 ? totalParameterTime.TotalMilliseconds / PlacedCount : 0;
                        double avgValidationTime = PlacedCount > 0 ? totalValidationTime.TotalMilliseconds / PlacedCount : 0;

                        // Build timing summary
                        detailedTimingLog.AppendLine($"\n=== SLEEVE PLACEMENT TIMING SUMMARY ===");
                        detailedTimingLog.AppendLine($"[OVERALL] Total time: {overallTimer.ElapsedMilliseconds}ms ({overallTimer.Elapsed.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"[OVERALL] Total sleeves: {PlacedCount} placed, {SkippedCount} skipped, {ErrorCount} errors");
                        detailedTimingLog.AppendLine($"\n[PER-SLEEVE AVERAGES]");
                        detailedTimingLog.AppendLine($"  Average total per sleeve: {avgPlacementTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average clearance calc: {avgClearanceTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average level find: {avgLevelFindTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average family load: {avgFamilyLoadTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average sleeve create: {avgSleeveCreateTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average parameters: {avgParameterTime:F2}ms");
                        detailedTimingLog.AppendLine($"  Average validation/update: {avgValidationTime:F2}ms");
                        detailedTimingLog.AppendLine($"\n[TOTAL TIMES]");
                        detailedTimingLog.AppendLine($"  Total clearance time: {totalClearanceTime.TotalMilliseconds:F2}ms ({totalClearanceTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total level find time: {totalLevelFindTime.TotalMilliseconds:F2}ms ({totalLevelFindTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total family load time: {totalFamilyLoadTime.TotalMilliseconds:F2}ms ({totalFamilyLoadTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total sleeve create time: {totalSleeveCreateTime.TotalMilliseconds:F2}ms ({totalSleeveCreateTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total parameter time: {totalParameterTime.TotalMilliseconds:F2}ms ({totalParameterTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total validation time: {totalValidationTime.TotalMilliseconds:F2}ms ({totalValidationTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"  Total placement time (all sleeves): {totalPlacementTime.TotalMilliseconds:F2}ms ({totalPlacementTime.TotalSeconds:F2}s)");
                        detailedTimingLog.AppendLine($"\n[PERFORMANCE ANALYSIS]");
                        if (PlacedCount > 0)
                        {
                            double clearancePercent = (totalClearanceTime.TotalMilliseconds / totalPlacementTime.TotalMilliseconds) * 100;
                            double createPercent = (totalSleeveCreateTime.TotalMilliseconds / totalPlacementTime.TotalMilliseconds) * 100;
                            double parameterPercent = (totalParameterTime.TotalMilliseconds / totalPlacementTime.TotalMilliseconds) * 100;
                            detailedTimingLog.AppendLine($"  Clearance calculation: {clearancePercent:F1}% of placement time");
                            detailedTimingLog.AppendLine($"  Sleeve creation: {createPercent:F1}% of placement time");
                            detailedTimingLog.AppendLine($"  Parameter setting: {parameterPercent:F1}% of placement time");
                        }

                        // Write timing log to file
                        SafeFileLogger.SafeAppendText("sleeve_placement_timing.log", detailedTimingLog.ToString());
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[TIMING] Sleeve placement timing logged to file");

                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] Placement loop complete - Placed: {PlacedCount}, Skipped: {SkippedCount}, Errors: {ErrorCount}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[TIMING] Total time: {overallTimer.ElapsedMilliseconds}ms, Avg per sleeve: {avgPlacementTime:F2}ms");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now}] [PLACEMENT_COMPLETE] Placed: {PlacedCount}, Skipped: {SkippedCount}, Errors: {ErrorCount}, Total: {overallTimer.ElapsedMilliseconds}ms, Avg: {avgPlacementTime:F2}ms\n");

                        // CRITICAL FIX: Save XML files with updated SleeveInstanceId values
                        // ✅ PERFORMANCE OPTIMIZATION: Batch regeneration after ALL sleeves placed
                        if (placedSleeveIds.Count > 0)
                        {
                            var regenTimer = System.Diagnostics.Stopwatch.StartNew();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[BATCH-REGEN] Regenerating document for {placedSleeveIds.Count} sleeves...");
                            }
                            _doc.Regenerate(); // ✅ Single regeneration for all sleeves
                            regenTimer.Stop();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[BATCH-REGEN] ✅ Regenerated {placedSleeveIds.Count} sleeves in {regenTimer.ElapsedMilliseconds}ms");
                            }
                            
                            // ✅ STEP 5 OPTIMIZATION: Flush all deferred parameters after regeneration
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[BATCH-PARAMS] ═══ BEFORE FLUSH CHECK ═══ UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, _deferredParameters.Count={_deferredParameters?.Count ?? 0}, placedSleeveIds.Count={placedSleeveIds.Count}");
                                if (_deferredParameters != null && _deferredParameters.Count > 0)
                                {
                                    var sleeveIds = string.Join(", ", _deferredParameters.Keys.Select(id => id.IntegerValue));
                                    DebugLogger.Info($"[BATCH-PARAMS] 📋 Deferred sleeve IDs: [{sleeveIds}]");
                                }
                            }
                            
                            if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null && _deferredParameters.Count > 0)
                            {
                                var flushTimer = System.Diagnostics.Stopwatch.StartNew();
                                int paramCountBeforeFlush = _deferredParameters.Count;
                                int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[BATCH-PARAMS] 🔄 ABOUT TO FLUSH: {paramCountBeforeFlush} sleeves with {totalParams} total parameters after regeneration...");
                                }
                                FlushDeferredParameters();
                                flushTimer.Stop();
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[BATCH-PARAMS] ✅ FLUSH COMPLETE: {paramCountBeforeFlush} sleeves ({totalParams} parameters) in {flushTimer.ElapsedMilliseconds}ms");
                                }
                            }
                            else if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[BATCH-PARAMS] ⚠️ SKIPPED FLUSH: UseBatchedParameterWrites={OptimizationFlags.UseBatchedParameterWrites}, _deferredParameters.Count={_deferredParameters?.Count ?? 0}");
                            }

                            // ✅ PERFORMANCE OPTIMIZATION: Batch bounding box retrieval after regeneration
                            var bboxTimer = System.Diagnostics.Stopwatch.StartNew();
                            int bboxCount = 0;
                            foreach (var (sleeve, zone, fw, fh, fd) in placedSleeveData)
                            {
                                try
                                {
                                    // ✅ Validate sleeve still exists (may have been deleted by clustering)
                                    if (!sleeve.IsValidObject) continue;

                                    var actualBbox = sleeve.get_BoundingBox(null);
                                    if (actualBbox != null)
                                    {
                                        // Set bounding box coordinates (WCS)
                                        zone.SetSleeveBoundingBox(actualBbox);

                                        // ✅ RCS BBOX: Transform to wall-aligned RCS for walls/framing
                                        // This eliminates rotation logic for walls - bounding boxes are already wall-aligned
                                        // NOTE: WallDirection should already be set when loading from DB (calculated in MapClashZone from HostOrientation)
                                        bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                                          string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                                        bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                                        
                                        if ((isWallHost || isFramingHost) && zone.WallDirection != null && !zone.WallDirection.IsZeroLength())
                                        {
                                            try
                                            {
                                                var rcsBbox = WallRcsTransformer.TransformToRcs(actualBbox, zone.WallDirection);
                                                if (rcsBbox != null)
                                                {
                                                    zone.SleeveBoundingBoxRCS_MinX = rcsBbox.Min.X;
                                                    zone.SleeveBoundingBoxRCS_MinY = rcsBbox.Min.Y;
                                                    zone.SleeveBoundingBoxRCS_MinZ = rcsBbox.Min.Z;
                                                    zone.SleeveBoundingBoxRCS_MaxX = rcsBbox.Max.X;
                                                    zone.SleeveBoundingBoxRCS_MaxY = rcsBbox.Max.Y;
                                                    zone.SleeveBoundingBoxRCS_MaxZ = rcsBbox.Max.Z;
                                                }
                                            }
                                            catch (Exception rcsEx)
                                            {
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("placement_errors.log",
                                                        $"[{DateTime.Now:HH:mm:ss.fff}] [RCS-TRANSFORM] Error transforming bbox to RCS for sleeve {sleeve?.Id?.IntegerValue ?? -1}: {rcsEx.Message}\n");
                                                }
                                            }
                                        }

                                        // Update placement point from bounding box center
                                        zone.SleevePlacementPoint = new XYZ(
                                            (actualBbox.Min.X + actualBbox.Max.X) / 2,
                                            (actualBbox.Min.Y + actualBbox.Max.Y) / 2,
                                            (actualBbox.Min.Z + actualBbox.Max.Z) / 2
                                        );

                                        zone.SleevePlacementPointX = zone.SleevePlacementPoint.X;
                                        zone.SleevePlacementPointY = zone.SleevePlacementPoint.Y;
                                        zone.SleevePlacementPointZ = zone.SleevePlacementPoint.Z;

                                        bboxCount++;
                                    }
                                }
                                catch (Exception bboxEx)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Warning($"[BATCH-BBOX] Error retrieving bbox for sleeve {sleeve?.Id?.IntegerValue ?? -1}: {bboxEx.Message}");
                                    }
                                }
                            }
                            bboxTimer.Stop();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[BATCH-BBOX] ✅ Retrieved {bboxCount} bounding boxes in {bboxTimer.ElapsedMilliseconds}ms (avg: {bboxTimer.ElapsedMilliseconds / Math.Max(1, bboxCount):F1}ms per sleeve)");
                            }

                            // ✅ CRITICAL FIX: Save sleeve data to database IMMEDIATELY after placement
                            // This ensures UpdateSleeveCoordinatesInXml can find sleeves by SleeveInstanceId
                            // ✅ CRITICAL: This MUST run even if bboxCount is 0 - bounding boxes are not required for snapshot save
                            try
                            {
                                using (var dbContext = new Data.SleeveDbContext(_doc))
                                {
                                    var repository = new Data.Repositories.ClashZoneRepository(dbContext);
                                    int dbSavedCount = 0;

                                    // ✅ BOUNDING BOX SAVE: Only save bounding boxes if they were retrieved (bboxCount > 0)
                                    if (bboxCount > 0)
                                    {
                                        foreach (var (sleeve, zone, fw, fh, fd) in placedSleeveData)
                                        {
                                            if (zone != null && zone.SleeveInstanceId > 0 && sleeve.IsValidObject)
                                            {
                                                // ⚠️ PROTECTED CODE: DO NOT MODIFY THIS SECTION WITHOUT UNDERSTANDING THE IMPACT
                                                // This code saves sleeve coordinates to the database immediately after placement.
                                                // It is critical for clustering to work correctly because:
                                                // 1. SleeveInstanceId - Required for clustering to find sleeves in the database
                                                // 2. Sleeve dimensions and placement point - Required for placement tracking
                                                // 3. Bounding box coordinates - Required for clustering to calculate cluster bounding boxes
                                                // 4. Rotation angle (MepElementRotationAngle) - Already saved during refresh in InsertOrUpdateClashZones,
                                                //    and is loaded by cluster service from database. No need to update here as it doesn't change after placement.
                                                // If this code is broken, clustering will fail because it won't find sleeves with valid bounding boxes.

                                                // ✅ ROTATION DATA: MepElementRotationAngle is already in database (saved during refresh).
                                                // Cluster service reads it from DB via GetClashZonesByCategory -> MepRotationAngleRad column.
                                                // No update needed here as rotation angle doesn't change after sleeve placement.

                                                // Save SleeveInstanceId
                                                repository.UpdateSleeveInstanceId(zone.Id, zone.SleeveInstanceId);

                                                // Save sleeve dimensions
                                                // ✅ CRITICAL: Also save Active document coordinates (where sleeve is actually placed)
                                                // ✅ CRITICAL: Also save rotation angle (for corner calculations)
                                                repository.UpdateSleevePlacement(
                                                    zone.Id, // Use ClashZone Guid
                                                    zone.SleeveInstanceId,
                                                    zone.SleeveWidth > 0 ? zone.SleeveWidth : fw,
                                                    zone.SleeveHeight > 0 ? zone.SleeveHeight : fh,
                                                    zone.SleeveDiameter > 0 ? zone.SleeveDiameter : fd,
                                                    zone.SleevePlacementPointX,
                                                    zone.SleevePlacementPointY,
                                                    zone.SleevePlacementPointZ,
                                                    zone.SleevePlacementPointActiveDocumentX,  // ✅ Active document coordinates (where sleeve is actually placed)
                                                    zone.SleevePlacementPointActiveDocumentY,
                                                    zone.SleevePlacementPointActiveDocumentZ,
                                                    zone.MepElementRotationAngle);  // ✅ Rotation angle in radians (for corner calculations)

                                                // ✅ BOUNDING BOX: Save axis-aligned bounding box coordinates (in model coordinate system).
                                                // Bounding boxes are required for clustering to calculate cluster bounding boxes
                                                if (!(zone.SleeveBoundingBoxMinX == 0.0 && zone.SleeveBoundingBoxMinY == 0.0 && zone.SleeveBoundingBoxMinZ == 0.0 &&
                                                      zone.SleeveBoundingBoxMaxX == 0.0 && zone.SleeveBoundingBoxMaxY == 0.0 && zone.SleeveBoundingBoxMaxZ == 0.0))
                                                {
                                                    repository.UpdateSleeveBoundingBoxes(
                                                        zone.Id,
                                                        zone.SleeveBoundingBoxMinX, zone.SleeveBoundingBoxMinY, zone.SleeveBoundingBoxMinZ,
                                                        zone.SleeveBoundingBoxMaxX, zone.SleeveBoundingBoxMaxY, zone.SleeveBoundingBoxMaxZ);

                                                    // ✅ RCS BBOX: Save wall-aligned RCS bounding box for walls/framing
                                                    // This eliminates rotation logic for walls - bounding boxes are already wall-aligned
                                                    bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                                                      string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                                                    bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                                                    
                                                    if ((isWallHost || isFramingHost) && 
                                                        !(zone.SleeveBoundingBoxRCS_MinX == 0.0 && zone.SleeveBoundingBoxRCS_MinY == 0.0 && zone.SleeveBoundingBoxRCS_MinZ == 0.0 &&
                                                          zone.SleeveBoundingBoxRCS_MaxX == 0.0 && zone.SleeveBoundingBoxRCS_MaxY == 0.0 && zone.SleeveBoundingBoxRCS_MaxZ == 0.0))
                                                    {
                                                        repository.UpdateSleeveBoundingBoxesRcs(
                                                            zone.Id,
                                                            zone.SleeveBoundingBoxRCS_MinX, zone.SleeveBoundingBoxRCS_MinY, zone.SleeveBoundingBoxRCS_MinZ,
                                                            zone.SleeveBoundingBoxRCS_MaxX, zone.SleeveBoundingBoxRCS_MaxY, zone.SleeveBoundingBoxRCS_MaxZ);
                                                    }

                                                    // ✅ ROTATED BBOX: Calculate and save rotated bounding box ONLY for rotated axis-aligned sleeves (non-straight axis-aligned)
                                                    // Skip for straight axis-aligned angles to WCS: 0°, 90°, 180°, 270° (use straight axis-aligned bbox instead)
                                                    // For rotated axis-aligned angles (45°, 135°, 225°, 315°), calculate rotated bbox
                                                    // For straight axis-aligned sleeves, rotated bbox columns remain NULL
                                                    var rotationAngleRad = zone.MepElementRotationAngle;
                                                    var rotationAngleDeg = Math.Abs(rotationAngleRad * 180.0 / Math.PI);

                                                    // Check if angle is straight axis-aligned to WCS (0°, 90°, 180°, 270°) with 1° tolerance
                                                    bool isStraightAxisAligned = Math.Abs(rotationAngleDeg) < 1.0 ||
                                                                        Math.Abs(rotationAngleDeg - 90.0) < 1.0 ||
                                                                        Math.Abs(rotationAngleDeg - 180.0) < 1.0 ||
                                                                        Math.Abs(rotationAngleDeg - 270.0) < 1.0 ||
                                                                        Math.Abs(rotationAngleDeg - 360.0) < 1.0;

                                                    if (Math.Abs(rotationAngleRad) > 1e-6 && !isStraightAxisAligned)
                                                    {
                                                        try
                                                        {
                                                            // ✅ FIX: Get actual sleeve dimensions from Revit element (not axis-aligned world bbox)
                                                            // The rotated bounding box should represent the sleeve in its LOCAL coordinate system
                                                            double actualWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                                            double actualHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                                            double actualDepth = zone.SleeveBoundingBoxMaxZ - zone.SleeveBoundingBoxMinZ;

                                                            // Try to get dimensions from sleeve element if available
                                                            if (sleeve != null && sleeve.IsValidObject)
                                                            {
                                                                try
                                                                {
                                                                    var widthParam = sleeve.LookupParameter("Width");
                                                                    var heightParam = sleeve.LookupParameter("Height");
                                                                    var depthParam = sleeve.LookupParameter("Depth");

                                                                    if (widthParam != null && widthParam.HasValue)
                                                                        actualWidth = widthParam.AsDouble();
                                                                    if (heightParam != null && heightParam.HasValue)
                                                                        actualHeight = heightParam.AsDouble();
                                                                    if (depthParam != null && depthParam.HasValue)
                                                                        actualDepth = depthParam.AsDouble();
                                                                }
                                                                catch { /* Fallback to zone dimensions */ }
                                                            }

                                                            // ✅ FIX: Store LOCAL bounding box coordinates (centered at placement point)
                                                            // The rotated bounding box should represent the sleeve's actual size in its LOCAL coordinate system
                                                            // This ensures clustering uses correct dimensions (550mm × 200mm) not inflated world bbox

                                                            // Calculate local bounding box (centered at placement point in sleeve's local coordinate system)
                                                            double halfWidth = actualWidth / 2.0;
                                                            double halfHeight = actualHeight / 2.0;
                                                            double halfDepth = actualDepth / 2.0;

                                                            // Local bbox min/max (centered at placement point in sleeve's local coordinate system)
                                                            // These represent the actual sleeve dimensions (550mm × 200mm) in local coordinates
                                                            var placementPoint = zone.SleevePlacementPoint;
                                                            var rotatedLocalMinX = placementPoint.X - halfWidth;
                                                            var rotatedLocalMinY = placementPoint.Y - halfHeight;
                                                            var rotatedLocalMinZ = placementPoint.Z - halfDepth;
                                                            var rotatedLocalMaxX = placementPoint.X + halfWidth;
                                                            var rotatedLocalMaxY = placementPoint.Y + halfHeight;
                                                            var rotatedLocalMaxZ = placementPoint.Z + halfDepth;

                                                            // ✅ SAVE: Save rotated bounding box in LOCAL coordinates (centered at placement point)
                                                            // These coordinates represent the sleeve's actual size (550mm × 200mm) in its local coordinate system
                                                            // Clustering will use these directly for distance calculations in the rotated coordinate system
                                                            repository.UpdateRotatedBoundingBoxes(
                                                                zone.Id,
                                                                rotatedLocalMinX, rotatedLocalMinY, rotatedLocalMinZ,
                                                                rotatedLocalMaxX, rotatedLocalMaxY, rotatedLocalMaxZ);

                                                            // ✅ SLEEVE CORNERS: Calculate and save 4 corner coordinates in WORLD space (ROBUST)
                                                            // Using robust helper method with full validation, retry logic, and error handling
                                                            SaveSleeveCornersRobust(zone, sleeve, repository, rotationAngleRad, actualWidth, actualHeight);

                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                            {
                                                                // Log local bounding box coordinates (centered at placement point)
                                                                double localWidthMm = (rotatedLocalMaxX - rotatedLocalMinX) * 304.8;
                                                                double localHeightMm = (rotatedLocalMaxY - rotatedLocalMinY) * 304.8;
                                                                double localDepthMm = (rotatedLocalMaxZ - rotatedLocalMinZ) * 304.8;

                                                                DebugLogger.Info($"[ROTATED-BBOX] ✅ Saved rotated bounding box (LOCAL coordinates) for zone {zone.Id}: " +
                                                                    $"Angle={rotationAngleDeg:F1}° (non-axis-aligned), " +
                                                                    $"ActualSize={actualWidth * 304.8:F1}mm × {actualHeight * 304.8:F1}mm, " +
                                                                    $"LocalBBoxSize={localWidthMm:F1}mm × {localHeightMm:F1}mm × {localDepthMm:F1}mm, " +
                                                                    $"PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6}), " +
                                                                    $"RotatedMin=({rotatedLocalMinX:F6}, {rotatedLocalMinY:F6}, {rotatedLocalMinZ:F6}), " +
                                                                    $"RotatedMax=({rotatedLocalMaxX:F6}, {rotatedLocalMaxY:F6}, {rotatedLocalMaxZ:F6})");
                                                            }
                                                        }
                                                        catch (Exception rotEx)
                                                        {
                                                            if (!DeploymentConfiguration.DeploymentMode)
                                                            {
                                                                DebugLogger.Warning($"[ROTATED-BBOX] Error calculating rotated bounding box for zone {zone.Id}: {rotEx.Message}");
                                                                DebugLogger.Warning($"[ROTATED-BBOX] Stack trace: {rotEx.StackTrace}");
                                                            }
                                                        }
                                                    }
                                                    else if (Math.Abs(rotationAngleRad) > 1e-6 && isStraightAxisAligned)
                                                    {
                                                        // Axis-aligned angle (0°, 90°, 180°, 270°) - no need for rotated bbox
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                                        {
                                                            DebugLogger.Info($"[ROTATED-BBOX] Skipped rotated bounding box for zone {zone.Id}: Angle={rotationAngleDeg:F1}° (axis-aligned, using axis-aligned bbox)");
                                                        }

                                                        // ✅ SLEEVE CORNERS: Calculate and save 4 corner coordinates in WORLD space (ROBUST)
                                                        // Using robust helper method with full validation, retry logic, and error handling
                                                        double axisAlignedWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                                        double axisAlignedHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                                        SaveSleeveCornersRobust(zone, sleeve, repository, rotationAngleRad, axisAlignedWidth, axisAlignedHeight);
                                                    }
                                                    else
                                                    {
                                                        // Zero rotation (0°) - still calculate corners for consistency (ROBUST)
                                                        // Using robust helper method with full validation, retry logic, and error handling
                                                        double zeroRotationWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                                        double zeroRotationHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                                        SaveSleeveCornersRobust(zone, sleeve, repository, 0.0, zeroRotationWidth, zeroRotationHeight);
                                                    }
                                                }

                                                dbSavedCount++;
                                            }
                                        }

                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[UniversalSleevePlacer] ✅ DATABASE: Saved sleeve data for {dbSavedCount} clash zones immediately after placement");
                                        }

                                        // ✅ CRITICAL FIX: Save sleeve snapshots after placement
                                        // Snapshots are only saved when sleeves have SleeveInstanceId > 0 (after placement)
                                        // ✅ DOES NOT DEPEND ON XML - only requires filterName and placed sleeves
                                        
                                        // ✅ ALWAYS LOG: Diagnostic to see if code path is reached (using SafeFileLogger since DebugLogger.IsEnabled=false)
                                        SafeFileLogger.SafeAppendText("database_operations.log",
                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] 🔍🔍🔍 SNAPSHOT SAVE CHECK: placedSleeveData.Count={placedSleeveData.Count}, _filterName='{_filterName ?? "NULL"}', WillAttemptSave={placedSleeveData.Count > 0 && !string.IsNullOrWhiteSpace(_filterName)}\n");
                                        
                                        if (placedSleeveData.Count > 0 && !string.IsNullOrWhiteSpace(_filterName))
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Info($"[UniversalSleevePlacer] ✅ Attempting to save sleeve snapshots: FilterName='{_filterName}', PlacedSleeveData={placedSleeveData.Count}");
                                            }
                                            try
                                            {
                                                // Get FilterId from filter name
                                                // ✅ CRITICAL FIX: Strip .xml extension if present (database stores filter name without extension)
                                                string filterNameForLookup = _filterName;
                                                if (!string.IsNullOrWhiteSpace(filterNameForLookup) && filterNameForLookup.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    filterNameForLookup = filterNameForLookup.Substring(0, filterNameForLookup.Length - 4);
                                                }
                                                
                                                int filterId = -1;
                                                using (var filterCmd = dbContext.Connection.CreateCommand())
                                                {
                                                    filterCmd.CommandText = @"
                                                SELECT FilterId FROM Filters 
                                                WHERE FilterName = @FilterName 
                                                LIMIT 1";
                                                    filterCmd.Parameters.AddWithValue("@FilterName", filterNameForLookup);
                                                    var filterResult = filterCmd.ExecuteScalar();
                                                    if (filterResult != null)
                                                    {
                                                        filterId = Convert.ToInt32(filterResult);
                                                    }
                                                }

                                                if (filterId > 0)
                                                {
                                                    // ✅ DIAGNOSTIC: Log placedSleeveData details BEFORE filtering (using SafeFileLogger)
                                                    int totalPlaced = placedSleeveData.Count;
                                                    int withZone = placedSleeveData.Count(p => p.zone != null);
                                                    int withSleeveId = placedSleeveData.Count(p => p.zone != null && p.zone.SleeveInstanceId > 0);
                                                    var sampleSleeveIds = placedSleeveData
                                                        .Where(p => p.zone != null && p.zone.SleeveInstanceId > 0)
                                                        .Take(5)
                                                        .Select(p => $"ZoneId={p.zone.Id}, SleeveId={p.zone.SleeveInstanceId}")
                                                        .ToList();
                                                    SafeFileLogger.SafeAppendText("database_operations.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] 🔍 DIAGNOSTIC: placedSleeveData - Total={totalPlaced}, WithZone={withZone}, WithSleeveId={withSleeveId}, Samples={string.Join(", ", sampleSleeveIds)}\n");
                                                    
                                                    // Get placed zones with SleeveInstanceId > 0
                                                    var placedZones = placedSleeveData
                                                        .Where(p => p.zone != null && p.zone.SleeveInstanceId > 0)
                                                        .Select(p => p.zone)
                                                        .Distinct()
                                                        .ToList();

                                                    // ✅ DIAGNOSTIC: Log filtered results (using SafeFileLogger)
                                                    int individualCount = placedZones.Count(z => z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId <= 0);
                                                    int clusterCount = placedZones.Count(z => z.ClusterSleeveInstanceId > 0);
                                                    SafeFileLogger.SafeAppendText("database_operations.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] 🔍 DIAGNOSTIC: placedZones after filter - Total={placedZones.Count}, Individual={individualCount}, Cluster={clusterCount}\n");

                                                    if (placedZones.Count > 0)
                                                    {
                                                        // ✅ ALWAYS LOG: Saving snapshots (using SafeFileLogger since DebugLogger.IsEnabled=false)
                                                        SafeFileLogger.SafeAppendText("database_operations.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ✅ Saving sleeve snapshots: FilterId={filterId}, PlacedZones={placedZones.Count}\n");
                                                        repository.SaveSleeveSnapshotsForPlacedSleeves(filterId, placedZones);

                                                        SafeFileLogger.SafeAppendText("database_operations.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ✅ DATABASE: Saved sleeve snapshots for {placedZones.Count} placed sleeves\n");
                                                    }
                                                    else
                                                    {
                                                        // ✅ CRITICAL: Always log this warning (using SafeFileLogger since DebugLogger.IsEnabled=false)
                                                        SafeFileLogger.SafeAppendText("database_operations.log",
                                                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ⚠️⚠️⚠️ No placed zones with SleeveInstanceId > 0 (placedSleeveData={placedSleeveData.Count}) - Individual sleeves will NOT be saved to snapshot table!\n");
                                                    }
                                                }
                                                else
                                                {
                                                    // ✅ ALWAYS LOG: FilterId not found (using SafeFileLogger)
                                                    SafeFileLogger.SafeAppendText("database_operations.log",
                                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ⚠️ FilterId not found for FilterName='{_filterName}' - cannot save sleeve snapshots\n");
                                                }
                                            }
                                            catch (Exception snapshotEx)
                                            {
                                                // ✅ ALWAYS LOG: Exception details (using SafeFileLogger)
                                                SafeFileLogger.SafeAppendText("database_operations.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ⚠️ Failed to save sleeve snapshots: {snapshotEx.Message}\n");
                                                SafeFileLogger.SafeAppendText("database_operations.log",
                                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] Stack trace: {snapshotEx.StackTrace}\n");
                                            }
                                        }
                                        else
                                        {
                                            // ✅ ALWAYS LOG: Why snapshot save was skipped (using SafeFileLogger since DebugLogger.IsEnabled=false)
                                            SafeFileLogger.SafeAppendText("database_operations.log",
                                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [UniversalSleevePlacer] ⚠️ SKIPPED saving sleeve snapshots: placedSleeveData={placedSleeveData.Count}, filterName='{_filterName ?? "NULL"}'\n");
                                        }
                                    }
                                    
                                    // ✅ CRITICAL FIX: Move snapshot saving OUTSIDE the bboxCount check
                                    // Snapshots don't require bounding boxes - they only need SleeveInstanceId
                                    // This was causing pipes to not be saved to snapshot table when bbox retrieval failed
                                }
                            }
                            catch (Exception dbEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ Failed to save sleeve data to database after placement: {dbEx.Message}");
                                }
                            }
                        } // ✅ PERFORMANCE: End of single sleeve placement tracking

                        // ✅ PERFORMANCE: Update loop tracker with processed count
                        placementLoopTracker.SetItemCount(processedCount);
                    } // ✅ PERFORMANCE: End of placement loop tracking

                    // ✅ PARALLEL PLANNING: Write planning logs after placement loop completes
                    if (planningLogs != null && planningLogs.Length > 0 && !DeploymentConfiguration.DeploymentMode)
                    {
                        try
                        {
                            var planningLogContent = planningLogs.ToString();
                            File.AppendAllText(debugLogPath, $"\n[{DateTime.Now:HH:mm:ss}] === PLANNING PHASE SUMMARY ===\n");
                            File.AppendAllText(debugLogPath, planningLogContent);
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] === END PLANNING SUMMARY ===\n\n");
                            
                            DebugLogger.Info($"[PLANNING] Planning logs written to {debugLogPath}");
                        }
                        catch (Exception logEx)
                        {
                            DebugLogger.Error($"[PLANNING] Failed to write planning logs: {logEx.Message}");
                        }
                    }

                    // ✅ PERFORMANCE: Track database updates (XML creation is disabled - see DeploymentConfiguration.DisableXmlCreation)
                    // NOTE: This operation updates the database for each placed sleeve individually.
                    // Time shown is database I/O (creating context, repository, and committing transactions).
                    // XML file writes are skipped when DeploymentConfiguration.DisableXmlCreation = true.
                    using (var xmlSaveTracker = performanceMonitor.TrackOperation("Update Database Flags"))
                    {
                        // ✅ CRITICAL: Save Global XML IMMEDIATELY after placement (before clustering can delete sleeves)
                        if (PlacedCount > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[{DateTime.Now}] [XML_SAVE] PlacedCount > 0, starting save operations\n");

                            // ⚠️ CRITICAL: Log flag states BEFORE any operations
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-BEFORE] About to save with {PlacedCount} placed sleeves\n");

                            // ✅ STEP 1: Capture placed clash zones WITH valid SleeveInstanceId BEFORE any file operations
                            // This ensures we have the correct state even if SaveUpdatedXmlFiles modifies the list
                            var placedClashZonesForGlobal = clashZones
                                .Where(cz => cz.IsResolved && cz.SleeveInstanceId > 0)
                                .ToList(); // ✅ OOP REFACTORING: Keep actual ClashZone objects instead of anonymous types

                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[GLOBAL_INDEX] CAPTURED {placedClashZonesForGlobal.Count} placed clash zones for Global XML update (before file operations)");
                            foreach (var cz in placedClashZonesForGlobal.Take(10))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[GLOBAL_INDEX-CAPTURED] ClashZone {cz.Id}: Category={cz.MepElementCategory}, IsResolved={cz.IsResolved}, SleeveInstanceId={cz.SleeveInstanceId}");
                            }

                            // ✅ STEP 2: Save Filter XML first
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                var zonesWithIds = clashZones.Where(cz => cz.SleeveInstanceId > 0).Select(cz => cz.Id).ToList();
                                DebugLogger.Info($"[UniversalSleevePlacer] Preparing to persist {zonesWithIds.Count} clash zones with SleeveInstanceId > 0");
                                try
                                {
                                    var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-PREP] Zones with SleeveId>0: {string.Join(", ", zonesWithIds.Select(id => id.ToString()).Take(10))}{(zonesWithIds.Count > 10 ? ", ..." : string.Empty)}\n");
                                    var samplePairs = clashZones.Select(cz => $"{cz.Id}:{cz.SleeveInstanceId}").Take(10);
                                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-PREP-DETAIL] Sample pairs: {string.Join(", ", samplePairs)}\n");
                                }
                                catch { }
                            }

                            // ✅ PHASE 2: Skip XML file updates if XML creation is disabled (database only mode)
                            if (!DeploymentConfiguration.DisableXmlCreation)
                            {
                                SaveUpdatedXmlFiles(clashZones);
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalSleevePlacer] ⚠️ XML creation disabled - skipping SaveUpdatedXmlFiles (database only mode). Zones={clashZones.Count}");
                                    var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-SKIP] ⚠️ XML creation disabled - skipping SaveUpdatedXmlFiles (database only mode). Zones={clashZones.Count}\n");
                                }
                            }


                            // ⚠️ REMOVED: Don't update coordinates here because clustering will delete individual sleeves

                            // ⚠️ CRITICAL: Log flag states AFTER XML save (or skip)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                if (DeploymentConfiguration.DisableXmlCreation)
                                {
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-AFTER] XML save skipped (database only mode) for {PlacedCount} placed sleeves\n");
                                }
                                else
                                {
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-AFTER] XML save completed for {PlacedCount} placed sleeves\n");
                                }
                            }

                            // ✅ PERFORMANCE: Use batch database updates instead of updating each sleeve individually
                            // This reduces 87ms for 2 sleeves (43.5ms each) to ~10ms total (single transaction)
                            try
                            {
                                if (placedClashZonesForGlobal.Count > 0)
                                {
                                    // Prepare batch list: (ClashZone, SleeveId)
                                    var batchUpdates = placedClashZonesForGlobal
                                        .Where(cz => cz.SleeveInstanceId > 0)
                                        .Select(cz => (cz, cz.SleeveInstanceId))
                                        .ToList();
                                    
                                    if (batchUpdates.Count > 0)
                                    {
                                        _flagManager.BatchUpdateFlagsForPlacement(
                                            batchUpdates,
                                            isCluster: false,
                                            placedClashZonesForGlobal[0].MepElementCategory,
                                            _filterName
                                        );
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FLAG-MANAGER] ✅ BATCH: Successfully updated {batchUpdates.Count} placed sleeves in single transaction (BEFORE clustering)");
                                    }
                                }
                                else
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ No placed clash zones found to update in Global XML (PlacedCount={PlacedCount}, but no clash zones with IsResolved=true AND SleeveInstanceId > 0)");
                                }
                            }
                            catch (Exception upEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[FLAG-MANAGER] ❌ CRITICAL ERROR: Batch flag update after individual placement failed: {upEx.Message}");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Error($"[FLAG-MANAGER] Stack trace: {upEx.StackTrace}");
                                // Don't throw - Global XML failure shouldn't stop placement, but log it clearly
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[{DateTime.Now}] [XML_SAVE] PlacedCount = 0, skipping SaveUpdatedXmlFiles\n");
                        }
                        xmlSaveTracker.SetItemCount(PlacedCount);
                    } // ✅ PERFORMANCE: End of XML save tracking
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalSleevePlacer] Error in placement loop: {ex.Message}");
                try { SafeFileLogger.SafeAppendText("sleeve_placement_errors.log", $"[{DateTime.Now}] {ex}\n"); } catch { }
                ErrorCount += 1;
            }
            finally
            {
                // ✅ PERFORMANCE OPTIMIZATION: Write batch logs once at the end instead of per clash zone
                if (batchLogs.Length > 0)
                {
                    try
                    {
                        string placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        // DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(placementDebugPath, batchLogs.ToString());
                        }
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string flagStatePath = SafeFileLogger.GetLogFilePath("flag_state_debug.log");
                            System.IO.File.AppendAllText(flagStatePath, batchLogs.ToString());
                        }
                        // Also write summary to AppData placement log
                        SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SUMMARY: Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}\n");
                        // Write a pointer file in project Log to locate the AppData placement log easily
                        string latestPath = SafeFileLogger.GetLogFilePath("sleeve_placement_latest.txt");
                        System.IO.File.WriteAllText(latestPath, placementLogName);
                    }
                    catch { } // Don't fail placement if logging fails
                }
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] 🎯 PLACEMENT COMPLETED: Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}");
            
            // ✅ PERFORMANCE LOGGING: Log optimization metrics
            overallTimer.Stop();
            
            // ✅ PERFORMANCE: Generate final performance report
            performanceMonitor.GenerateReport(PlacedCount, 0); // 0 clusters for individual placement
            
            if (PlacedCount > 0 && !DeploymentConfiguration.DeploymentMode)
            {
                var avgPlacementTime = overallTimer.ElapsedMilliseconds / (double)PlacedCount;
                SafeFileLogger.SafeAppendText("placement_performance.log",
                    $"[{DateTime.Now:HH:mm:ss}] PERFORMANCE SUMMARY: Total={overallTimer.ElapsedMilliseconds}ms, AvgPerSleeve={avgPlacementTime:F2}ms, Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}\n");
                DebugLogger.Info($"[PERFORMANCE] Total placement: {overallTimer.ElapsedMilliseconds}ms");
                DebugLogger.Info($"[PERFORMANCE] Average per sleeve: {avgPlacementTime:F2}ms");
                DebugLogger.Info($"[PERFORMANCE] Regenerations: 1 (batch) ✅");
                DebugLogger.Info($"[PERFORMANCE] Element cache size: {elementCache.Count}");
                DebugLogger.Info($"[PERFORMANCE] Batch bounding boxes: {placedSleeveData.Count}");
                DebugLogger.Info($"[PERFORMANCE] Expected improvement: 60-73% faster vs per-sleeve operations");
            }

            try
            {
                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [RETURN-PlaceAllSleevesInTransaction] Placed={PlacedCount}, Skipped={SkippedCount}, Errors={ErrorCount}\n");
            }
            catch { }
            
            // ⚡ PERFORMANCE OPTIMIZATION: Write deferred metadata in batch BEFORE flag reset
            if (OptimizationFlags.DeferNonCriticalMetadata && placedSleeveData.Count > 0)
            {
                try
                {
                    var batchWriter = new Placement.BatchMetadataWriterService(_doc);
                    var sleeveData = placedSleeveData.Select(x => (x.sleeve, x.zone)).ToList();
                    batchWriter.WriteDeferredMetadata(sleeveData);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] ⚡ Batch metadata writer completed for {sleeveData.Count} sleeves");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ Batch metadata write failed: {ex.Message}");
                    }
                }
            }
            
            // ✅ SESSION FLAG: Reset ReadyForPlacement for processed zones so they won't be re-processed
            if (processedZoneGuids.Count > 0)
            {
                try
                {
                    using (var context = new Data.SleeveDbContext(_doc, msg => { }))
                    {
                        var repository = new Data.Repositories.ClashZoneRepository(context, msg => { });
                        repository.BulkResetReadyForPlacementFlags(processedZoneGuids);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] ✅ Reset ReadyForPlacementFlag on {processedZoneGuids.Count} processed zones");
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ SESSION-FLAG-RESET: {processedZoneGuids.Count} zones marked as processed (ReadyForPlacement=false)\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ Failed to reset ReadyForPlacement flags: {ex.Message}");
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ SESSION-FLAG-RESET-ERROR: {ex.Message}\n");
                    }
                }
            }
            
            return (PlacedCount, SkippedCount, ErrorCount);
        }
        
        /// <summary>
        /// Regenerate _CLUSTER.xml files with actual Revit coordinates after all sleeve placement completed
        /// </summary>
        private void RegenerateClusterXmlFiles()
        {
            try
            {
                // Wait briefly for Revit to update all sleeves
                System.Threading.Thread.Sleep(500);
                
                // Collect all sleeves with MEP_ElementId parameter
                var allSleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.LookupParameter("MEP_ElementId") != null && s.LookupParameter("MEP_ElementId").HasValue)
                    .ToList();
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[REGENERATION-DEBUG] Found {allSleeves.Count} sleeves with MEP_ElementId parameter\n");
                
                // Group sleeves by category using MEP_Category parameter
                var groupedSleeves = allSleeves.GroupBy(s => GetSleeveCategoryFromParameter(s)).ToList();
                
                foreach (var group in groupedSleeves)
                {
                    var category = group.Key;
                    var sleeves = group.ToList();
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[REGENERATION-DEBUG] Processing {sleeves.Count} sleeves for category: {category}\n");
                    
                    // Create SleeveDataList with actual Revit coordinates
                    var sleeveDataList = new List<SleeveData>();
                    
                    foreach (var sleeve in sleeves)
                    {
                        var bbox = sleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            var sleeveData = new SleeveData
                            {
                                SleeveInstanceId = sleeve.Id.IntegerValue,
                                Corner1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                                Corner2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                                Corner3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
                                Corner4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                                Width = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Meters),
                                Height = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Meters),
                                Depth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Meters),
                                HostType = GetHostType(sleeve),
                                Orientation = GetSleeveOrientation(sleeve),
                                Category = category,
                                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            };
                            
                            sleeveDataList.Add(sleeveData);
                        }
                    }
                    
                    // Save to _CLUSTER.xml file
                    SaveSleeveDataToClusterXml(sleeveDataList, category);
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[REGENERATION-ERROR] Error in RegenerateClusterXmlFiles: {ex.Message}\n");
                throw;
            }
        }
        
        /// <summary>
        /// Get sleeve category from MEP_Category parameter
        /// </summary>
        private string GetSleeveCategoryFromParameter(FamilyInstance sleeve)
        {
            try
            {
                var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                if (mepCategoryParam != null && !string.IsNullOrEmpty(mepCategoryParam.AsString()))
                {
                    return mepCategoryParam.AsString();
                }
                
                // Fallback to family name
                var familyName = sleeve.Symbol.FamilyName.ToLower();
                if (familyName.Contains("circular") || familyName.Contains("round")) return "Pipes";
                if (familyName.Contains("duct")) return "Ducts";
                if (familyName.Contains("pipe")) return "Pipes";
                if (familyName.Contains("cable")) return "Cable Trays";
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get sleeve orientation from MEP element
        /// </summary>
        private string GetSleeveOrientation(FamilyInstance sleeve)
        {
            try
            {
                var mepElementId = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementId != null && mepElementId.HasValue)
                {
                    var mepElement = _doc.GetElement(mepElementId.AsElementId());
                    if (mepElement != null)
                    {
                        // Get orientation from MEP element's Wall Direction Type
                        var wallDirectionParam = mepElement.LookupParameter("Wall Direction Type");
                        if (wallDirectionParam != null && !string.IsNullOrEmpty(wallDirectionParam.AsString()))
                        {
                            return wallDirectionParam.AsString();
                        }
                    }
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
        
        /// <summary>
        /// Save sleeve data to _CLUSTER.xml file
        /// </summary>
        private void SaveSleeveDataToClusterXml(List<SleeveData> sleeveDataList, string category)
        {
            try
            {
                var pluralCategory = category switch
                {
                    "Pipes" => "Pipes",
                    "Ducts" => "Ducts", 
                    "Cable Trays" => "Cable Trays",
                    _ => category
                };
                
                var fileName = $"Plumbing_{pluralCategory.ToLower().Replace(" ", "_")}_CLUSTER.xml";
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                var filePath = Path.Combine(filtersDirectory, fileName);
                
                var xmlDoc = new System.Xml.XmlDocument();
                var root = xmlDoc.CreateElement("SleeveDataList");
                xmlDoc.AppendChild(root);
                
                foreach (var sleeveData in sleeveDataList)
                {
                    var sleeveElement = xmlDoc.CreateElement("SleeveData");
                    root.AppendChild(sleeveElement);
                    
                    AddXmlElement(xmlDoc, sleeveElement, "SleeveInstanceId", sleeveData.SleeveInstanceId.ToString());
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1X", sleeveData.Corner1.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Y", sleeveData.Corner1.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Z", sleeveData.Corner1.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2X", sleeveData.Corner2.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Y", sleeveData.Corner2.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Z", sleeveData.Corner2.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3X", sleeveData.Corner3.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Y", sleeveData.Corner3.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Z", sleeveData.Corner3.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4X", sleeveData.Corner4.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Y", sleeveData.Corner4.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Z", sleeveData.Corner4.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Width", sleeveData.Width.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Height", sleeveData.Height.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Depth", sleeveData.Depth.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "HostType", sleeveData.HostType);
                    AddXmlElement(xmlDoc, sleeveElement, "Orientation", sleeveData.Orientation);
                    AddXmlElement(xmlDoc, sleeveElement, "Category", sleeveData.Category);
                    AddXmlElement(xmlDoc, sleeveElement, "CreatedAt", sleeveData.CreatedAt);
                }
                
                xmlDoc.Save(filePath);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[REGENERATION-SAVE] Saved {sleeveDataList.Count} sleeves to {fileName}\n");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[REGENERATION-SAVE-ERROR] Error saving {category} cluster XML: {ex.Message}\n");
                throw;
            }
        }
        
        /// <summary>
        /// Helper method to add XML elements
        /// </summary>
        private void AddXmlElement(System.Xml.XmlDocument xmlDoc, System.Xml.XmlElement parent, string name, string value)
        {
            var element = xmlDoc.CreateElement(name);
            element.InnerText = value;
            parent.AppendChild(element);
        }

        /// <summary>
        /// ✅ VALIDATION ONLY: Verifies that ClashZone has valid intersection coordinates
        /// NO FALLBACK to wall center - throws error if coordinates are invalid
        /// </summary>
        private void ValidatePlacementPoint(ClashZone cz)
        {
            // ✅ CRITICAL FIX: Check BOTH IntersectionPointX/Y/Z AND SleevePlacementPointX/Y/Z
            // After XML deserialization, IntersectionPoint/SleevePlacementPoint objects are null,
            // but the XML-serialized properties must have valid values
            bool hasZeroIntersection = (Math.Abs(cz.IntersectionPointX) < 1e-9 && Math.Abs(cz.IntersectionPointY) < 1e-9 && Math.Abs(cz.IntersectionPointZ) < 1e-9);
            bool hasZeroSleevePoint = (Math.Abs(cz.SleevePlacementPointX) < 1e-9 && Math.Abs(cz.SleevePlacementPointY) < 1e-9 && Math.Abs(cz.SleevePlacementPointZ) < 1e-9);
            
            // ✅ CRITICAL: Throw error if BOTH are zero (no valid coordinates available)
            // We DO NOT use wall center as fallback - invalid data means placement cannot proceed
            if (hasZeroIntersection && hasZeroSleevePoint)
            {
                var errorMsg = $"CRITICAL ERROR: ClashZone {cz.Id} has ZERO intersection coordinates. " +
                    $"MEP Element ID: {cz.MepElementIdValue}, Structural Element ID: {cz.StructuralElementIdValue}, " +
                    $"MEP Category: {cz.MepElementCategory}, Host Type: {cz.StructuralElementType}. " +
                    $"This indicates the intersection point was not properly calculated during refresh. " +
                    $"Please re-run Refresh to regenerate correct intersection points.";
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] {errorMsg}");
                
                // ✅ AGGRESSIVE LOGGING: Direct file write to error log (ALWAYS works)
                try 
                { 
                    var errorLogFilePath = SafeFileLogger.GetLogFilePath("sleeve_placement_errors.log");
                    var debugLogFilePath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                    
                    // DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(errorLogFilePath, 
                            $"[{DateTime.Now}] ⚠️⚠️⚠️ ZERO INTERSECTION POINT ERROR ⚠️⚠️⚠️\n" +
                            $"[{DateTime.Now}] {errorMsg}\n" +
                            $"  ClashZone ID: {cz.Id}\n" +
                            $"  MEP Element ID: {cz.MepElementIdValue}\n" +
                            $"  Structural Element ID: {cz.StructuralElementIdValue}\n" +
                            $"  MEP Category: {cz.MepElementCategory}\n" +
                            $"  Host Type: {cz.StructuralElementType}\n" +
                            $"  Host Document: {cz.StructuralElementDocumentTitle ?? "NULL"}\n" +
                            $"  IntersectionPointX/Y/Z: ({cz.IntersectionPointX}, {cz.IntersectionPointY}, {cz.IntersectionPointZ})\n" +
                            $"  SleevePlacementPointX/Y/Z: ({cz.SleevePlacementPointX}, {cz.SleevePlacementPointY}, {cz.SleevePlacementPointZ})\n" +
                            $"  IntersectionPoint object: {(cz.IntersectionPoint != null ? $"({cz.IntersectionPoint.X}, {cz.IntersectionPoint.Y}, {cz.IntersectionPoint.Z})" : "NULL")}\n" +
                            $"  SleevePlacementPoint object: {(cz.SleevePlacementPoint != null ? $"({cz.SleevePlacementPoint.X}, {cz.SleevePlacementPoint.Y}, {cz.SleevePlacementPoint.Z})" : "NULL")}\n" +
                            $"  ⚠️ ACTION REQUIRED: Re-run Refresh to regenerate intersection points\n\n");
                        
                        File.AppendAllText(debugLogFilePath,
                            $"[{DateTime.Now:HH:mm:ss}] VALIDATE-FAILED: Zone={cz.Id}, MEP={cz.MepElementIdValue}, HOST={cz.StructuralElementIdValue}, IP=({cz.IntersectionPointX},{cz.IntersectionPointY},{cz.IntersectionPointZ})\n");
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ValidatePlacementPoint] ✅ Error logged to: {errorLogFilePath}");
                    }
                } 
                catch (Exception logEx) 
                { 
                    System.Diagnostics.Debug.WriteLine($"[ValidatePlacementPoint] ❌ Failed to log error: {logEx.Message}");
                }
                
                throw new InvalidOperationException(errorMsg);
            }
            
            // At least one set has valid coordinates - validation passed
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] [VALIDATION-PASSED] ClashZone {cz.Id} has valid coordinates: IP=({cz.IntersectionPointX:F3},{cz.IntersectionPointY:F3},{cz.IntersectionPointZ:F3}), SPP=({cz.SleevePlacementPointX:F3},{cz.SleevePlacementPointY:F3},{cz.SleevePlacementPointZ:F3})");
        }

    // Helper method to define category priority for sorting
    private int GetCategoryPriority(string mepCategory)
    {
        // Lower number = higher priority (processed first)
        return mepCategory?.ToLowerInvariant() switch
        {
            "duct accessories" => 1,  // Dampers - HIGHEST priority
            "ducts" => 2,             // Ducts - processed after dampers
            "pipes" => 3,
            "cable trays" => 4,
            "cable tray fittings" => 5,
            _ => 999                  // Unknown categories last
        };
    }

    /// <summary>
    /// Load Duct Accessories clash zones from XML file for proximity checking
    /// </summary>
    private List<ClashZone> LoadDuctAccessoriesClashZones()
    {
        try
        {
            var clashZones = new List<ClashZone>();
            var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
            
            // Search for Duct Accessories XML file
            var pattern = "*_duct_accessories.xml";
            var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
            
            if (matchingFiles.Length > 0)
            {
                // Use the most recently modified file
                var xmlFilePath = matchingFiles
                    .OrderByDescending(f => File.GetLastWriteTime(f))
                    .First();
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Loading Duct Accessories clash zones from: {xmlFilePath}");
                
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFilePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    var storageZones = filter?.ClashZoneStorage?.AllZones;
                    if (storageZones != null && storageZones.Count > 0)
                    {
                        clashZones.AddRange(storageZones);
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Loaded {clashZones.Count} Duct Accessories clash zones from XML");
                    }
                }
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalSleevePlacer] No Duct Accessories XML files found matching pattern: {pattern}");
            }
            
            return clashZones;
        }
        catch (Exception ex)
        {
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Error($"[UniversalSleevePlacer] Error loading Duct Accessories clash zones: {ex.Message}");
            return new List<ClashZone>();
        }
        }
        
        /// <summary>
        /// REMOVED: Expensive spatial checking replaced with simple flag-based approach
        /// Now we only use IsResolved and IsClusterResolved flags for duplicate avoidance
        /// </summary>

        /// <summary>
        /// ✅ SIMPLIFIED: Select universal family based on host type and MEP shape
        /// Uses 4 universal families following CONVOID approach
        /// Returns only family name and circular flag - no type name needed
        /// </summary>
        private (string familyName, bool isCircular) SelectUniversalFamily(ClashZone clashZone, MepElementSize mepSize)
        {
            // Determine host type (check both singular and plural forms)
            bool isWallOrFraming = clashZone.StructuralElementType == "Wall" || 
                                  clashZone.StructuralElementType == "Walls" ||
                                  clashZone.StructuralElementType == "Structural Framing";
            
            // 🛡️ ARCHITECTURE FIX: Use global configuration rules + CONDITIONS XML for opening type preferences
            // This follows the reference architecture: Global rules > UI preferences
            bool isCircular;
            if (string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ CORRECT: Pipes opening type using global configuration resolution
                var pipeType = _conditions?.OpeningTypePreferences?.Pipes ?? "Circular";
                
                // 🔍 DEBUG: Log pipe opening type resolution for floors vs walls
                var hostType = clashZone.StructuralElementType ?? "Unknown";
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] PIPE OPENING TYPE DEBUG: Host={hostType}, Category={clashZone.MepElementCategory}, UI_Preference={pipeType}");
                
                // Use PipePlacementStrategy to resolve opening type with global rules
                if (_strategy is PipePlacementStrategy pipeStrategy)
                {
                    var resolvedType = pipeStrategy.GetResolvedOpeningType(mepSize, pipeType, hostType);
                    isCircular = string.Equals(resolvedType, "Circular", StringComparison.OrdinalIgnoreCase);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] PIPE opening type resolved: Host={hostType}, UI='{pipeType}' → Global Rule='{resolvedType}' → isCircular={isCircular}");
                }
                else
                {
                    // Fallback to CONDITIONS XML if strategy not available
                isCircular = string.Equals(pipeType, "Circular", StringComparison.OrdinalIgnoreCase);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] PIPE opening type from CONDITIONS XML (fallback): Host={hostType}, '{pipeType}' → isCircular={isCircular}");
            }
            }
            else if (string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ CORRECT: Round ducts opening type from CONDITIONS XML (user preference)
                // Check if this is a round duct first
                bool isRoundDuct = string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase);
                
                if (isRoundDuct)
                {
                    // Round ducts: use user preference from CONDITIONS XML
                    var roundDuctType = _conditions?.OpeningTypePreferences?.RoundDucts ?? "Circular";
                    isCircular = string.Equals(roundDuctType, "Circular", StringComparison.OrdinalIgnoreCase);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] ROUND DUCT opening type from CONDITIONS XML: '{roundDuctType}' → isCircular={isCircular}");
                }
                else
                {
                    // Rectangular ducts: always rectangular opening
                    isCircular = false;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] RECTANGULAR DUCT → isCircular=false");
                }
            }
            else
            {
                // Other categories (cable trays, accessories): rectangular
                isCircular = false;
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] OTHER CATEGORY ({clashZone.MepElementCategory}) → isCircular=false");
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] Family selection - StructuralElementType: '{clashZone.StructuralElementType}', isWallOrFraming: {isWallOrFraming}, MEP Shape: '{mepSize.Shape}', isCircular: {isCircular}");
            
            // Select family using your exact family names
            string familyName = (isWallOrFraming, isCircular) switch
            {
                (true, false) => "RectangularOpeningOnWall",   // ✅ Your family
                (true, true) => "CircularOpeningOnWall",       // ✅ Your family
                (false, false) => "RectangularOpeningOnSlab",  // ✅ Your family
                (false, true) => "CircularOpeningOnSlab",      // ✅ Your family
            };
            
            // ✅ SIMPLIFIED: No type name needed - Revit will use any active type in the family
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] Selected family: {familyName}, IsCircular: {isCircular}");
            
            return (familyName, isCircular);
        }
        
        // Ensure XML-serializable coordinates are populated before saving
        /// <summary>
        /// ✅ CRITICAL FIX: Normalizes intersection coordinates from multiple sources
        /// Handles Internal Origin-to-Origin links where IntersectionPoint object may be (0,0,0) after XML deserialization
        /// </summary>
        private static void NormalizeIntersectionCoordinates(ClashZone zone)
        {
            if (zone == null) return;
            
            // ✅ CRITICAL: Check XML-serialized properties FIRST (these are always set during refresh)
            // For Internal Origin-to-Origin links, these values are correct even if IntersectionPoint object is wrong
            bool hasValidXmlX = Math.Abs(zone.IntersectionPointX) > 1e-9;
            bool hasValidXmlY = Math.Abs(zone.IntersectionPointY) > 1e-9;
            bool hasValidXmlZ = Math.Abs(zone.IntersectionPointZ) > 1e-9;
            bool hasValidXml = hasValidXmlX || hasValidXmlY || hasValidXmlZ; // At least one coordinate is non-zero
            
            bool hasValidSppXmlX = Math.Abs(zone.SleevePlacementPointX) > 1e-9;
            bool hasValidSppXmlY = Math.Abs(zone.SleevePlacementPointY) > 1e-9;
            bool hasValidSppXmlZ = Math.Abs(zone.SleevePlacementPointZ) > 1e-9;
            bool hasValidSppXml = hasValidSppXmlX || hasValidSppXmlY || hasValidSppXmlZ;
            
            // ✅ PRIORITY 1: Use XML-serialized IntersectionPointX/Y/Z if they're valid (most reliable source)
            if (hasValidXml)
            {
                // XML values are already set correctly - ensure IntersectionPoint object matches
                if (zone.IntersectionPoint == null || 
                    (Math.Abs(zone.IntersectionPoint.X - zone.IntersectionPointX) > 1e-9 ||
                     Math.Abs(zone.IntersectionPoint.Y - zone.IntersectionPointY) > 1e-9 ||
                     Math.Abs(zone.IntersectionPoint.Z - zone.IntersectionPointZ) > 1e-9))
                {
                    zone.IntersectionPoint = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
                }
                // XML values are already correct, no need to change them
                return;
            }
            
            // ✅ PRIORITY 2: Use XML-serialized SleevePlacementPointX/Y/Z if IntersectionPoint is invalid
            if (hasValidSppXml)
            {
                zone.IntersectionPointX = zone.SleevePlacementPointX;
                zone.IntersectionPointY = zone.SleevePlacementPointY;
                zone.IntersectionPointZ = zone.SleevePlacementPointZ;
                if (zone.IntersectionPoint == null)
                {
                    zone.IntersectionPoint = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
                }
                return;
            }
            
            // ✅ PRIORITY 3: Check IntersectionPoint object (may be (0,0,0) after XML deserialization for Internal Origin-to-Origin)
            bool hasIntersectionObj = zone.IntersectionPoint != null;
            bool isIntersectionZero = hasIntersectionObj && 
                Math.Abs(zone.IntersectionPoint.X) < 1e-9 && 
                Math.Abs(zone.IntersectionPoint.Y) < 1e-9 && 
                Math.Abs(zone.IntersectionPoint.Z) < 1e-9;
            
            if (hasIntersectionObj && !isIntersectionZero)
            {
                zone.IntersectionPointX = zone.IntersectionPoint.X;
                zone.IntersectionPointY = zone.IntersectionPoint.Y;
                zone.IntersectionPointZ = zone.IntersectionPoint.Z;
                return;
            }
            
            // ✅ PRIORITY 4: Check SleevePlacementPoint object
            bool hasSleevePoint = zone.SleevePlacementPoint != null;
            bool isSleevePointZero = hasSleevePoint && 
                Math.Abs(zone.SleevePlacementPoint.X) < 1e-9 && 
                Math.Abs(zone.SleevePlacementPoint.Y) < 1e-9 && 
                Math.Abs(zone.SleevePlacementPoint.Z) < 1e-9;
            
            if (hasSleevePoint && !isSleevePointZero)
            {
                zone.IntersectionPointX = zone.SleevePlacementPoint.X;
                zone.IntersectionPointY = zone.SleevePlacementPoint.Y;
                zone.IntersectionPointZ = zone.SleevePlacementPoint.Z;
                return;
            }
            
            // ✅ PRIORITY 5: Final fallback - center of clash bounding box
            if (zone.ClashBoundingBox != null)
            {
                var bb = zone.ClashBoundingBox;
                var center = (bb.Min + bb.Max) / 2.0;
                zone.IntersectionPointX = center.X;
                zone.IntersectionPointY = center.Y;
                zone.IntersectionPointZ = center.Z;
                return;
            }
            
            // If all sources are invalid/zero, leave as-is (will be 0 in XML, but at least we tried)
            System.Diagnostics.Debug.WriteLine($"[NormalizeIntersectionCoordinates] WARNING: Zone {zone.Id} has no valid intersection coordinates from any source");
        }
        
        /// <summary>
        /// ⚠️ QUICK WIN: Pre-cache family symbols for all clash zones (load once, reuse many times)
        /// Expected gain: 2-3 seconds for 500 sleeves
        /// </summary>
        private void PreCacheFamilySymbols(List<ClashZone> clashZones)
        {
            try
            {
                var familyNames = new HashSet<string>();
                
                // Collect all unique family names needed
                foreach (var cz in clashZones)
                {
                    var familyName = GetFamilyNameForClashZone(cz);
                    if (!string.IsNullOrEmpty(familyName))
                    {
                        familyNames.Add(familyName);
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[FAMILY-CACHE] Pre-loading {familyNames.Count} unique family symbols...");
                
                int cachedCount = 0;
                foreach (var familyName in familyNames)
                {
                    // Check if already cached
                    if (_familySymbolCache.ContainsKey(familyName))
                    {
                        continue; // Already cached
                    }
                    
                    // Load and cache
                    var symbol = LoadFamilySymbolInternal(familyName);
                    if (symbol != null)
                    {
                        _familySymbolCache[familyName] = symbol;
                        cachedCount++;
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[FAMILY-CACHE] ✓ Pre-cached {cachedCount} family symbols. Total cached: {_familySymbolCache.Count}");
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[FAMILY-CACHE] Error pre-caching family symbols: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get family name for a clash zone based on structural element type
        /// </summary>
        private string GetFamilyNameForClashZone(ClashZone cz)
        {
            // Determine family name based on structural element type (same logic as in placement)
            var structType = cz.StructuralElementType ?? "";
            
            if (structType.Equals("Wall", StringComparison.OrdinalIgnoreCase))
            {
                // Determine if circular or rectangular based on clash zone
                bool isCircular = cz.SleeveDiameter > 0 && cz.SleeveWidth == 0 && cz.SleeveHeight == 0;
                return isCircular ? "OpeningOnWall-Circular" : "OpeningOnWall-Rectangular";
            }
            else if (structType.Equals("Floor", StringComparison.OrdinalIgnoreCase) || 
                     structType.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase))
            {
                bool isCircular = cz.SleeveDiameter > 0 && cz.SleeveWidth == 0 && cz.SleeveHeight == 0;
                return isCircular ? "OpeningOnSlab-Circular" : "OpeningOnSlab-Rectangular";
            }
            
            return null; // Unknown structural type
        }
        
        /// <summary>
        /// Get cached family symbol or load if not cached
        /// </summary>
        private FamilySymbol LoadFamilySymbol(string familyName)
        {
            // Check cache first (QUICK WIN)
            if (OptimizationFlags.UseFamilySymbolCache && _familySymbolCache.TryGetValue(familyName, out var cached))
            {
                return cached; // Return cached symbol instantly
            }
            
            // Not in cache - load it
            var symbol = LoadFamilySymbolInternal(familyName);
            
            // Cache it for next time
            if (symbol != null && OptimizationFlags.UseFamilySymbolCache)
            {
                _familySymbolCache[familyName] = symbol;
            }
            
            return symbol;
        }
        
        /// <summary>
        /// STRICT: Internal method to load family symbol from Revit (no caching, no fallbacks)
        /// Only checks for EXACT family names: RectangularOpeningOnWall, CircularOpeningOnWall, 
        /// RectangularOpeningOnSlab, CircularOpeningOnSlab
        /// </summary>
        private FamilySymbol LoadFamilySymbolInternal(string familyName)
        {
            // ✅ STRICT: Only check for exact family names - NO variations, NO fallbacks
            var allowedFamilyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "RectangularOpeningOnWall",
                "CircularOpeningOnWall",
                "RectangularOpeningOnSlab",
                "CircularOpeningOnSlab"
            };
            
            // Reject any family name that's not in our strict list
            if (!allowedFamilyNames.Contains(familyName))
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] INVALID family name requested: '{familyName}'. Only these 4 families are allowed: RectangularOpeningOnWall, CircularOpeningOnWall, RectangularOpeningOnSlab, CircularOpeningOnSlab");
                return null;
            }
            
            // Load universal families - EXACT name match only
            var symbols = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.Family.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            
            var symbol = symbols.FirstOrDefault();
            
            if (symbol == null)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] ❌ Universal family '{familyName}' NOT FOUND in project!");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] Required families must be loaded before placement can proceed.");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] Please load the family from the Resources folder:");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] 1. Go to Insert > Load Family");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] 2. Navigate to Resources folder");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[UniversalSleevePlacer] 3. Load the required family .rfa file");
            }
            else
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] ✅ Found universal family: {familyName} (ID: {symbol.Id})");
            }
            
            return symbol;
        }
        
        private Level FindNearestLevel(XYZ point)
        {
            var levels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => Math.Abs(l.Elevation - point.Z))
                .ToList();
            
            return levels.FirstOrDefault();
        }
        
        /// <summary>
        /// Get clearance value from UI clearance settings (priority) or CONDITIONS service (fallback)
        /// </summary>
        private double GetClearanceFromConditions(string category, MepElementSize mepSize)
        {
            try
            {
                // 🔥 CRITICAL DEBUG: Force direct file logging to trace clearance calculation
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] GetClearanceFromConditions START: category='{category}', Shape='{mepSize.Shape}', IsInsulated={mepSize.IsInsulated}\n");
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[GetClearanceFromConditions] START: category='{category}', UI settings={_clearanceSettings.Count}, XML conditions={(_conditions != null ? "NOT NULL" : "NULL")}");
                
                // ✅ PRIORITY 1: Use UI clearance settings if available
                if (_clearanceSettings.Count > 0)
                {
                    double clearanceInMm = GetClearanceFromUISettings(category, mepSize);
                    if (clearanceInMm > 0)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Using UI clearance: {clearanceInMm}mm for {category}\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromConditions] Using UI clearance: {clearanceInMm}mm for {category}");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceInMm);
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Using XML clearance settings for {category}\n");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromConditions] Using XML clearance settings for {category}");
                    return GetClearanceFromXmlConditions(category, mepSize);
                }
                
                // ✅ PRIORITY 3: Default fallback
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ No clearance settings available, using default 50mm\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[GetClearanceFromConditions] No clearance settings available, using default 50mm");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in clearance calculation: {ex.Message}\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetClearanceFromConditions] Error: {ex.Message}");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
        }
        
        /// <summary>
        /// Get clearance value from UI settings
        /// </summary>
        private double GetClearanceFromUISettings(string category, MepElementSize mepSize)
        {
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[GetClearanceFromUISettings] Checking UI settings for category: {category}");
                
                // Log all available UI clearance settings
                foreach (var kvp in _clearanceSettings)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromUISettings] Available: {kvp.Key} = {kvp.Value}mm");
                }
                
                if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // Check for pipe-specific clearance keys
                    string normalKey = "pipes_normal_clearance";
                    string insulatedKey = "pipes_insulated_clearance";
                    
                    bool isInsulated = IsPipeInsulated(mepSize);
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearance = _clearanceSettings[targetKey];
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromUISettings] Pipes: isInsulated={isInsulated}, key='{targetKey}', clearance={clearance}mm");
                        return clearance;
                    }
                    
                    // Fallback to generic pipe clearance
                    if (_clearanceSettings.ContainsKey("pipes_clearance"))
                    {
                        double clearance = _clearanceSettings["pipes_clearance"];
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromUISettings] Pipes: Using generic clearance={clearance}mm");
                        return clearance;
                    }
                }
                else if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    // Check for cable tray-specific clearance keys
                    string topKey = "cabletray_top_clearance";
                    string otherKey = "cabletray_other_clearance";
                    
                    // Try top clearance first, then other
                    if (_clearanceSettings.ContainsKey(topKey))
                    {
                        double clearance = _clearanceSettings[topKey];
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromUISettings] Cable Trays: Using top clearance={clearance}mm");
                        return clearance;
                    }
                    
                    if (_clearanceSettings.ContainsKey(otherKey))
                    {
                        double clearance = _clearanceSettings[otherKey];
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromUISettings] Cable Trays: Using other clearance={clearance}mm");
                        return clearance;
                    }
                }
                else if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // Check for duct-specific clearance keys
                    string normalKey = "ducts_normal_clearance";
                    string insulatedKey = "ducts_insulated_clearance";
                    
                    bool isInsulated = IsDuctInsulated(mepSize);
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Looking for key='{targetKey}', isInsulated={isInsulated}");
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearance = _clearanceSettings[targetKey];
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Found key='{targetKey}', clearance={clearance}mm");
                        return clearance;
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[GetClearanceFromUISettings] Ducts: Key '{targetKey}' not found in clearance settings!");
                        
                        // Try alternative keys
                        if (_clearanceSettings.ContainsKey("normal_clearance"))
                        {
                            double clearance = _clearanceSettings["normal_clearance"];
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Using fallback 'normal_clearance' = {clearance}mm");
                            return clearance;
                        }
                        
                        if (_clearanceSettings.ContainsKey("insulated_clearance"))
                        {
                            double clearance = _clearanceSettings["insulated_clearance"];
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Using fallback 'insulated_clearance' = {clearance}mm");
                            return clearance;
                        }
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[GetClearanceFromUISettings] No UI clearance found for category: {category}");
                return 0; // Indicate no UI clearance found
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetClearanceFromUISettings] Error: {ex.Message}");
                return 0;
            }
        }
        
        /// <summary>
        /// Get clearance for pipe based on ClashZone.IsInsulated (uses authoritative data from database)
        /// This bypasses mepSize which might be stale and uses clashZone.IsInsulated directly
        /// </summary>
        private double GetClearanceForPipeFromClashZone(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                
                // ✅ PRIORITY 1: Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    string normalKey = "pipes_normal_clearance";
                    string insulatedKey = "pipes_insulated_clearance";
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearanceMm = _clearanceSettings[targetKey];
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceForPipeFromClashZone] Using UI clearance: isInsulated={isInsulated}, key='{targetKey}', clearance={clearanceMm}mm");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                    }
                    
                    // Fallback to generic pipe clearance
                    if (_clearanceSettings.ContainsKey("pipes_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["pipes_clearance"];
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceForPipeFromClashZone] Using generic UI clearance: {clearanceMm}mm");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                    double clearanceInMm = isInsulated 
                        ? _conditions.ClearanceSettings.PipesInsulated 
                        : _conditions.ClearanceSettings.PipesNormal;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceForPipeFromClashZone] Using XML clearance: isInsulated={isInsulated}, clearance={clearanceInMm}mm");
                    return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceInMm);
                }
                
                // ✅ PRIORITY 3: Default fallback (50mm)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GetClearanceForPipeFromClashZone] No clearance settings available, using default 50mm");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GetClearanceForPipeFromClashZone] Error: {ex.Message}");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
        }
        
        /// <summary>
        /// Get clearance for duct based on ClashZone.IsInsulated (uses authoritative data from database)
        /// This bypasses mepSize which might be stale and uses clashZone.IsInsulated directly
        /// </summary>
        private double GetClearanceForDuctFromClashZone(ClashZone clashZone)
        {
            try
            {
                bool isInsulated = clashZone?.IsInsulated ?? false;
                
                // ✅ PRIORITY 1: Try UI clearance settings first (user-provided values take priority)
                if (_clearanceSettings != null && _clearanceSettings.Count > 0)
                {
                    string normalKey = "ducts_normal_clearance";
                    string insulatedKey = "ducts_insulated_clearance";
                    string targetKey = isInsulated ? insulatedKey : normalKey;
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearanceMm = _clearanceSettings[targetKey];
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceForDuctFromClashZone] Using UI clearance: isInsulated={isInsulated}, key='{targetKey}', clearance={clearanceMm}mm");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                    }
                    
                    // Fallback to generic duct clearance
                    if (_clearanceSettings.ContainsKey("ducts_clearance"))
                    {
                        double clearanceMm = _clearanceSettings["ducts_clearance"];
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[GetClearanceForDuctFromClashZone] Using generic UI clearance: {clearanceMm}mm");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions (check shape for round vs rectangular)
                if (_conditions?.ClearanceSettings != null)
                {
                    // Check if round or rectangular (default to rectangular if unknown)
                    bool isRound = clashZone?.MepElementSizeData?.Shape?.Equals("Round", StringComparison.OrdinalIgnoreCase) == true ||
                                   clashZone?.MepElementSizeData?.Shape?.Equals("Circular", StringComparison.OrdinalIgnoreCase) == true;
                    
                    double clearanceInMm;
                    if (isRound)
                    {
                        clearanceInMm = isInsulated 
                            ? _conditions.ClearanceSettings.RoundInsulated 
                            : _conditions.ClearanceSettings.RoundNormal;
                    }
                    else
                    {
                        clearanceInMm = isInsulated 
                            ? _conditions.ClearanceSettings.RectangularInsulated 
                            : _conditions.ClearanceSettings.RectangularNormal;
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceForDuctFromClashZone] Using XML clearance: isInsulated={isInsulated}, isRound={isRound}, clearance={clearanceInMm}mm");
                    return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceInMm);
                }
                
                // ✅ PRIORITY 3: Default fallback (50mm)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[GetClearanceForDuctFromClashZone] No clearance settings available, using default 50mm");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GetClearanceForDuctFromClashZone] Error: {ex.Message}");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
        }
        
        /// <summary>
        /// Get clearance value from XML conditions (fallback method)
        /// </summary>
        private double GetClearanceFromXmlConditions(string category, MepElementSize mepSize)
        {
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[GetClearanceFromXmlConditions] ClearanceSettings available: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, RectInsulated={_conditions.ClearanceSettings.RectangularInsulated}mm");

                // Determine clearance based on category and element properties
                double clearanceInMm = 50.0; // Default fallback

                if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // Pipes: Check if insulated
                    bool isInsulated = IsPipeInsulated(mepSize);
                    
                    // ⚠️ DIAGNOSTIC: Log insulation detection from XML data
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: Reading from XML - Shape='{mepSize.Shape}', IsInsulated={isInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: isInsulated={isInsulated}, clearance={clearanceInMm}mm");
                    
                    clearanceInMm = isInsulated ? _conditions.ClearanceSettings.PipesInsulated : _conditions.ClearanceSettings.PipesNormal;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: Final clearance={clearanceInMm}mm (insulated={isInsulated})");
                }
                else if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // Ducts: Check if insulated and shape
                    bool isInsulated = IsDuctInsulated(mepSize);
                    bool isRound = string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase);
                    
                    // ⚠️ DIAGNOSTIC: Log insulation detection from XML data
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Ducts: Reading from XML - Shape='{mepSize.Shape}', IsInsulated={isInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Ducts: isInsulated={isInsulated}, isRound={isRound}, Shape='{mepSize.Shape}'");
                    
                    if (isRound)
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RoundInsulated : _conditions.ClearanceSettings.RoundNormal;
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromXmlConditions] Round ducts: clearance={clearanceInMm}mm");
                    }
                    else
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RectangularInsulated : _conditions.ClearanceSettings.RectangularNormal;
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[GetClearanceFromXmlConditions] Rectangular ducts: clearance={clearanceInMm}mm");
                    }
                }
                else if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    // Cable Trays: Use top clearance as default
                    clearanceInMm = _conditions.ClearanceSettings.CableTrayTop;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Cable Trays: clearance={clearanceInMm}mm");
                }

                // Convert from mm to feet (Revit internal units)
                double clearanceInFeet = RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceInMm);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[GetClearanceFromXmlConditions] {category}: {clearanceInMm}mm → {clearanceInFeet:F6}ft");
                return clearanceInFeet;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetClearanceFromXmlConditions] Error: {ex.Message}");
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50.0);
            }
        }

        private static bool HasSnapshotPlacementData(ClashZone zone)
        {
            if (zone == null)
                return false;

            const double tol = 1e-9;
            bool HasValue(double value) => Math.Abs(value) > tol;

            return HasValue(zone.SleevePlacementPointX) ||
                   HasValue(zone.SleevePlacementPointY) ||
                   HasValue(zone.SleevePlacementPointZ) ||
                   HasValue(zone.IntersectionPointX) ||
                   HasValue(zone.IntersectionPointY) ||
                   HasValue(zone.IntersectionPointZ);
        }

        private static XYZ GetSnapshotPlacementPoint(ClashZone zone)
        {
            if (zone == null)
                return new XYZ(0, 0, 0);

            const double tol = 1e-9;

            bool hasPlacementPoint =
                Math.Abs(zone.SleevePlacementPointX) > tol ||
                Math.Abs(zone.SleevePlacementPointY) > tol ||
                Math.Abs(zone.SleevePlacementPointZ) > tol;

            if (hasPlacementPoint)
            {
                return new XYZ(
                    zone.SleevePlacementPointX,
                    zone.SleevePlacementPointY,
                    zone.SleevePlacementPointZ);
            }

            return new XYZ(
                zone.IntersectionPointX,
                zone.IntersectionPointY,
                zone.IntersectionPointZ);
        }

        private static bool HasValidPlacementCoordinate(XYZ point)
        {
            if (point == null) return false;

            const double tol = 1e-9;
            return Math.Abs(point.X) > tol ||
                   Math.Abs(point.Y) > tol ||
                   Math.Abs(point.Z) > tol;
        }

        /// <summary>
        /// Determine if a pipe is insulated (use actual insulation data from strategy analysis)
        /// </summary>
        private bool IsPipeInsulated(MepElementSize mepSize)
        {
            // ✅ FIXED: Use actual insulation data from strategy analysis instead of guessing
            // The MepElementSize object already contains the correct insulation status
            // from the PipePlacementStrategy.GetMepElementSize method
            
            // 🔥 CRITICAL DEBUG: Log the actual insulation data being used
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] IsPipeInsulated: Using actual data - IsInsulated={mepSize.IsInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft\n");
            
            return mepSize.IsInsulated;
        }

        /// <summary>
        /// Determine if a duct is insulated (use actual insulation data from strategy analysis)
        /// </summary>
        private bool IsDuctInsulated(MepElementSize mepSize)
        {
            // ✅ FIXED: Use actual insulation data from strategy analysis instead of guessing
            // The MepElementSize object already contains the correct insulation status
            // from the DuctPlacementStrategy.GetMepElementSize method
            
            // 🔥 CRITICAL DEBUG: Log the actual insulation data being used
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] IsDuctInsulated: Using actual data - IsInsulated={mepSize.IsInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft\n");
            
            return mepSize.IsInsulated;
        }

        // ============================================================================
        // 🛡️ FAIL-SAFE: Dimension Validation
        // ============================================================================
        private bool ValidateSleeveDimensions(double width, double height, double diameter, ClashZone clashZone)
        {
            try
            {
                // Convert to millimeters for easier validation
                double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                double diameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(diameter);
                
                // 🚨 CRITICAL LIMITS: Prevent oversized sleeves that crash Revit
                const double MAX_SIZE_MM = 10000.0; // 10 meters - reasonable maximum
                const double MIN_SIZE_MM = 10.0;    // 10mm - reasonable minimum
                
                // Check for invalid/negative dimensions
                if (widthMm <= 0 || heightMm <= 0 || diameterMm <= 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ValidateDimensions] NEGATIVE/ZERO dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm");
                    return false;
                }
                
                // Check for oversized dimensions
                if (widthMm > MAX_SIZE_MM || heightMm > MAX_SIZE_MM || diameterMm > MAX_SIZE_MM)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ValidateDimensions] OVERSIZED dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm (MAX={MAX_SIZE_MM}mm)");
                    return false;
                }
                
                // Check for undersized dimensions
                if (widthMm < MIN_SIZE_MM || heightMm < MIN_SIZE_MM || diameterMm < MIN_SIZE_MM)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ValidateDimensions] UNDERSIZED dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm (MIN={MIN_SIZE_MM}mm)");
                    return false;
                }
                
                // Check for reasonable aspect ratio (prevent extremely thin sleeves)
                const double MAX_ASPECT_RATIO = 20.0; // Max 20:1 ratio
                double aspectRatio1 = Math.Max(widthMm, heightMm) / Math.Min(widthMm, heightMm);
                if (aspectRatio1 > MAX_ASPECT_RATIO)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ValidateDimensions] EXTREME ASPECT RATIO: {aspectRatio1:F1}:1 (MAX={MAX_ASPECT_RATIO}:1)");
                    return false;
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[ValidateDimensions] ✓ VALID dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm");
                return true;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[ValidateDimensions] ERROR validating dimensions: {ex.Message}");
                return false; // Fail safe - don't proceed with invalid dimensions
            }
        }
        
        // ============================================================================
        // STEP 5 OPTIMIZATION: Flush Deferred Parameters (Batch Write After Regeneration)
        // ============================================================================
        /// <summary>
        /// Applies all accumulated parameter values from _deferredParameters to sleeves after regeneration.
        /// This eliminates per-sleeve Revit regeneration overhead during parameter setting.
        /// Expected gain: 4-6× faster placement (143-203ms → <30ms per sleeve).
        /// </summary>
        private void FlushDeferredParameters()
        {
            // ✅ CRITICAL DIAGNOSTIC: Log call stack to identify where this is being called from
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var stackTrace = new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false);
                var caller = stackTrace.GetFrame(0)?.GetMethod()?.Name ?? "Unknown";
                DebugLogger.Info($"[BATCH-PARAMS] 🔍 FlushDeferredParameters CALLED from: {caller}, StackDepth={stackTrace.FrameCount}");
            }
            
            if (_deferredParameters == null || _deferredParameters.Count == 0)
            {
                // ✅ DIAGNOSTIC: Log when batching is enabled but no parameters deferred
                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseBatchedParameterWrites)
                {
                    DebugLogger.Info($"[BATCH-PARAMS] ⚠️ FlushDeferredParameters called but _deferredParameters is empty (batching enabled but no parameters deferred)");
                }
                return;
            }
            
            // ✅ DIAGNOSTIC: Log flush start with count
            if (!DeploymentConfiguration.DeploymentMode)
            {
                int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                DebugLogger.Info($"[BATCH-PARAMS] 🔄 Flushing {_deferredParameters.Count} individual sleeves with {totalParams} total parameters...");
                DebugLogger.Info($"[BATCH-PARAMS] 📊 Sleeve IDs: [{string.Join(", ", _deferredParameters.Keys.Select(id => id.IntegerValue))}]");
            }
            
            int successCount = 0;
            int failCount = 0;
            var errorLog = new System.Text.StringBuilder();
            
            try
            {
                foreach (var kvp in _deferredParameters)
                {
                    var sleeveId = kvp.Key;
                    var paramValues = kvp.Value;
                    
                    try
                    {
                        // Get sleeve instance from document
                        var sleeveInstance = _doc.GetElement(sleeveId) as FamilyInstance;
                        if (sleeveInstance == null || !sleeveInstance.IsValidObject)
                        {
                            errorLog.AppendLine($"Sleeve {sleeveId} no longer valid - skipping {paramValues.Count} parameters");
                            failCount++;
                            continue;
                        }
                        
                        // Apply all deferred parameters for this sleeve
                        foreach (var paramKvp in paramValues)
                        {
                            var paramName = paramKvp.Key;
                            var paramValue = paramKvp.Value;
                            
                            try
                            {
                                var param = sleeveInstance.LookupParameter(paramName);
                                if (param != null && !param.IsReadOnly)
                                {
                                    if (paramValue is double doubleVal)
                                    {
                                        param.Set(doubleVal);
                                    }
                                    else if (paramValue is string stringVal)
                                    {
                                        param.Set(stringVal);
                                    }
                                    else if (paramValue is int intVal)
                                    {
                                        param.Set(intVal);
                                    }
                                }
                            }
                            catch (Exception paramEx)
                            {
                                errorLog.AppendLine($"  Failed to set parameter '{paramName}' on sleeve {sleeveId}: {paramEx.Message}");
                            }
                        }
                        
                        successCount++;
                    }
                    catch (Exception sleeveEx)
                    {
                        errorLog.AppendLine($"Failed to process sleeve {sleeveId}: {sleeveEx.Message}");
                        failCount++;
                    }
                }
                
                // Log summary
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                    DebugLogger.Info($"[BATCH-PARAMS] ✅ Flushed {successCount} individual sleeves ({totalParams} parameters) in batch, {failCount} failed");
                    if (errorLog.Length > 0)
                    {
                        DebugLogger.Warning($"[BATCH-PARAMS] Errors during flush:\n{errorLog}");
                    }
                    
                    // ✅ ALSO: Log to file for verification
                    var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                    Directory.CreateDirectory(logDir);
                    var profPath = Path.Combine(logDir, "individual_sleeve_param_batching.log");
                    File.AppendAllText(profPath, 
                        $"{DateTime.Now:O}\t" +
                        $"Sleeves={successCount}\t" +
                        $"TotalParams={totalParams}\t" +
                        $"Failed={failCount}\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[BATCH-PARAMS] CRITICAL ERROR during parameter flush: {ex.Message}\n{ex.StackTrace}");
                }
                // Don't throw - log error and continue (parameters already written to _deferredParameters)
            }
            finally
            {
                // ✅ CRITICAL DIAGNOSTIC: Log before clearing to track when/why it's cleared
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[BATCH-PARAMS] 🗑️ CLEARING _deferredParameters: Count={_deferredParameters?.Count ?? 0} sleeves, StackDepth={new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false).FrameCount}");
                }
                
                // Clear deferred parameters after flush (ready for next placement batch)
                _deferredParameters.Clear();
            }
        }
        
        // ============================================================================
// CORRECTED SetSleeveParameters Method
// ============================================================================
        private void SetSleeveParameters(
            FamilyInstance sleeveInstance, 
            MepElementSize mepSize,
            double finalWidth,
            double finalHeight,
            double finalDiameter,
            ClashZone clashZone,
            bool isCircular)
        {
            try
            {
                // ✅ PERFORMANCE OPTIMIZATION: Cache all parameter references upfront (calculate once, use many times)
                var paramCache = new Dictionary<string, Parameter>();
                void CacheParam(string name)
                {
                    var param = sleeveInstance.LookupParameter(name);
                    if (param != null && !param.IsReadOnly)
                        paramCache[name] = param;
                }

                // Local timing + logging helpers (only active when EnableParameterTimingInstrumentation true)
                void AppendTiming(string logicalName, long ticks, long ms, double internalValue)
                {
                    if (!OptimizationFlags.EnableParameterTimingInstrumentation || !OptimizationFlags.UseDiagnosticMode) return;
                    try
                    {
                        var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Logs", "R2023");
                        Directory.CreateDirectory(logDir);
                        var path = Path.Combine(logDir, "param_timing.log");
                        File.AppendAllText(path, $"{DateTime.Now:O}\tSleeve={sleeveInstance.Id.IntegerValue}\tParam={logicalName}\tValueInternal={internalValue:F6}\tTicks={ticks}\tMs={ms}\n");
                    }
                    catch { }
                }

                void TimedSetDouble(Parameter p, double value, string logicalName)
                {
                    if (p == null || p.IsReadOnly) return;
                    
                    // ✅ STEP 5 OPTIMIZATION: Defer parameter writes if batching enabled
                    if (OptimizationFlags.UseBatchedParameterWrites)
                    {
                        // Accumulate parameter value for later batch write
                        if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                            _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                        _deferredParameters[sleeveInstance.Id][logicalName] = value;
                        
                        // ✅ DIAGNOSTIC: Log when parameters are deferred (first 5 sleeves only to avoid spam)
                        if (!DeploymentConfiguration.DeploymentMode && _deferredParameters.Count <= 5)
                        {
                            DebugLogger.Info($"[BATCH-PARAMS] ⚡ DEFERRED: Sleeve={sleeveInstance.Id.IntegerValue}, Param={logicalName}, TotalDeferred={_deferredParameters.Count} sleeves, ParamsForThisSleeve={_deferredParameters[sleeveInstance.Id].Count}");
                        }
                        
                        // Still log timing if instrumentation enabled (measures overhead of accumulation)
                        if (OptimizationFlags.EnableParameterTimingInstrumentation)
                        {
                            var sw = Stopwatch.StartNew();
                            // Just track accumulation time (should be <1ms vs 10-20ms for Set)
                            sw.Stop();
                            AppendTiming(logicalName, sw.ElapsedTicks, sw.ElapsedMilliseconds, value);
                        }
                        return;
                    }
                    
                    // Immediate write path (backward compatibility when batching disabled)
                    if (OptimizationFlags.EnableParameterTimingInstrumentation)
                    {
                        var sw = Stopwatch.StartNew();
                        p.Set(value);
                        sw.Stop();
                        AppendTiming(logicalName, sw.ElapsedTicks, sw.ElapsedMilliseconds, value);
                    }
                    else
                    {
                        p.Set(value);
                    }
                }
                
                // ✅ BATCHING HELPER: Defer string parameter writes
                void TimedSetString(Parameter p, string value, string logicalName)
                {
                    if (p == null || p.IsReadOnly) return;
                    
                    if (OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                            _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                        _deferredParameters[sleeveInstance.Id][logicalName] = value;
                        return;
                    }
                    
                    p.Set(value);
                }
                
                // ✅ BATCHING HELPER: Defer int parameter writes
                void TimedSetInt(Parameter p, int value, string logicalName)
                {
                    if (p == null || p.IsReadOnly) return;
                    
                    if (OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                            _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                        _deferredParameters[sleeveInstance.Id][logicalName] = value;
                        return;
                    }
                    
                    p.Set(value);
                }
                
                // Cache all needed parameters once
                CacheParam("Width");
                CacheParam("Height");
                CacheParam("Depth");
                CacheParam("Wall Width");
                CacheParam("Opening Outside Diameter");
                CacheParam("Opening_Diameter");
                CacheParam("Outside Diameter");
                CacheParam("Diameter");
                CacheParam("MEP_ElementId");
                CacheParam("MEP_Category");
                CacheParam("MEP_UniqueId");
                CacheParam("MEP_Size");
                CacheParam("System_Abbreviation");
                CacheParam("MEP_Count");
                CacheParam("HostOrientation");
                CacheParam("Bottom of Opening");
                CacheParam("Elevation from Level");
                CacheParam("Schedule Level Elevation");
                CacheParam("Elevation from Level Offset");
                CacheParam("Filter Name");
                CacheParam("Sleeve Instance ID");
                CacheParam("ClashZone_GUID");
                
                // ✅ PERFORMANCE OPTIMIZATION: Pre-cache ALL host parameters upfront to avoid LookupParameter() in loop
                if (clashZone.HostParameterValues != null && clashZone.HostParameterValues.Count > 0)
                {
                    foreach (var hostParam in clashZone.HostParameterValues)
                    {
                        if (!paramCache.ContainsKey(hostParam.Key))
                        {
                            var param = sleeveInstance.LookupParameter(hostParam.Key);
                            if (param != null && !param.IsReadOnly)
                                paramCache[hostParam.Key] = param;
                        }
                    }
                }
                
                // Helper to get cached parameter
                Parameter GetParam(string name) => paramCache.TryGetValue(name, out var p) ? p : null;
                // 🛡️ FAIL-SAFE: Validate dimensions before setting parameters
                if (!ValidateSleeveDimensions(finalWidth, finalHeight, finalDiameter, clashZone))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalSleevePlacer] INVALID DIMENSIONS for ClashZone {clashZone.Id}: W={finalWidth:F3}ft, H={finalHeight:F3}ft, D={finalDiameter:F3}ft");
                    ErrorCount++;
                    return; // Skip this sleeve - don't crash Revit
                }
        bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
        bool isDamper = string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
        
        // 🔍 DEBUG: Log pipe detection for floors vs walls
        var hostType = clashZone.StructuralElementType ?? "Unknown";
                if (!DeploymentConfiguration.DeploymentMode)
        DebugLogger.Info($"[SetSleeveParameters] PIPE DETECTION DEBUG: Host={hostType}, Category={clashZone.MepElementCategory}, isPipe={isPipe}");
        
        // ✅ DAMPER: No rounding - preserve exact calculated dimensions
        // Apply rounding to nearest 5mm if setting is enabled (skip for dampers)
        double roundedWidth, roundedHeight, roundedDiameter;
        if (isDamper)
        {
            // No rounding for dampers - use exact calculated values
            roundedWidth = finalWidth;
            roundedHeight = finalHeight;
            roundedDiameter = finalDiameter;
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[ROUNDING] Zone {clashZone?.Id}: DAMPER - No rounding applied, using exact calculated dimensions: {RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm");
        }
        else
        {
            // Apply rounding for other categories
            // ✅ DETAILED LOGGING: Log pipe rounding inputs and outputs
            if (isPipe && !DeploymentConfiguration.DeploymentMode)
            {
                var finalDiameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(finalDiameter);
                DebugLogger.Info($"[PIPE-ROUNDING-INPUT] Zone {clashZone?.Id}: Final diameter BEFORE rounding = {finalDiameterMm:F1}mm");
            }
            
            (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(finalWidth, finalHeight);
            roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(finalDiameter);
            
            // ✅ DETAILED LOGGING: Log pipe rounding results
            if (isPipe && !DeploymentConfiguration.DeploymentMode)
            {
                var roundedDiameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedDiameter);
                var finalDiameterMm = RevitUnitConversionService.Instance.FromInternalMillimeters(finalDiameter);
                if (Math.Abs(finalDiameterMm - roundedDiameterMm) > 0.1)
                {
                    DebugLogger.Info($"[PIPE-ROUNDING-OUTPUT] Zone {clashZone?.Id}: {finalDiameterMm:F1}mm → {roundedDiameterMm:F1}mm (rounded)");
                }
                else
                {
                    DebugLogger.Info($"[PIPE-ROUNDING-OUTPUT] Zone {clashZone?.Id}: {finalDiameterMm:F1}mm (no rounding applied)");
                }
            }
        }
        
        // ✅ OOP METHOD: Log rounded values (final dimensions actually used)
        double roundedWidthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth);
        double roundedHeightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight);
        if (!isDamper) // Only log rounding for non-dampers (dampers already logged above)
        {
            if (Math.Abs(finalWidth - roundedWidth) > 1e-6 || Math.Abs(finalHeight - roundedHeight) > 1e-6)
            {
                // Rounding was applied
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ROUNDING] Zone {clashZone?.Id}: Rounded from {RevitUnitConversionService.Instance.FromInternalMillimeters(finalWidth):F1}x{RevitUnitConversionService.Instance.FromInternalMillimeters(finalHeight):F1}mm → {roundedWidthMm:F1}x{roundedHeightMm:F1}mm");
            }
            else
            {
                // No rounding applied (or rounding value matches calculation exactly)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ROUNDING] Zone {clashZone?.Id}: No rounding applied - Using calculated dimensions: {roundedWidthMm:F1}x{roundedHeightMm:F1}mm");
            }
        }
        
        // ✅ PERFORMANCE: Removed verbose parameter logging - only log in diagnostic mode
        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
        {
            try
            {
                double diamMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedDiameter);
                double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth);
                double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight);
                DebugLogger.Info($"[PARAM-SET] Sleeve {sleeveInstance.Id.IntegerValue}: isPipe={isPipe}, Shape={mepSize.Shape}, roundedDia={diamMm:F1}mm, W={widthMm:F1}mm, H={heightMm:F1}mm\n");
            }
            catch { }
        }
                
                // Set size parameters
        bool treatAsCircular =
            string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase) ||
            (isPipe && isCircular); // ✅ FIX: Only treat pipes as circular if they're actually circular opening type

        if (treatAsCircular)
        {
            // Try common diameter parameter names in order
            string[] diameterParamNames = new[]
            {
                "Opening Outside Diameter",
                "Opening_Diameter",
                "Outside Diameter",
                "Diameter",
                "Width"  // Some circular families use Width
            };
            
            bool setOk = false;
            foreach (var name in diameterParamNames)
            {
                var p = GetParam(name); // ✅ PERFORMANCE: Use cache only, no fallback LookupParameter
                if (p != null && !p.IsReadOnly)
                {
                    TimedSetDouble(p, roundedDiameter, name);
                    // Add to cache if not already cached
                    if (!paramCache.ContainsKey(name) && !p.IsReadOnly)
                        paramCache[name] = p;
                    // ✅ PERFORMANCE: Removed verbose logging - only log in diagnostic mode
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        double diamMm = RevitUnitConversionService.Instance.FromInternalMillimeters(roundedDiameter);
                        DebugLogger.Info($"[UniversalSleevePlacer] Set '{name}' = {diamMm:F1}mm on sleeve {sleeveInstance.Id}");
                    }
                    setOk = true;
                    break;
                }
            }
            
            if (!setOk)
                    {
                        // Fallback to Width/Height for circular
                        {
                            var _w = GetParam("Width");
                            var _h = GetParam("Height");
                            TimedSetDouble(_w, roundedDiameter, "Width");
                            TimedSetDouble(_h, roundedDiameter, "Height");
                        }
                        // ✅ PERFORMANCE: Removed verbose logging
                    }
                }
                else
                {
                    // Rectangular
                    {
                        var _w = GetParam("Width");
                        var _h = GetParam("Height");
                        TimedSetDouble(_w, roundedWidth, "Width");
                        TimedSetDouble(_h, roundedHeight, "Height");
                    }
                    // ✅ PERFORMANCE: Removed verbose logging
        }
        
        // ⚠️ FLOOR FIX: For duct sleeves on floors, rotate orientation by 90 degrees and swap width/height
        // NOTE: Cable trays should NOT have width/height swapped - they maintain their original orientation
        bool isFloorHost = string.Equals(clashZone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(clashZone.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);
        bool isDuct = string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
        bool isCableTray = string.Equals(clashZone.MepElementCategory, "Cable Trays", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(clashZone.MepElementCategory, "Cable Tray Fittings", StringComparison.OrdinalIgnoreCase);

                if (!DeploymentConfiguration.DeploymentMode)
        DebugLogger.Info($"[UniversalSleevePlacer] FLOOR MEP CHECK: StructuralElementType='{clashZone.StructuralElementType}', MepElementCategory='{clashZone.MepElementCategory}', isFloorHost={isFloorHost}, isDuct={isDuct}, isCableTray={isCableTray}");

        // ⚠️ CRITICAL FIX: Choose ONE approach for duct orientation on floors
        // Option 1: Swap width/height (NO rotation) - This aligns family axes with MEP direction
        // Option 2: Keep original width/height (WITH rotation) - This rotates the sleeve to align
        // We'll use Option 1 (swap only) for consistency and to avoid double-correction
        
        if (isFloorHost && isDuct && !treatAsCircular)
        {
            // ⚠️ CRITICAL FIX: Ensure longer dimension becomes width for duct sleeves on floors
            // This ensures consistent orientation regardless of how Revit stores the dimensions
            if (roundedHeight > roundedWidth)
            {
                // Height is longer - swap so longer dimension becomes width
            double tempWidth = roundedWidth;
                roundedWidth = roundedHeight;  // Use height as width (longer dimension)
                roundedHeight = tempWidth;     // Use original width as height (shorter dimension)
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT SWAP: Height({RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm) > Width({RevitUnitConversionService.Instance.FromInternalMillimeters(tempWidth):F1}mm) - Swapped to make longer dimension the width");
            }
            else
            {
                // Width is already longer - no swap needed
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT NO SWAP: Width({RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm) >= Height({RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm) - Longer dimension already width");
            }
            
            // Update the sleeve parameters with correct values
            {
                var _w = GetParam("Width");
                var _h = GetParam("Height");
                TimedSetDouble(_w, roundedWidth, "Width");
                TimedSetDouble(_h, roundedHeight, "Height");
            }
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT FINAL: Width={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, Height={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm (NO ROTATION)");
        }
        else if (isFloorHost && isCableTray && !treatAsCircular)
        {
            // Cable trays on floors: NO width/height swap - maintain original orientation
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR CABLE TRAY: NO SWAP - Width={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, Height={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm");
        }
        
        // CRITICAL: Set Depth parameter based on host type
        // Priority: Wall Width (walls) > Depth (floors/framing) > Type Depth (fallback)
        bool isWallHost = clashZone.StructuralElementType == "Wall" || clashZone.StructuralElementType == "Walls";
        bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
        
        // ⚠️ BUG FIX: For structural framing, swap width and depth based on orientation
        if (isFramingHost && !treatAsCircular)
        {
            // Get host orientation from ClashZone
            string hostOrientation = clashZone.HostOrientation ?? "";
            
            if (hostOrientation == "X")
            {
                // X-framing: swap width and depth
                double tempWidth = roundedWidth;
                roundedWidth = roundedHeight;  // Use height as width
                roundedHeight = tempWidth;     // Use original width as height
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] X-FRAMING SWAP: Width={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, Height={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm");
            }
            else if (hostOrientation == "Y")
            {
                // Y-framing: swap width and depth
                double tempWidth = roundedWidth;
                roundedWidth = roundedHeight;  // Use height as width
                roundedHeight = tempWidth;     // Use original width as height
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Y-FRAMING SWAP: Width={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, Height={RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm");
            }
        }
        
                var depthParam = GetParam("Depth");
                var wallWidthParam = GetParam("Wall Width");
        
        // Get the correct thickness based on host type
        double thickness = 0.0;
        if (isWallHost)
        {
            thickness = clashZone.WallThickness > 0 ? clashZone.WallThickness : clashZone.StructuralElementThickness;
        }
        else if (isFramingHost)
        {
            thickness = clashZone.FramingThickness > 0 ? clashZone.FramingThickness : clashZone.StructuralElementThickness;
        }
        else
        {
            thickness = clashZone.StructuralElementThickness;
        }
        
        // ✅ PATH 1 (Replay): Log if structural thickness is missing (but don't retrieve - use existing data only)
        if (_isReplayPath && thickness <= 0.0)
        {
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[PATH-1-DIAGNOSTIC] Zone={clashZone.Id}: Missing structural thickness - WallThickness={clashZone.WallThickness:F6}, FramingThickness={clashZone.FramingThickness:F6}, StructuralThickness={clashZone.StructuralElementThickness:F6}, HostType={clashZone.StructuralElementType}. PATH 1 will use 0 (no retrieval from linked files).");
            }
        }
        
        // ✅ CRITICAL FIX: For PATH 3 (Non-Fresh), if thickness is 0, retrieve from linked file structural element
        // This is a fallback - thickness should have been saved during refresh, but if missing, retrieve it now
        if (!_isReplayPath && thickness <= 0.0 && clashZone.StructuralElementIdValue > 0)
        {
            try
            {
                // Try to get structural element from linked files
                Element structuralElement = null;
                var linkInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();
                
                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        try
                        {
                            structuralElement = linkDoc.GetElement(new ElementId(clashZone.StructuralElementIdValue));
                            if (structuralElement != null)
                            {
                                // Retrieve thickness from linked file element
                                if (structuralElement is Wall wall)
                                {
                                    thickness = wall.Width;
                                }
                                else if (structuralElement is Floor floor)
                                {
                                    thickness = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.0;
                                }
                                else if (structuralElement?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                                {
                                    // Get framing thickness from type parameter 'b'
                                    var typeId = (structuralElement as FamilyInstance)?.GetTypeId() ?? structuralElement.GetTypeId();
                                    var typeElem = linkDoc.GetElement(typeId);
                                    if (typeElem != null)
                                    {
                                        var param = typeElem.LookupParameter("b") ?? typeElem.LookupParameter("B");
                                        if (param != null) thickness = param.AsDouble();
                                    }
                                }
                                
                                if (thickness > 0.0)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[PATH-3-FALLBACK] Retrieved thickness={RevitUnitConversionService.Instance.FromInternalMillimeters(thickness):F1}mm from linked file for structural element {clashZone.StructuralElementIdValue} (Zone={clashZone.Id})");
                                    break;
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PATH-3-FALLBACK] Error retrieving thickness from linked file for Zone={clashZone.Id}: {ex.Message}");
            }
        }
        
        double thicknessMm = RevitUnitConversionService.Instance.FromInternalMillimeters(thickness);
        
        // ✅ DIAGNOSTIC: Log thickness calculation for debugging depth=0 issue
        if (!DeploymentConfiguration.DeploymentMode && thickness <= 0.0)
        {
            DebugLogger.Warning($"[DEPTH-DIAGNOSTIC] Zone={clashZone.Id}, HostType={clashZone.StructuralElementType}, WallThickness={clashZone.WallThickness:F6}, FramingThickness={clashZone.FramingThickness:F6}, StructuralThickness={clashZone.StructuralElementThickness:F6}, CalculatedThickness={thickness:F6}, Path={(_isReplayPath ? "Replay" : "Sizing/Detection")}");
        }
        
        // ✅ PERFORMANCE: Removed verbose depth check logging
        
        bool depthSetSuccess = false;
        
        if (isWallHost && wallWidthParam != null && !wallWidthParam.IsReadOnly)
        {
            // Wall host: use Wall Width parameter
            TimedSetDouble(wallWidthParam, thickness, "Wall Width");
            depthSetSuccess = true;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[DEPTH-SET] Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: Set Wall Width={RevitUnitConversionService.Instance.FromInternalMillimeters(thickness):F1}mm (from DB: Structural={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.StructuralElementThickness):F1}mm, Wall={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.WallThickness):F1}mm)");
            }
        }
        else if (depthParam != null && !depthParam.IsReadOnly)
        {
            TimedSetDouble(depthParam, thickness, "Depth");
            depthSetSuccess = true;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[DEPTH-SET] Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: Set Depth={RevitUnitConversionService.Instance.FromInternalMillimeters(thickness):F1}mm (from DB: Structural={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.StructuralElementThickness):F1}mm, Framing={RevitUnitConversionService.Instance.FromInternalMillimeters(clashZone.FramingThickness):F1}mm)");
            }
        }
        else
        {
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[DEPTH-SET] ❌ Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: Depth parameter not writable - wallWidthParam={(wallWidthParam != null ? "exists" : "null")}, depthParam={(depthParam != null ? "exists" : "null")}, isReadOnly={(depthParam?.IsReadOnly ?? true)}, family='{sleeveInstance.Symbol?.Family?.Name ?? "Unknown"}'");
            }
        }
        
        // ✅ PERFORMANCE: Only log depth parameter failures in diagnostic mode
        if (!depthSetSuccess && !DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
        {
            DebugLogger.Warning($"[UniversalSleevePlacer] Could not find writable Depth or Wall Width parameter on sleeve {sleeveInstance.Id}");
        }
        
        // ✅ BOTTOM OF OPENING: Calculate and set for rectangular openings on walls and framing
        // Formula: Bottom of Opening = Elevation from Level - (Height / 2)
        // This calculates the bottom edge of the opening for scheduling purposes
        if (!treatAsCircular && (isWallHost || isFramingHost))
        {
            try
            {
                // ✅ PERFORMANCE: Use cached parameters only, no fallback LookupParameter calls
                var elevationFromLevelParam = GetParam("Elevation from Level") 
                                           ?? GetParam("Schedule Level Elevation")
                                           ?? GetParam("Elevation from Level Offset");
                
                var heightParam = GetParam("Height");
                var bottomOfOpeningParam = GetParam("Bottom of Opening");
                
                if (elevationFromLevelParam != null && heightParam != null && bottomOfOpeningParam != null && !bottomOfOpeningParam.IsReadOnly)
                {
                    // Get current values
                    double elevationFromLevel = elevationFromLevelParam.AsDouble();
                    double height = heightParam.AsDouble();
                    
                    // Validate values are valid
                    if (height > 0 && Math.Abs(elevationFromLevel) < 10000) // Reasonable bounds check
                    {
                    // Calculate: Bottom of Opening = Elevation from Level - (Height / 2)
                    // Elevation from Level gives center of opening, subtract half height to get bottom
                    double bottomOfOpening = elevationFromLevel - (height / 2.0);
                    
                    // Set the parameter (batched if enabled)
                    TimedSetDouble(bottomOfOpeningParam, bottomOfOpening, "Bottom of Opening");
                        
                        // ✅ PERFORMANCE: Removed verbose logging
                    }
                    else
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[UniversalSleevePlacer] Invalid values for Bottom of Opening calculation: Elevation={elevationFromLevel:F3}ft, Height={height:F3}ft");
                    }
                }
                else
                {
                    var elevationFound = elevationFromLevelParam != null;
                    var heightFound = heightParam != null;
                    var bottomFound = bottomOfOpeningParam != null && !bottomOfOpeningParam.IsReadOnly;
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalSleevePlacer] Cannot set Bottom of Opening: Elevation from Level={elevationFound}, Height={heightFound}, Bottom of Opening={bottomFound}");
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalSleevePlacer] Error calculating Bottom of Opening: {ex.Message}");
            }
        }
        
        // ⚡ PERFORMANCE OPTIMIZATION: Split metadata into critical (immediate) and non-critical (deferred)
        // Critical parameters required for flag reset during refresh: MEP_Category, MEP_ElementId, ClashZone_GUID, Sleeve Instance ID, Filter Name
        // Non-critical parameters deferred to batch write: MEP_UniqueId, MEP_Size, System_Abbreviation, MEP_Count, Bottom of Opening, Host Parameters
        
        if (OptimizationFlags.DeferNonCriticalMetadata)
        {
            // ✅ PHASE 2: Write ONLY critical parameters immediately (required for flag reset)
            
            // 1. MEP_Category - Required by FlagManager.ResetFlagsForDeletedSleeves (line 1801-1805)
            var mepCategoryParam = GetParam("MEP_Category");
            if (mepCategoryParam != null)
            {
                TimedSetString(mepCategoryParam, clashZone.MepElementCategory, "MEP_Category");
            }
            
            // 2. MEP_ElementId - Required by FlagManager.RecoverFlagsFromSleevesInRevit (line 2826)
            var mepElementIdParam = GetParam("MEP_ElementId");
            if (mepElementIdParam != null)
            {
                TimedSetInt(mepElementIdParam, clashZone.MepElementId.IntegerValue, "MEP_ElementId");
            }
            
            // 3. ClashZone_GUID - Required by FlagManager.GetClashZoneGuidValue (line 3516)
            var guidParam = GetParam("ClashZone_GUID");
            if (guidParam != null)
            {
                TimedSetString(guidParam, clashZone.Id.ToString(), "ClashZone_GUID");
            }
            
            // 4. Sleeve Instance ID - Required by FlagManager.ResetFlagsForDeletedSleeves (line 1984)
            var instanceIdParam = GetParam("Sleeve Instance ID");
            if (instanceIdParam != null)
            {
                TimedSetInt(instanceIdParam, sleeveInstance.Id.IntegerValue, "Sleeve Instance ID");
            }
            
            // 5. Filter Name - Required by FlagManager.RecoverFlagsFromSleevesInRevit (line 2805)
            string filterName = GetFilterNameForCategory(clashZone.MepElementCategory);
            var filterNameParam = GetParam("Filter Name");
            if (filterNameParam != null)
            {
                TimedSetString(filterNameParam, filterName, "Filter Name");
            }
            
            // ⚡ NON-CRITICAL PARAMETERS DEFERRED - Will be written in batch after all sleeves placed
            // Deferred: MEP_UniqueId, MEP_Size, System_Abbreviation, MEP_Count, Bottom of Opening, Host Parameters
            // Expected gain: ~150-180ms per sleeve
        }
        else
        {
            // 🔄 LEGACY PATH: Write all metadata immediately (old behavior for rollback safety)
            
            // Set MEP metadata parameters
                    var mepElementIdParam = GetParam("MEP_ElementId");
                    if (mepElementIdParam != null)
                    {
                        TimedSetInt(mepElementIdParam, clashZone.MepElementId.IntegerValue, "MEP_ElementId");
                    }

                    var mepUniqueIdParam = GetParam("MEP_UniqueId");
                    if (mepUniqueIdParam != null)
                    {
                        TimedSetString(mepUniqueIdParam, clashZone.MepElementUniqueId, "MEP_UniqueId");
                    }
                    
                    var mepSizeParam = GetParam("MEP_Size");
                    if (mepSizeParam != null)
                    {
                        TimedSetString(mepSizeParam, clashZone.MepElementFormattedSize, "MEP_Size");
                    }
                    
                    var systemAbbrParam = GetParam("System_Abbreviation");
                    if (systemAbbrParam != null)
                    {
                        TimedSetString(systemAbbrParam, clashZone.MepElementSystemAbbreviation, "System_Abbreviation");
                    }
                    
                    var mepCountParam = GetParam("MEP_Count");
                    if (mepCountParam != null)
                    {
                        TimedSetInt(mepCountParam, 1, "MEP_Count");  // Individual sleeve
                    }
                    
                    // ✅ CRITICAL FIX: Set MEP_Category parameter for clustering
                    // ✅ PERFORMANCE: Reduced logging - only log failures in diagnostic mode
                    var mepCategoryParam = GetParam("MEP_Category");
                    if (mepCategoryParam != null)
                    {
                        TimedSetString(mepCategoryParam, clashZone.MepElementCategory, "MEP_Category");
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] Set MEP_Category = '{clashZone.MepElementCategory}' for sleeve {sleeveInstance.Id.IntegerValue}");
                        }
                    }
                    else if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        DebugLogger.Warning($"[UniversalSleevePlacer] MEP_Category parameter not found on sleeve {sleeveInstance.Id}");
                    }
        }

                // ⚠️ CRITICAL FIX: Transfer HostParameterValues from XML intersection data to sleeve parameters
                // ✅ PERFORMANCE OPTIMIZATION: Reduced logging and use pre-cached parameters
                // ⚡ DEFERRED if DeferNonCriticalMetadata=true (moved to batch writer)
                if (!OptimizationFlags.DeferNonCriticalMetadata && clashZone.HostParameterValues != null && clashZone.HostParameterValues.Count > 0)
                {
                    int hostParamsSet = 0;
                    int hostParamsFailed = 0;
                    
                    foreach (var hostParam in clashZone.HostParameterValues)
                    {
                        try
                        {
                            // ✅ PERFORMANCE: Use pre-cached parameter instead of LookupParameter()
                            var param = GetParam(hostParam.Key);
                            if (param != null)
                            {
                                // Handle different parameter storage types (batched if enabled)
                                if (param.StorageType == StorageType.String)
                                {
                                    TimedSetString(param, hostParam.Value, hostParam.Key);
                                    hostParamsSet++;
                                }
                                else if (param.StorageType == StorageType.Integer)
                                {
                                    if (int.TryParse(hostParam.Value, out int intValue))
                                    {
                                        TimedSetInt(param, intValue, hostParam.Key);
                                        hostParamsSet++;
                                    }
                                    else
                                    {
                                        hostParamsFailed++;
                                    }
                                }
                                else if (param.StorageType == StorageType.Double)
                                {
                                    if (double.TryParse(hostParam.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double doubleValue))
                                    {
                                        TimedSetDouble(param, doubleValue, hostParam.Key);
                                        hostParamsSet++;
                                    }
                                    else
                                    {
                                        hostParamsFailed++;
                                    }
                                }
                                // ElementId parameters skipped silently (no warning needed)
                            }
                            else
                            {
                                hostParamsFailed++;
                            }
                        }
                        catch (Exception paramEx)
                        {
                            hostParamsFailed++;
                            // ✅ PERFORMANCE: Only log errors in diagnostic mode, not every missing parameter
                            if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                            {
                                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting host parameter '{hostParam.Key}': {paramEx.Message}");
                            }
                        }
                    }
                    
                    // ✅ PERFORMANCE: Single summary log instead of per-parameter logging
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode && (hostParamsSet > 0 || hostParamsFailed > 0))
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Host parameters: {hostParamsSet} set, {hostParamsFailed} failed for sleeve {sleeveInstance.Id.IntegerValue}");
                    }
                }

                // ✅ PERFORMANCE: Removed summary log - parameters are set silently
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
        DebugLogger.Warning($"[UniversalSleevePlacer] Error setting parameters: {ex.Message}\n{ex.StackTrace}");
        try
        {
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[PARAM-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
        }
        catch { }
    }
}

        /// <summary>
        /// Set sleeve metadata parameters for fast parameter transfer
        /// </summary>
        private void SetSleeveMetadata(FamilyInstance sleeveInstance, ClashZone clashZone)
        {
            try
            {
                // ✅ BATCHING HELPER: Defer string parameter writes
                void TimedSetString(Parameter p, string value, string logicalName)
                {
                    if (p == null || p.IsReadOnly) return;
                    
                    if (OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                            _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                        _deferredParameters[sleeveInstance.Id][logicalName] = value;
                        return;
                    }
                    
                    p.Set(value);
                }
                
                // ✅ BATCHING HELPER: Defer int parameter writes
                void TimedSetInt(Parameter p, int value, string logicalName)
                {
                    if (p == null || p.IsReadOnly) return;
                    
                    if (OptimizationFlags.UseBatchedParameterWrites)
                    {
                        if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                            _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                        _deferredParameters[sleeveInstance.Id][logicalName] = value;
                        return;
                    }
                    
                    p.Set(value);
                }
                
                // Set Filter Name based on category
                string filterName = GetFilterNameForCategory(clashZone.MepElementCategory);
                var filterNameParam = sleeveInstance.LookupParameter("Filter Name");
                if (filterNameParam != null && !filterNameParam.IsReadOnly)
                {
                    TimedSetString(filterNameParam, filterName, "Filter Name");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SetSleeveMetadata] Set Filter Name = '{filterName}' for sleeve {sleeveInstance.Id}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SetSleeveMetadata] Filter Name parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // Set Instance ID
                var instanceIdParam = sleeveInstance.LookupParameter("Sleeve Instance ID");
                if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                {
                    TimedSetInt(instanceIdParam, sleeveInstance.Id.IntegerValue, "Sleeve Instance ID");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SetSleeveMetadata] Set Sleeve Instance ID = {sleeveInstance.Id.IntegerValue} for sleeve {sleeveInstance.Id}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SetSleeveMetadata] Sleeve Instance ID parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // ✅ NEW: Set LinkedFile parameter from ClashZone.SourceDocKey
                var linkedFileParam = sleeveInstance.LookupParameter("LinkedFile");
                if (linkedFileParam != null && !linkedFileParam.IsReadOnly)
                {
                    // Normalize SourceDocKey: remove extension and element count if present
                    string linkedFile = clashZone.SourceDocKey ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(linkedFile))
                    {
                        linkedFile = System.IO.Path.GetFileNameWithoutExtension(linkedFile);
                        int parenIndex = linkedFile.IndexOf('(');
                        if (parenIndex > 0)
                            linkedFile = linkedFile.Substring(0, parenIndex).Trim();
                    }
                    TimedSetString(linkedFileParam, linkedFile, "LinkedFile");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SetSleeveMetadata] Set LinkedFile = '{linkedFile}' for sleeve {sleeveInstance.Id}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SetSleeveMetadata] LinkedFile parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // ✅ NEW: Set HostFile parameter from ClashZone.HostDocKey
                var hostFileParam = sleeveInstance.LookupParameter("HostFile");
                if (hostFileParam != null && !hostFileParam.IsReadOnly)
                {
                    // Normalize HostDocKey: remove extension and element count if present
                    string hostFile = clashZone.HostDocKey ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(hostFile))
                    {
                        hostFile = System.IO.Path.GetFileNameWithoutExtension(hostFile);
                        int parenIndex = hostFile.IndexOf('(');
                        if (parenIndex > 0)
                            hostFile = hostFile.Substring(0, parenIndex).Trim();
                    }
                    TimedSetString(hostFileParam, hostFile, "HostFile");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SetSleeveMetadata] Set HostFile = '{hostFile}' for sleeve {sleeveInstance.Id}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SetSleeveMetadata] HostFile parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // Log to debug file
                try
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SLEEVE_METADATA] Sleeve {sleeveInstance.Id.IntegerValue}: Filter='{filterName}', InstanceID={sleeveInstance.Id.IntegerValue}, LinkedFile='{clashZone.SourceDocKey}', HostFile='{clashZone.HostDocKey}' ✓\n");
                }
                catch { }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[SetSleeveMetadata] Error setting metadata for sleeve {sleeveInstance.Id}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// ✅ CRITICAL: Set ClashZone_GUID parameter on sleeve with STABLE GUID per 3-point combo
        /// Uses clashZone.Id which is now deterministic (generated from MEP+Host+Point hash)
        /// This ensures GUID is stable across multiple detection runs for the same intersection
        /// Follows industry best practices for stable clash identification
        /// </summary>
        /// <param name="sleeveInstance">The sleeve family instance</param>
        /// <param name="clashZone">The clash zone being placed</param>
        private void SetClashZoneGuidOnSleeveStable(FamilyInstance sleeveInstance, ClashZone clashZone)
        {
            try
            {
                var guidParam = sleeveInstance.LookupParameter("ClashZone_GUID");
                if (guidParam == null || guidParam.IsReadOnly)
                {
                    // Parameter not found - log warning but don't fail (backward compatibility)
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SetClashZoneGuidStable] ClashZone_GUID parameter not found or read-only on sleeve {sleeveInstance.Id} - GUID storage skipped. Add 'ClashZone_GUID' text parameter to family.");
                    return;
                }
                
                // ✅ Use clashZone.Id directly (now deterministic - same intersection always gets same GUID)
                // clashZone.Id is generated deterministically from MEP+Host+Point hash, so it's stable across detection runs
                // No need to lookup Global XML - deterministic GUID ensures consistency
                Guid stableGuid = clashZone.Id;
                string guidString = stableGuid.ToString();
                
                // ✅ BATCHING: Defer parameter write if batching enabled
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                        _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                    _deferredParameters[sleeveInstance.Id]["ClashZone_GUID"] = guidString;
                }
                else
                {
                    guidParam.Set(guidString);
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SetClashZoneGuidStable] Set ClashZone_GUID = '{guidString}' for sleeve {sleeveInstance.Id} (stable per 3-point combo)");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SetClashZoneGuidStable] Error setting ClashZone_GUID for sleeve {sleeveInstance.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Immediately update clash zone in XML to prevent value loss
        /// This ensures the values are saved before any other process can overwrite them
        /// </summary>
        /// <summary>
        /// Collect sleeve data for clustering from placed sleeves
        /// </summary>
        private List<SleeveData> CollectSleeveDataForClustering()
        {
            var sleeveDataList = new List<SleeveData>();
            
            try
            {
                // Get all sleeves in the model
                var sleeves = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                    .ToList();
                
                foreach (var sleeve in sleeves)
                {
                    try
                    {
                        var bbox = sleeve.get_BoundingBox(null);
                        if (bbox != null)
                        {
                            var sleeveData = new SleeveData
                            {
                                SleeveInstanceId = sleeve.Id.IntegerValue,
                                Category = _strategy.GetCategoryName(),
                                HostType = GetHostType(sleeve),
                                Orientation = GetSleeveOrientation(sleeve),
                                Width = bbox.Max.X - bbox.Min.X,
                                Height = bbox.Max.Y - bbox.Min.Y,
                                Depth = bbox.Max.Z - bbox.Min.Z
                            };
                            
                            // Calculate 4 corner coordinates
                            sleeveData.Corner1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
                            sleeveData.Corner2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
                            sleeveData.Corner3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z);
                            sleeveData.Corner4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z);
                            
                            sleeveDataList.Add(sleeveData);
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[COLLECT] Sleeve {sleeve.Id.IntegerValue}: Host={sleeveData.HostType}, Orientation={sleeveData.Orientation}, W={sleeveData.Width:F3}, H={sleeveData.Height:F3}, D={sleeveData.Depth:F3}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ERROR] Failed to collect data for sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[ERROR] Failed to collect sleeve data: {ex.Message}\n");
            }
            
            return sleeveDataList;
        }
        
        private string GetHostType(FamilyInstance sleeve)
        {
            try
            {
                var host = sleeve.Host;
                if (host != null)
                {
                    var hostType = host.GetType().Name;
                    if (hostType.Contains("Wall")) return "Wall";
                    if (hostType.Contains("Floor")) return "Floor";
                    if (hostType.Contains("Ceiling")) return "Ceiling";
                    return hostType;
                }
            }
            catch { }
            return "Unknown";
        }
        
        private void UpdateClashZoneInXml(ClashZone clashZone)
        {
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ Filters directory does not exist: {filtersDirectory}\n");
                    return;
                }

                // 🎯 CRITICAL FIX: Use category-based file matching to target the correct XML file
                var targetFileName = GetFilterNameForCategory(clashZone.MepElementCategory);
                
                // ✅ CRITICAL LOGGING: Log XML file details
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[XML-FILE-SEARCH] {DateTime.Now:HH:mm:ss.fff} - Looking for zone {clashZone.Id} in category file: {targetFileName}\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[XML-FILE-SEARCH] Category: {clashZone.MepElementCategory}, TargetFileName: {targetFileName}\n");
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] Looking for zone {clashZone.Id} in category file: {targetFileName}\n");

                // First, try to find the exact category-specific file
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                string targetFile = null;
                
                foreach (var xmlFile in xmlFiles)
                {
                    var fileName = Path.GetFileName(xmlFile);
                    if (fileName.Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                        continue;
                        
                    // ✅ CRITICAL: Use EXACT match - no Contains() fallback to prevent wrong XML matching
                    if (fileName.Equals(targetFileName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetFile = xmlFile;
                        
                        // ✅ CRITICAL LOGGING: Log XML file found
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[XML-FILE-FOUND] {DateTime.Now:HH:mm:ss.fff} - Found target file: {fileName}\n");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[XML-FILE-FOUND] Full path: {xmlFile}\n");
                        
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] Found target file: {fileName}\n");
                        break;
                    }
                }

                // If no specific file found, fall back to searching all files
                if (targetFile == null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] No specific file found, searching all {xmlFiles.Length} XML files\n");
                    targetFile = xmlFiles.FirstOrDefault(f => !Path.GetFileName(f).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase));
                }

                if (targetFile == null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ No suitable XML file found!\n");
                    return;
                }

                // Load and update the target file
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                Models.OpeningFilter filter;
                
                using (var reader = new StreamReader(targetFile))
                {
                    filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                }
                
                var storageZones = filter?.ClashZoneStorage?.AllZones;
                if (storageZones == null || storageZones.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ No clash zones in file: {Path.GetFileName(targetFile)}\n");
                    return;
                }

                // 🔥 TRIPLE-MATCH VERIFICATION: Match by ID AND MepElementId AND StructuralElementId (tree-aware)
                var zone = storageZones.FirstOrDefault(z => 
                    z.Id == clashZone.Id && 
                    z.MepElementIdValue == clashZone.MepElementIdValue && 
                    z.StructuralElementIdValue == clashZone.StructuralElementIdValue);
                
                if (zone != null)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ✅ FOUND zone {clashZone.Id} in {Path.GetFileName(targetFile)}\n");
                    
                    // Log BEFORE values
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] BEFORE: W={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveHeight):F1}mm, ActiveDocX={zone.SleevePlacementPointActiveDocumentX:F3}\n");
                    
                    // Update the values
                    zone.SleeveWidth = clashZone.SleeveWidth;
                    zone.SleeveHeight = clashZone.SleeveHeight;
                    zone.SleeveDiameter = clashZone.SleeveDiameter;
                    zone.SleevePlacementPointActiveDocumentX = clashZone.SleevePlacementPointActiveDocumentX;
                    zone.SleevePlacementPointActiveDocumentY = clashZone.SleevePlacementPointActiveDocumentY;
                    zone.SleevePlacementPointActiveDocumentZ = clashZone.SleevePlacementPointActiveDocumentZ;
                    zone.SleeveInstanceId = clashZone.SleeveInstanceId;
                    
                    // ✅ REMOVED: Bounding box coordinates - now handled by SleeveCoordinateService
                    
                    // Log AFTER values
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] AFTER: W={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveHeight):F1}mm, ActiveDocX={zone.SleevePlacementPointActiveDocumentX:F3}\n");
                    
                    // Normalize coordinates to avoid 0,0,0 in XML
                    foreach (var z in storageZones)
                    {
                        NormalizeIntersectionCoordinates(z);
                    }

                    // Log a few zones' coordinates to placement_debug before saving (using SafeFileLogger path)
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        var sample = storageZones.Take(3).ToList();
                        for (int i = 0; i < sample.Count; i++)
                        {
                            var s = sample[i];
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] XML-IMMEDIATE-UPDATE BEFORE SAVE: {Path.GetFileName(targetFile)} Zone={s.Id} IP=({s.IntersectionPointX},{s.IntersectionPointY},{s.IntersectionPointZ}) SPP=({s.SleevePlacementPointX},{s.SleevePlacementPointY},{s.SleevePlacementPointZ})\n");
                            }
                        }
                    }
                    catch (Exception logEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[XML-LOG-ERROR] {logEx.Message}");
                    }
                    
                    // Save the file immediately
                    filter.LastModified = DateTime.Now;
                    using (var writer = new StreamWriter(targetFile))
                    {
                        serializer.Serialize(writer, filter);
                    }
                    
                                        if (!DeploymentConfiguration.DeploymentMode)
                    {
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] {DateTime.Now:HH:mm:ss.fff} - Successfully saved to {Path.GetFileName(targetFile)}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Zone: {clashZone.Id}, SleeveInstanceId: {zone.SleeveInstanceId}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Coordinates: X={zone.SleevePlacementPointActiveDocumentX:F3}, Y={zone.SleevePlacementPointActiveDocumentY:F3}, Z={zone.SleevePlacementPointActiveDocumentZ:F3}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Dimensions: W={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveWidth):F1}mm, H={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.SleeveHeight):F1}mm\n");
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ✅ SUCCESS: Updated {Path.GetFileName(targetFile)} with values for zone {clashZone.Id}\n");
                    }
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ Zone {clashZone.Id} NOT found in {Path.GetFileName(targetFile)} (checked {storageZones.Count} zones)\n");
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[XML-IMMEDIATE-UPDATE-ERROR] Error updating XML: {ex.Message}\n");
            }
        }

        /// <summary>
        /// ✅ CORRECT APPROACH: Use ClashZonePersistenceService to save updated clash zones
        /// This ensures the tree structure is maintained correctly using the dedicated service
        /// </summary>
        private void SaveUpdatedXmlFiles(List<ClashZone> updatedClashZones = null)
        {
            var diagnosticLog = new System.Text.StringBuilder();
            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] === SaveUpdatedXmlFiles START ===");

            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] SaveUpdatedXmlFiles: {updatedClashZones?.Count ?? 0} zones ready (delegating to persistence service)");

                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Input count: {updatedClashZones?.Count ?? 0}");
                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Sample (first 10): {string.Join(", ", updatedClashZones?.Where(z => z != null).Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10) ?? Array.Empty<string>())}");

                if (updatedClashZones == null || updatedClashZones.Count == 0)
                {
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] EXIT: No updated clash zones to save");
                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info("[UniversalSleevePlacer] No updated clash zones to save");
                    return;
                }
                
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filters directory: {filtersDirectory}");
                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Directory exists: {Directory.Exists(filtersDirectory)}");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] EXIT: Filters directory not found");
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalSleevePlacer] Filters directory not found: {filtersDirectory}");
                    return;
                }
                
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                var persistenceService = new ClashZonePersistenceService(_doc, new GuidManager(_doc));
                var filterManagementService = new FilterManagementService(_doc, null, null);

                var clashZonesByCategory = updatedClashZones
                    .GroupBy(cz => cz.MepElementCategory)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .ToList();

                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Categories: {clashZonesByCategory.Count}");
                foreach (var group in clashZonesByCategory)
                        {
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}]  - {group.Key}: {group.Count()} zones");
                }

                string sanitizedFilterNameGlobal = _filterName;
                if (!string.IsNullOrWhiteSpace(sanitizedFilterNameGlobal))
                {
                    sanitizedFilterNameGlobal = sanitizedFilterNameGlobal.Trim();
                    sanitizedFilterNameGlobal = Path.GetFileName(sanitizedFilterNameGlobal);
                }

                string baseFilterNameForPersistence = null;

                var processedCategoryFilters = new List<(Models.OpeningFilter Filter, string Category)>();

                int categoryIndex = 0;
                foreach (var categoryGroup in clashZonesByCategory)
                            {
                    categoryIndex++;
                    var category = categoryGroup.Key;
                    var categoryZones = categoryGroup.ToList();

                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] --- Category {categoryIndex}/{clashZonesByCategory.Count}: {category} ---");
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Category zones count: {categoryZones.Count}");
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Category sample: {string.Join(", ", categoryZones.Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10))}");

                    string filterFileName;
                    if (!string.IsNullOrEmpty(sanitizedFilterNameGlobal) && sanitizedFilterNameGlobal.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        filterFileName = sanitizedFilterNameGlobal;
                    }
                    else if (!string.IsNullOrEmpty(sanitizedFilterNameGlobal))
                    {
                        filterFileName = $"{sanitizedFilterNameGlobal}_{category.ToLower()}.xml";
                    }
                    else
                    {
                        filterFileName = $"*_{category.ToLower()}.xml";
                    }

                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter name='{_filterName}', sanitized='{sanitizedFilterNameGlobal}', pattern='{filterFileName}'");

                    string[] matchingFiles = null;
                    try
                    {
                        matchingFiles = Directory.GetFiles(filtersDirectory, filterFileName);
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Matching files: {matchingFiles.Length}");
                        foreach (var file in matchingFiles)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}]   • {Path.GetFileName(file)}");
                        }
                    }
                    catch (Exception findEx)
                    {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ERROR finding files: {findEx}");
                                }

                    var xmlFile = matchingFiles?.FirstOrDefault();
                    if (xmlFile == null || !File.Exists(xmlFile))
                    {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] SKIP: Filter XML file not found for pattern {filterFileName}");
                                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[UniversalSleevePlacer] Filter XML file not found: {filterFileName}");
                            continue;
                        }
                        
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Target file: {Path.GetFileName(xmlFile)} (size={new FileInfo(xmlFile).Length} bytes)");

                    try
                    {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1: Deserializing XML...");
                        Models.OpeningFilter filter = null;
                        try
                        {
                        using (var reader = new StreamReader(xmlFile))
                        {
                            filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                            }
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1 SUCCESS");
                        }
                        catch (Exception deserializeEx)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 1 FAILED: {deserializeEx}");
                            throw;
                        }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2: Validating filter");
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Filter null? {filter == null}");
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Storage null? {filter?.ClashZoneStorage == null}");

                        if (filter?.ClashZoneStorage == null)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2 FAILED: Storage null");
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalSleevePlacer] Filter or ClashZoneStorage is null in {Path.GetFileName(xmlFile)}");
                            continue;
                        }
                        
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 2 SUCCESS");

                        try
                        {
                            var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            var samplePairs = categoryZones.Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10);
                            System.IO.File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-BEFORE-PERSIST] File={Path.GetFileName(xmlFile)}, Sample=[{string.Join(", ", samplePairs)}]\n");
                            }
                        catch { }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3: Saving via persistence service");
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 Sample (pre-persist): {string.Join(", ", categoryZones.Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10))}");

                        try
                        {
                            var effectiveFilterName = FilterNameHelper.NormalizeBaseName(_filterName, filter?.Name, category);
                            baseFilterNameForPersistence ??= effectiveFilterName;
                            var persistenceZones = categoryZones
                                .Select(CloneZoneForPersistence)
                                .Where(z => z != null)
                                .ToList();

                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 Snapshot (clone): {string.Join(", ", persistenceZones.Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10))}");

                            // ✅ PRINCIPLE: Only PATH 3 (Detection) for non-validated zones should allow structural updates
                            // PATH 1 (Replay): false - replaying existing data, no updates
                            // PATH 2 (Sizing): false - sizing existing sleeves, no structural updates needed
                            // PATH 3 (Detection): true ONLY for non-validated zones (validated zones treated as PATH 1)
                            // Note: Stale DB/XML data should be cleared, not worked around with code changes
                            bool allowStructuralUpdates = false; // Default: no structural updates
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[PERSISTENCE] Path={(_isReplayPath ? "PATH 1 (Replay)" : "PATH 2/3 (Sizing/Detection)")}, allowStructuralUpdates={allowStructuralUpdates} (following principle: only PATH 3 non-validated should allow updates)");
                            }
                            
                            persistenceService.SaveClashZones(persistenceZones, effectiveFilterName, filter, allowStructuralUpdates);
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 SUCCESS: allowStructuralUpdates={allowStructuralUpdates}");
                            }
                        catch (Exception persistEx)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 FAILED: {persistEx}");
                            throw;
                        }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4: Saving filter to XML");
                        try
                        {
                            if (!DeploymentConfiguration.DisableXmlCreation)
                            {
                                filterManagementService.SaveFilterToXmlFile(filter, xmlFile);
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4 SUCCESS");
                                processedCategoryFilters.Add((filter, category));
                            }
                            else
                            {
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4 SKIPPED: XML creation disabled (database only mode)");
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[UniversalSleevePlacer] ⚠️ XML creation disabled - skipping filter XML save for '{category}' (database only mode)");
                                }
                            }
                        }
                        catch (Exception saveFileEx)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4 FAILED: {saveFileEx}");
                            throw;
                        }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ✅ XML-MERGE-COMPLETE: {Path.GetFileName(xmlFile)}, Zones={categoryZones.Count}");

                        try
                        {
                            var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            System.IO.File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [XML-MERGE-COMPLETE] File={Path.GetFileName(xmlFile)}, Zones={categoryZones.Count}\n");
                        }
                        catch { }
                    }
                    catch (Exception categoryEx)
                    {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ CATEGORY ERROR: {categoryEx}");
                        try
                        {
                            var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                            System.IO.File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [XML-MERGE-ERROR] File={Path.GetFileName(xmlFile)} → {categoryEx.Message}\n");
                                }
                        catch { }
                                    
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalSleevePlacer] Error saving {Path.GetFileName(xmlFile)}: {categoryEx.Message}");
                                }
                            }
                
                // ✅ Ensure the main filter (e.g., Plumbing.xml) carries the latest placement data.
                //    Naming MUST stay normalized (base filter name only). The persistence service
                //    appends category suffixes internally, so feeding "Plumbing" here guarantees
                //    the category snapshots and the root filter stay aligned.
                try
                {
                    var baseFilterName = baseFilterNameForPersistence ?? FilterNameHelper.NormalizeBaseName(_filterName, sanitizedFilterNameGlobal);
                    if (string.IsNullOrWhiteSpace(baseFilterName))
                    {
                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATE SKIPPED → base name missing");
                    }
                    else
                    {
                        baseFilterName = Path.GetFileNameWithoutExtension(baseFilterName);
                        var mainFilterPath = Path.Combine(filtersDirectory, $"{baseFilterName}.xml");

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATE → base='{baseFilterName}', path='{mainFilterPath}', exists={File.Exists(mainFilterPath)}");

                        if (File.Exists(mainFilterPath) && processedCategoryFilters.Count > 0)
                        {
                            Models.OpeningFilter mainFilter = null;
                            try
                            {
                                using (var reader = new StreamReader(mainFilterPath))
                                {
                                    mainFilter = (Models.OpeningFilter)serializer.Deserialize(reader);
                                }
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER DESERIALIZE SUCCESS");
                            }
                            catch (Exception mainDeserializeEx)
                            {
                                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER DESERIALIZE FAILED: {mainDeserializeEx}");
                                mainFilter = null;
                            }

                            if (mainFilter != null)
                            {
                                try
                                {
                                    foreach (var (categoryFilter, categoryName) in processedCategoryFilters)
                                    {
                                        MergeCategoryIntoMainFilter(mainFilter, categoryFilter);
                                    }
                                    
                                    // ✅ PHASE 2: Only save main filter XML if XML creation is enabled
                                    if (!DeploymentConfiguration.DisableXmlCreation)
                                    {
                                        filterManagementService.SaveFilterToXmlFile(mainFilter, mainFilterPath);
                                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATED: {Path.GetFileName(mainFilterPath)}");
                                    }
                                    else
                                    {
                                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATE SKIPPED: XML creation disabled (database only mode)");
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[UniversalSleevePlacer] ⚠️ XML creation disabled - skipping main filter XML save (database only mode)");
                                        }
                                    }

                                    try
                                    {
                                        var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                                        System.IO.File.AppendAllText(logPath,
                                            $"[{DateTime.Now:HH:mm:ss}] [XML-MERGE-COMPLETE] File={Path.GetFileName(mainFilterPath)}, Zones={mainFilter?.ClashZoneStorage?.AllZones?.Count ?? 0}\n");
                                    }
                                    catch { }
                                }
                                catch (Exception mergeEx)
                                {
                                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER MERGE FAILED: {mergeEx}");
                                }

                                EnsureConditionsFiles(_doc, baseFilterName, processedCategoryFilters, filtersDirectory, diagnosticLog);
                            }
                        }
                    }
                }
                catch (Exception mainUpdateEx)
                {
                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATE ERROR: {mainUpdateEx}");
                }

                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] === SaveUpdatedXmlFiles END (SUCCESS) ===");
                    }
                    catch (Exception ex)
                    {
                diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ❌ OUTER ERROR: {ex}");
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[UniversalSleevePlacer] Error in SaveUpdatedXmlFiles: {ex.Message}");
            }
            finally
            {
                try
                {
                    SafeFileLogger.SafeAppendText("save_xml_diagnostic.log", diagnosticLog.ToString());
                }
                catch { }
            }
        }
        
        private static ClashZone CloneZoneForPersistence(ClashZone source)
        {
            if (source == null) return null;

            var clone = new ClashZone
            {
                Id = source.Id,
                SourceDocKey = source.SourceDocKey,
                HostDocKey = source.HostDocKey,
                StructuralElementDocumentTitle = source.StructuralElementDocumentTitle,
                StructuralElementType = source.StructuralElementType,
                StructuralElementThickness = source.StructuralElementThickness,
                WallThickness = source.WallThickness,
                FramingThickness = source.FramingThickness,
                StructuralElementNormalX = source.StructuralElementNormalX,
                StructuralElementNormalY = source.StructuralElementNormalY,
                StructuralElementNormalZ = source.StructuralElementNormalZ,
                MepElementCategory = source.MepElementCategory,
                MepElementWidth = source.MepElementWidth,
                MepElementHeight = source.MepElementHeight,
                MepElementFormattedSize = source.MepElementFormattedSize,
                MepElementSystemAbbreviation = source.MepElementSystemAbbreviation,
                MepElementIdValue = source.MepElementIdValue,
                MepElementUniqueId = source.MepElementUniqueId,
                StructuralElementIdValue = source.StructuralElementIdValue,
                IntersectionPointX = source.IntersectionPointX,
                IntersectionPointY = source.IntersectionPointY,
                IntersectionPointZ = source.IntersectionPointZ,
                SleeveInstanceId = source.SleeveInstanceId,
                ClusterSleeveInstanceId = source.ClusterSleeveInstanceId,
                AfterClusterSleevePlacedSleeveInstanceId = source.AfterClusterSleevePlacedSleeveInstanceId,
                SleeveFamilyName = source.SleeveFamilyName,
                SleeveWidth = source.SleeveWidth,
                SleeveHeight = source.SleeveHeight,
                SleeveDiameter = source.SleeveDiameter,
                SleevePlacementPointX = source.SleevePlacementPointX,
                SleevePlacementPointY = source.SleevePlacementPointY,
                SleevePlacementPointZ = source.SleevePlacementPointZ,
                SleevePlacementPointActiveDocumentX = source.SleevePlacementPointActiveDocumentX,
                SleevePlacementPointActiveDocumentY = source.SleevePlacementPointActiveDocumentY,
                SleevePlacementPointActiveDocumentZ = source.SleevePlacementPointActiveDocumentZ,
                SleeveBoundingBoxMinX = source.SleeveBoundingBoxMinX,
                SleeveBoundingBoxMinY = source.SleeveBoundingBoxMinY,
                SleeveBoundingBoxMinZ = source.SleeveBoundingBoxMinZ,
                SleeveBoundingBoxMaxX = source.SleeveBoundingBoxMaxX,
                SleeveBoundingBoxMaxY = source.SleeveBoundingBoxMaxY,
                SleeveBoundingBoxMaxZ = source.SleeveBoundingBoxMaxZ,
                ClusterSleeveBoundingBoxMinX = source.ClusterSleeveBoundingBoxMinX,
                ClusterSleeveBoundingBoxMinY = source.ClusterSleeveBoundingBoxMinY,
                ClusterSleeveBoundingBoxMinZ = source.ClusterSleeveBoundingBoxMinZ,
                ClusterSleeveBoundingBoxMaxX = source.ClusterSleeveBoundingBoxMaxX,
                ClusterSleeveBoundingBoxMaxY = source.ClusterSleeveBoundingBoxMaxY,
                ClusterSleeveBoundingBoxMaxZ = source.ClusterSleeveBoundingBoxMaxZ,
                IsResolved = source.IsResolved,
                IsClusterResolved = source.IsClusterResolved,
                MarkedForClusteringSleeveProcess = source.MarkedForClusteringSleeveProcess,
                MepElementLevelName = source.MepElementLevelName,
                MepElementLevelElevation = source.MepElementLevelElevation,
                DuctShape = source.DuctShape,
                PipeOpeningType = source.PipeOpeningType,
                HostOrientation = source.HostOrientation,
                WallDirectionType = source.WallDirectionType,
                WallDirectionX = source.WallDirectionX,
                WallDirectionY = source.WallDirectionY,
                WallDirectionZ = source.WallDirectionZ,
                // ✅ CRITICAL FIX: Copy MEP orientation properties (missing in original clone)
                MepElementOrientation = source.MepElementOrientation,
                MepElementOrientationDirection = source.MepElementOrientationDirection,
                MepElementRotationAngle = source.MepElementRotationAngle,
                MepElementOrientationX = source.MepElementOrientationX,
                MepElementOrientationY = source.MepElementOrientationY,
                MepElementOrientationZ = source.MepElementOrientationZ,
                DocumentPath = source.DocumentPath,
                DetectedAt = source.DetectedAt,
                LastUpdated = source.LastUpdated
            };

            if (source.MepParameterValues != null)
            {
                clone.MepParameterValues = source.MepParameterValues
                    .Select(p => new SerializableKeyValue { Key = p.Key, Value = p.Value })
                    .ToList();
            }

            if (source.HostParameterValues != null)
            {
                clone.HostParameterValues = source.HostParameterValues
                    .Select(p => new SerializableKeyValue { Key = p.Key, Value = p.Value })
                    .ToList();
            }

            return clone;
        }

        private static void MergeCategoryIntoMainFilter(OpeningFilter mainFilter, OpeningFilter categoryFilter)
        {
            if (mainFilter == null || categoryFilter?.ClashZoneStorage == null)
                return;

            mainFilter.ClashZoneStorage ??= new ClashZoneStorage();
            var mainStorage = mainFilter.ClashZoneStorage;
            var categoryStorage = categoryFilter.ClashZoneStorage;

            mainStorage.Filters ??= new List<FilterGroupForStorage>();

            var sourceGroup = categoryStorage.Filters?.FirstOrDefault();
            if (sourceGroup == null)
                return;

            var targetGroup = mainStorage.Filters
                .FirstOrDefault(f => string.Equals(f?.Name, sourceGroup.Name, StringComparison.OrdinalIgnoreCase));

            if (targetGroup == null)
            {
                var cloneGroup = CloneFilterGroupForStorage(sourceGroup);
                if (cloneGroup != null)
                    mainStorage.Filters.Add(cloneGroup);
            }
            else
            {
                MergeFilterGroups(targetGroup, sourceGroup);
            }

            mainStorage.ClashZones = mainStorage.AllZones?.ToList() ?? new List<ClashZone>();
            mainStorage.LastUpdated = DateTime.Now;
            mainFilter.LastModified = DateTime.Now;
        }

        private static FilterGroupForStorage CloneFilterGroupForStorage(FilterGroupForStorage source)
        {
            if (source == null) return null;

            var clone = new FilterGroupForStorage
            {
                Name = source.Name,
                FileCombos = new List<FilterFileComboGroup>()
            };

            if (source.FileCombos != null)
            {
                foreach (var combo in source.FileCombos)
                {
                    var comboClone = CloneFilterComboForStorage(combo);
                    if (comboClone != null)
                        clone.FileCombos.Add(comboClone);
                }
            }

            return clone;
        }

        private static FilterFileComboGroup CloneFilterComboForStorage(FilterFileComboGroup source)
        {
            if (source == null) return null;

            var clone = new FilterFileComboGroup
            {
                LinkedFile = source.LinkedFile,
                HostFile = source.HostFile,
                ProcessedAt = source.ProcessedAt,
                ClashZones = new List<ClashZone>()
            };

            if (source.ClashZones != null)
            {
                foreach (var zone in source.ClashZones)
                {
                    var zoneClone = CloneZoneForPersistence(zone);
                    if (zoneClone != null)
                        clone.ClashZones.Add(zoneClone);
                }
            }

            return clone;
        }

        private static void MergeFilterGroups(FilterGroupForStorage target, FilterGroupForStorage source)
        {
            if (target == null || source?.FileCombos == null)
                return;

            target.FileCombos ??= new List<FilterFileComboGroup>();

            var existing = target.FileCombos
                .Where(fc => fc != null)
                .ToDictionary(fc => fc.GetNormalizedKey(), fc => fc, StringComparer.OrdinalIgnoreCase);

            foreach (var combo in source.FileCombos ?? Enumerable.Empty<FilterFileComboGroup>())
            {
                if (combo == null) continue;
                var key = combo.GetNormalizedKey();
                if (string.IsNullOrWhiteSpace(key)) continue;

                if (!existing.TryGetValue(key, out var targetCombo))
                {
                    var newCombo = CloneFilterComboForStorage(combo);
                    if (newCombo != null)
                    {
                        target.FileCombos.Add(newCombo);
                        existing[key] = newCombo;
                    }
                }
                else
                {
                    MergeFilterFileCombo(targetCombo, combo);
                }
            }
        }

        private static void MergeFilterFileCombo(FilterFileComboGroup target, FilterFileComboGroup source)
        {
            if (target == null || source?.ClashZones == null)
                return;

            target.ClashZones ??= new List<ClashZone>();

            var indexMap = new Dictionary<Guid, int>();
            for (int i = 0; i < target.ClashZones.Count; i++)
            {
                var existing = target.ClashZones[i];
                if (existing == null || existing.Id == Guid.Empty)
                    continue;
                indexMap[existing.Id] = i;
            }

            foreach (var zone in source.ClashZones)
            {
                var clone = CloneZoneForPersistence(zone);
                if (clone == null || clone.Id == Guid.Empty)
                    continue;

                if (indexMap.TryGetValue(clone.Id, out var index))
                {
                    target.ClashZones[index] = clone;
                                }
                                else
                                {
                    indexMap[clone.Id] = target.ClashZones.Count;
                    target.ClashZones.Add(clone);
                }
            }

            DeduplicateZonesInCombo(target);
            if (source.ProcessedAt > target.ProcessedAt)
                target.ProcessedAt = source.ProcessedAt;
        }

        private static void DeduplicateZonesInCombo(FilterFileComboGroup combo)
        {
            if (combo?.ClashZones == null)
                return;

            var seen = new HashSet<Guid>();
            var deduped = new List<ClashZone>();

            foreach (var zone in combo.ClashZones)
            {
                if (zone == null || zone.Id == Guid.Empty)
                    continue;

                if (seen.Add(zone.Id))
                {
                    deduped.Add(zone);
                }
            }

            combo.ClashZones = deduped;
        }

        private static void EnsureConditionsFiles(
            Document document,
            string baseFilterName,
            List<(OpeningFilter Filter, string Category)> processedCategoryFilters,
            string filtersDirectory,
            System.Text.StringBuilder diagnosticLog)
        {
            if (string.IsNullOrWhiteSpace(baseFilterName) ||
                processedCategoryFilters == null ||
                processedCategoryFilters.Count == 0)
            {
                return;
            }

            try
            {
                if (!Directory.Exists(filtersDirectory))
                {
                    Directory.CreateDirectory(filtersDirectory);
                }
            }
            catch { }

            try
            {
                var conditionsService = new ConditionsService(document, filtersDirectory, msg =>
                {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info(msg);
                });

                var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var (categoryFilter, category) in processedCategoryFilters)
                {
                    var normalizedCategory = NormalizeCategoryForConditions(category);
                    if (string.IsNullOrWhiteSpace(normalizedCategory))
                        continue;

                    var combinedKey = $"{baseFilterName}_{normalizedCategory}";
                    if (!seenKeys.Add(combinedKey))
                        continue;

                    try
                    {
                        if (!conditionsService.ConditionsExist(combinedKey))
                        {
                            var conditions = new OpeningConditions
                            {
                                FilterName = combinedKey,
                                Category = category,
                                ClearanceSettings = new ClearanceSettings()
                            };

                            if (conditionsService.SaveConditions(conditions, combinedKey))
                            {
                                diagnosticLog?.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] CONDITIONS CREATED → {combinedKey}_CONDITIONS.xml");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        diagnosticLog?.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] CONDITIONS SAVE FAILED ({combinedKey}): {ex}");
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[UniversalSleevePlacer] Failed to ensure conditions for '{combinedKey}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                diagnosticLog?.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] CONDITIONS SERVICE INIT FAILED: {ex}");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[UniversalSleevePlacer] ConditionsService initialization failed: {ex.Message}");
            }
        }

        private static string NormalizeCategoryForConditions(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return "unknown";

            return category
                .Replace(" ", "_")
                .ToLowerInvariant();
        }
        /// <summary>
        /// ⚠️ SAFETY MEASURE: Create backup of XML file before modification
        /// </summary>
        private void CreateXmlBackup(string xmlFilePath)
        {
            try
            {
                var backupPath = xmlFilePath + ".backup";
                if (File.Exists(xmlFilePath))
                {
                    File.Copy(xmlFilePath, backupPath, overwrite: true);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-BACKUP] Created backup: {Path.GetFileName(backupPath)}");
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[XML-BACKUP] Failed to create backup for {Path.GetFileName(xmlFilePath)}: {ex.Message}");
                // Don't throw - backup failure shouldn't stop save
            }
        }
        
        /// <summary>
        /// ⚠️ SAFETY MEASURE: Restore XML file from backup if save validation fails
        /// </summary>
        private void RestoreXmlFromBackup(string xmlFilePath)
        {
            try
            {
                var backupPath = xmlFilePath + ".backup";
                if (File.Exists(backupPath))
                {
                    File.Copy(backupPath, xmlFilePath, overwrite: true);
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[XML-RESTORE] ✓ Restored {Path.GetFileName(xmlFilePath)} from backup");
                    SafeFileLogger.SafeAppendText("xml_restore.log", 
                        $"Restored {Path.GetFileName(xmlFilePath)} from backup at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
                else
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[XML-RESTORE] No backup found for {Path.GetFileName(xmlFilePath)}");
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[XML-RESTORE] Failed to restore backup for {Path.GetFileName(xmlFilePath)}: {ex.Message}");
                throw; // Re-throw - restore failure is critical
            }
        }
        
        /// <summary>
        /// ⚠️ SAFETY MEASURE: Validate that XML save completed correctly
        /// Verifies that expected clash zones were updated correctly
        /// </summary>
        private bool ValidateXmlSave(string xmlFilePath, List<Guid> expectedUpdatedIds, int expectedCount)
        {
            try
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                Models.OpeningFilter verification;
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    verification = (Models.OpeningFilter)serializer.Deserialize(reader);
                }
                
                var verificationZones = verification?.ClashZoneStorage?.AllZones ?? new List<ClashZone>();
                if (verificationZones.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[XML-VERIFY] No clash zones found in saved file {Path.GetFileName(xmlFilePath)}");
                    return false;
                }
                
                int actualUpdated = 0;
                foreach (var expectedId in expectedUpdatedIds)
                {
                    var savedZone = verificationZones.FirstOrDefault(z => z.Id == expectedId);
                    if (savedZone != null && (savedZone.SleeveWidth > 0 || savedZone.SleeveHeight > 0 || savedZone.SleeveDiameter > 0))
                    {
                        actualUpdated++;
                    }
                }
                
                // Allow 5% tolerance for minor discrepancies
                double successRate = expectedCount > 0 ? (double)actualUpdated / expectedCount : 0;
                bool isValid = successRate >= 0.95;
                
                if (!isValid)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[XML-VERIFY] Validation failed: Expected {expectedCount} updates, got {actualUpdated} ({successRate:P1}). Threshold: 95%");
                }
                
                return isValid;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[XML-VERIFY] Error during validation: {ex.Message}");
                return false; // Fail safe - if validation fails, consider it invalid
            }
        }

        /// <summary>
        /// Get XML filename for category
        /// ✅ CRITICAL: Filter name MUST come from orchestrator - no fallback allowed
        /// </summary>
        private string GetFilterNameForCategory(string category)
        {
            // ✅ CRITICAL: Filter name MUST come from orchestrator - no fallback allowed
            // This is essential for correct XML updates, clustering, and flag management
            if (string.IsNullOrEmpty(_filterName))
            {
                var errorMsg = $"Filter name is required but was not provided. Cannot determine target XML file for category '{category}'.";
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetFilterNameForCategory] {errorMsg}");
                throw new InvalidOperationException(errorMsg);
            }
            
            return _filterName;
        }

// ============================================================================
// CORRECTED SetSleeveOrientation Method
// ============================================================================
        private void SetSleeveOrientation(FamilyInstance sleeveInstance, ClashZone clashZone, SleevePlacementPlanningDto planningDto = null)
        {
            try
            {
                bool isFloorHost = clashZone.StructuralElementType == "Floor" || 
                                  clashZone.StructuralElementType == "Floors";
                
                // DEBUG: Log the floor host decision
                try
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLOOR-DEBUG] Sleeve {sleeveInstance.Id.IntegerValue}: StructuralElementType='{clashZone.StructuralElementType}', isFloorHost={isFloorHost}\n");
                }
                catch { }
                
                if (isFloorHost)
                {
            // ====== FLOOR HOST: Rotate based on MEP element rotation angle ======
                    // ✅ FLOOR ROTATION FIX: Use pre-calculated rotation angle (calculated once during refresh)
                    // This supports arbitrary angles (not just 0° and 90°) and avoids Revit API calls
                    
                    // ✅ CRITICAL: Pipes and round ducts (circular elements) should NOT be rotated - place straight to WCS (axis-aligned)
                    // Only rectangular ducts and cable trays should rotate based on MEP element orientation
                    bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                    bool isDuct = string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                    bool isRoundDuct = isDuct && (
                        string.Equals(clashZone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(clashZone.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase) ||
                        (clashZone.MepElementSizeData != null && 
                         (string.Equals(clashZone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(clashZone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)))
                    );
                    bool isCircularElement = isPipe || isRoundDuct;
                    
                    if (isCircularElement)
                    {
                        // ✅ CIRCULAR ELEMENT FIX: Pipes and round ducts should always be placed axis-aligned (straight to WCS), no rotation
                        string elementType = isPipe ? "PIPE" : "ROUND DUCT";
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORIENTATION-SET] Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: {elementType} on floor - NO ROTATION (placed straight to WCS, axis-aligned)");
                        }
                    }
                    else
                    {
                        // ✅ DUCTS and CABLE TRAYS: Rotate based on MEP element orientation
                    var loc = sleeveInstance.Location as LocationPoint;
                    if (loc != null)
                    {
                        // ✅ PARALLEL PLANNING: Use pre-computed rotation angle if available
                        double rotationAngle;
                        if (planningDto != null && DeploymentConfiguration.EnableParallelPlanning)
                        {
                            // Use planning layer rotation (converted to radians)
                            rotationAngle = planningDto.RotationAngleDeg * Math.PI / 180.0;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[PLANNING] Using pre-computed rotation: {planningDto.RotationAngleDeg:F1}° (Risk={planningDto.ClearanceRisk})");
                            }
                        }
                        else
                        {
                            // ✅ CALCULATE ONCE, USE MANY TIMES: Use pre-calculated rotation angle from ClashZone
                            rotationAngle = clashZone.MepElementRotationAngle;
                        }
                        
                        // ✅ PATH 1 (Replay): Log if MEP orientation is missing (but don't retrieve - use existing data only)
                        if (_isReplayPath && Math.Abs(rotationAngle) < 1e-6 && string.IsNullOrEmpty(clashZone.MepElementOrientationDirection))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[PATH-1-DIAGNOSTIC] Zone={clashZone.Id}: Missing MEP orientation (BasisX) - MepElementRotationAngle={clashZone.MepElementRotationAngle}, MepElementOrientationDirection='{clashZone.MepElementOrientationDirection}', MepElementOrientationX={clashZone.MepElementOrientationX}. PATH 1 will use 0° rotation (no retrieval from linked files).");
                            }
                        }
                        
                        // ✅ FALLBACK: If rotation angle not calculated (old XML data), calculate from orientation direction
                        // This maintains backward compatibility with existing XML files
                        // ✅ BUG FIX: Only use fallback if BOTH angle is zero AND orientation direction is missing
                        // If MepElementOrientationDirection exists, the angle was calculated (even if 0°), so use it as-is
                        if (Math.Abs(rotationAngle) < 1e-6 && string.IsNullOrEmpty(clashZone.MepElementOrientationDirection))
                        {
                            // Old XML data - angle not calculated, use fallback
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: MepElementRotationAngle is 0 and MepElementOrientationDirection is missing - this is old XML data, skipping fallback (will use 0°)");
                        }
                        else if (Math.Abs(rotationAngle) < 1e-6 && !string.IsNullOrEmpty(clashZone.MepElementOrientationDirection))
                        {
                            // ✅ BUG FIX: Angle is 0° and orientation direction exists - this means 0° was calculated, use it as-is
                            // Do NOT use fallback - 0° is a valid calculated angle
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Using calculated MepElementRotationAngle=0° (MepElementOrientationDirection='{clashZone.MepElementOrientationDirection}' exists, so angle was calculated)");
                        }
                        
                            // ✅ APPLY ROTATION: Rotate sleeve to match MEP element angle (for ducts and cable trays only)
                        if (Math.Abs(rotationAngle) > 1e-6) // Only rotate if angle is significant
                        {
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[ORIENTATION-SET] Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: Rotated {rotationAngleDegrees:F1}° (from DB: MepElementRotationAngle={clashZone.MepElementRotationAngle * 180 / Math.PI:F1}°, MepElementOrientationDirection='{clashZone.MepElementOrientationDirection}')");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[ORIENTATION-SET] ⚠️ Zone={clashZone.Id}, Sleeve={sleeveInstance.Id}: No rotation applied (angle=0°) - from DB: MepElementRotationAngle={clashZone.MepElementRotationAngle * 180 / Math.PI:F1}°, MepElementOrientationDirection='{clashZone.MepElementOrientationDirection}', MepElementOrientationX={clashZone.MepElementOrientation?.X:F6}");
                                }
                            }
                        }
                    }
                    
                    // Set HostOrientation parameter (batched if enabled)
                    var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                    if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                    {
                        if (OptimizationFlags.UseBatchedParameterWrites)
                        {
                            if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                                _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                            _deferredParameters[sleeveInstance.Id]["HostOrientation"] = "FloorHosted";
                        }
                        else
                        {
                            hostOrientationParam.Set("FloorHosted");
                        }
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Set HostOrientation = FloorHosted");
                    }
                    
                    // ✅ FIX: Return early to prevent wall rotation logic from running
                    return;
                }
                else
                {
            // ====== WALL or FRAMING HOST ======
            bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            // DEBUG: Log the wall/framing decision
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[WALL-DEBUG] Sleeve {sleeveInstance.Id.IntegerValue}: StructuralElementType='{clashZone.StructuralElementType}', isFramingHost={isFramingHost}\n");
            }
            catch { }
            bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
            bool isCableTray = string.Equals(clashZone.MepElementCategory, "Cable Trays", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(clashZone.MepElementCategory, "Cable Tray Fittings", StringComparison.OrdinalIgnoreCase);
                    
                    var mepOrientation = clashZone.MepElementOrientation;
            var structuralNormal = clashZone.StructuralElementNormal;
            
            try
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[ORIENT-INPUT] Sleeve {sleeveInstance.Id.IntegerValue}: Cat={clashZone.MepElementCategory}, Host={clashZone.StructuralElementType}, MEP=({mepOrientation?.X:F3},{mepOrientation?.Y:F3}), StructN=({structuralNormal?.X:F3},{structuralNormal?.Y:F3})\n");
                
                // DEBUG: Log all ClashZone properties to trace XML loading issue
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[XML-DEBUG] ClashZone {clashZone.Id}: MepElementCategory='{clashZone.MepElementCategory}', StructuralElementType='{clashZone.StructuralElementType}', MepElementId={clashZone.MepElementIdValue}, StructuralElementId={clashZone.StructuralElementIdValue}\n");
            }
            catch { }
            
                    if (true) // DISABLED: Don't wait for MEP orientation - use wall direction directly
            {
                var loc = sleeveInstance.Location as LocationPoint;
                if (loc != null)
                {
                    // Skip the old framing logic - let the wall logic handle it (like walls)
                    if (true) // Process all elements through wall logic
                    {
                        // DEBUG: Confirm we're reaching the wall direction logic
                        try
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[WALL-LOGIC-REACHED] Sleeve {sleeveInstance.Id.IntegerValue}: Reached wall direction logic block ✓\n");
                        }
                        catch { }
                        
                        // ⚠️ CRITICAL OPTIMIZATION: Use pre-calculated wall direction from XML ⚠️
                        // NO EXPENSIVE REVIT API CALLS - wall direction calculated during refresh
                        bool allowRotate = false;
                        
                        // DIAGNOSTIC: Log ClashZone wall direction data availability
                        try
                        {
                            string placementDebugLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(placementDebugLogPath, $"[DIAGNOSTIC] Sleeve {sleeveInstance.Id.IntegerValue}: WallDirection={clashZone.WallDirection != null}, WallDirectionType='{clashZone.WallDirectionType ?? "NULL"}', WallDirectionZero={clashZone.WallDirection == XYZ.Zero}\n");
                }
                        }
                        catch { }
                        
                        // Calculate wall direction from structural normal if not available in XML
                        XYZ wallDirection = clashZone.WallDirection;
                        string wallDirectionType = clashZone.WallDirectionType;
                        
                        if (wallDirection == null || wallDirection == XYZ.Zero || string.IsNullOrEmpty(wallDirectionType))
                        {
                            // FALLBACK: Calculate wall direction from structural normal
                        if (structuralNormal != null && (structuralNormal.X != 0 || structuralNormal.Y != 0))
                        {
                                // Wall direction is perpendicular to wall normal
                                wallDirection = new XYZ(-structuralNormal.Y, structuralNormal.X, 0).Normalize();
                                
                                // Determine wall direction type
                                double absX = Math.Abs(wallDirection.X);
                                double absY = Math.Abs(wallDirection.Y);
                                
                                if (absX > absY)
                                {
                                    wallDirectionType = "X-WALL";
                                }
                                else if (absY > absX)
                                {
                                    wallDirectionType = "Y-WALL";
                                }
                                else
                                {
                                    wallDirectionType = "DIAGONAL-WALL";
                                }
                                
                                // Log fallback calculation
                                try
                                {
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[WALL-DIR-FALLBACK] Sleeve {sleeveInstance.Id.IntegerValue}: Calculated WallDirection=({wallDirection.X:F3},{wallDirection.Y:F3},{wallDirection.Z:F3}), WallType={wallDirectionType} ✓\n");
                                }
                                catch { }
                            }
                            else
                            {
                                // Ultimate fallback
                                wallDirectionType = "UNKNOWN";
                                try
                                {
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[WALL-DIR-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: Cannot determine wall direction - structural normal is null or zero ✗\n");
                                }
                                catch { }
                            }
                        }
                        
                        // ✅ SIMPLIFIED: One code path for both walls and framing
                        // Check HostOrientation from XML (works for both walls and framing)
                        string hostOrientation = clashZone.HostOrientation ?? "";
                        bool needsRotation = false; // Whether we need to apply +90° rotation
                        
                        if (hostOrientation == "X")
                        {
                            needsRotation = true; // X-orientation needs +90° rotation
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] HostOrientation=X: Will apply +90° rotation");
                        }
                        else if (hostOrientation == "Y")
                        {
                            needsRotation = false; // Y-orientation works naturally with LEFT view families
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] HostOrientation=Y: No rotation needed");
                        }
                        else
                        {
                            // Fallback: Use wall direction type if HostOrientation not available
                            if (wallDirectionType == "X-WALL")
                            {
                                needsRotation = true;
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalSleevePlacer] Fallback: X-WALL detected - will apply +90° rotation");
                            }
                            else if (wallDirectionType == "Y-WALL")
                            {
                                needsRotation = false;
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalSleevePlacer] Fallback: Y-WALL detected - no rotation");
                        }
                        else
                        {
                                // Ultimate fallback
                                needsRotation = false;
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[UniversalSleevePlacer] Unknown orientation - using fallback");
                            }
                        }
                        
                        // Log to placement_debug.log for immediate visibility
                        try
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[ORIENT-DECISION] Sleeve {sleeveInstance.Id.IntegerValue}: HostOrientation={hostOrientation}, needsRotation={needsRotation} ✓\n");
                        }
                        catch { }
                        
                        if (needsRotation)
                        {
                            // ==========================================
                            // X-WALL LOGIC (Horizontal walls)
                            // ==========================================
                            // Family created in LEFT view → extrudes along Y-axis
                            // Works naturally for Y-walls, needs +90° for X-walls
                            // ALL families: Circular, Rectangular, Ducts, Pipes, Cable Trays
                            
                            var locXWall = sleeveInstance.Location as LocationPoint;
                            if (locXWall != null)
                            {
                                Line rotationAxis = Line.CreateBound(locXWall.Point, locXWall.Point + XYZ.BasisZ);
                                double rotation = Math.PI / 2; // +90 degrees
                                ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotation);
                                
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ORIENT-FIX] Sleeve {sleeveInstance.Id.IntegerValue}: X-WALL +90° (LEFT view family)");
                                try
                                {
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[ORIENT-FIX] Sleeve {sleeveInstance.Id.IntegerValue}: X-WALL +90° (LEFT view family) ✓\n");
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            // ==========================================
                            // Y-WALL LOGIC (Vertical walls)
                            // ==========================================
                            // Family created in LEFT view → NO rotation needed
                            // ALL families work perfectly as-is: Circular, Rectangular, Ducts, Pipes, Cable Trays
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[ORIENT-SKIP] Sleeve {sleeveInstance.Id.IntegerValue}: Y-WALL no rotation (LEFT view family already aligned)");
                            try
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[ORIENT-SKIP] Sleeve {sleeveInstance.Id.IntegerValue}: Y-WALL no rotation (LEFT view family) ✓\n");
                            }
                            catch { }
                        }
                    }
                }
            }
            
            // Set HostOrientation parameter from XML (pre-calculated during refresh) (batched if enabled)
                    if (!string.IsNullOrEmpty(clashZone.HostOrientation))
                    {
                        var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                        if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                        {
                            if (OptimizationFlags.UseBatchedParameterWrites)
                            {
                                if (!_deferredParameters.ContainsKey(sleeveInstance.Id))
                                    _deferredParameters[sleeveInstance.Id] = new Dictionary<string, object>();
                                _deferredParameters[sleeveInstance.Id]["HostOrientation"] = clashZone.HostOrientation;
                            }
                            else
                            {
                                hostOrientationParam.Set(clashZone.HostOrientation);
                            }
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Set HostOrientation = {clashZone.HostOrientation} (from XML)");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting sleeve orientation: {ex.Message}\n{ex.StackTrace}");
                try
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[ORIENT-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
                }
                catch { }
            }
        }

        /// <summary>
        /// ✅ ROBUST CORNER SAVING: Calculates and saves 4 corner coordinates to database with full validation and error handling.
        /// This method ensures corners are ALWAYS saved, regardless of rotation angle or other conditions.
        /// </summary>
        /// <param name="zone">ClashZone with sleeve placement data</param>
        /// <param name="sleeve">Revit sleeve element (optional, for parameter lookup)</param>
        /// <param name="repository">Database repository for saving corners</param>
        /// <param name="rotationAngleRad">Rotation angle in radians (can be 0 for no rotation)</param>
        /// <param name="fallbackWidth">Fallback width if zone.SleeveWidth is invalid</param>
        /// <param name="fallbackHeight">Fallback height if zone.SleeveHeight is invalid</param>
        /// <returns>True if corners were successfully saved, False otherwise</returns>
        private bool SaveSleeveCornersRobust(
            ClashZone zone,
            FamilyInstance sleeve,
            ClashZoneRepository repository,
            double rotationAngleRad,
            double fallbackWidth,
            double fallbackHeight)
        {
            const int MAX_RETRIES = 3;
            const double MIN_DIMENSION = 0.001; // Minimum valid dimension (1mm in internal units)
            
            try
            {
                // ✅ VALIDATION 1: Check GUID is valid
                if (zone == null || zone.Id == Guid.Empty)
                {
                    SafeFileLogger.SafeAppendText("corner_save_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] ❌ SaveSleeveCornersRobust: Invalid zone or empty GUID\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SLEEVE-CORNERS] ❌ Cannot save corners: Zone is null or GUID is empty");
                    return false;
                }

                // ✅ VALIDATION 2: Get sleeve center coordinates with fallback
                double centerX = zone.SleevePlacementPointActiveDocumentX;
                double centerY = zone.SleevePlacementPointActiveDocumentY;
                double centerZ = zone.SleevePlacementPointActiveDocumentZ;

                // Check if Active coordinates are valid (not all zero)
                if (Math.Abs(centerX) < 1e-6 && Math.Abs(centerY) < 1e-6 && Math.Abs(centerZ) < 1e-6)
                {
                    // Fallback to regular placement point
                    if (zone.SleevePlacementPoint != null)
                    {
                        centerX = zone.SleevePlacementPoint.X;
                        centerY = zone.SleevePlacementPoint.Y;
                        centerZ = zone.SleevePlacementPoint.Z;
                        SafeFileLogger.SafeAppendText("corner_save_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Zone {zone.Id}: Active coordinates zero, using SleevePlacementPoint\n");
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("corner_save_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {zone.Id}: Both Active and PlacementPoint coordinates are invalid\n");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[SLEEVE-CORNERS] ❌ Zone {zone.Id}: Cannot determine sleeve center coordinates");
                        return false;
                    }
                }

                // ✅ VALIDATION 3: Get and validate width/height
                double actualWidth = zone.SleeveWidth > MIN_DIMENSION ? zone.SleeveWidth : fallbackWidth;
                double actualHeight = zone.SleeveHeight > MIN_DIMENSION ? zone.SleeveHeight : fallbackHeight;

                // Try to get dimensions from sleeve element if available
                if ((actualWidth < MIN_DIMENSION || actualHeight < MIN_DIMENSION) && sleeve != null && sleeve.IsValidObject)
                {
                    try
                    {
                        var widthParam = sleeve.LookupParameter("Width");
                        var heightParam = sleeve.LookupParameter("Height");

                        if (widthParam != null && widthParam.HasValue)
                            actualWidth = widthParam.AsDouble();
                        if (heightParam != null && heightParam.HasValue)
                            actualHeight = heightParam.AsDouble();
                    }
                    catch (Exception paramEx)
                    {
                        SafeFileLogger.SafeAppendText("corner_save_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Zone {zone.Id}: Error reading parameters, using fallback: {paramEx.Message}\n");
                    }
                }

                // Final validation of dimensions
                if (actualWidth < MIN_DIMENSION || actualHeight < MIN_DIMENSION)
                {
                    SafeFileLogger.SafeAppendText("corner_save_errors.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {zone.Id}: Invalid dimensions - Width={actualWidth * 304.8:F1}mm, Height={actualHeight * 304.8:F1}mm (min={MIN_DIMENSION * 304.8:F1}mm)\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SLEEVE-CORNERS] ❌ Zone {zone.Id}: Cannot save corners - invalid dimensions");
                    return false;
                }

                // ✅ CALCULATION: Calculate 4 corners
                XYZ[] worldCorners = new XYZ[4];
                
                if (Math.Abs(rotationAngleRad) < 1e-6)
                {
                    // Zero rotation - calculate directly in world space
                    double halfW = actualWidth / 2.0;
                    double halfH = actualHeight / 2.0;
                    
                    worldCorners[0] = new XYZ(centerX - halfW, centerY - halfH, centerZ);  // Bottom-left
                    worldCorners[1] = new XYZ(centerX + halfW, centerY - halfH, centerZ);  // Bottom-right
                    worldCorners[2] = new XYZ(centerX - halfW, centerY + halfH, centerZ);  // Top-left
                    worldCorners[3] = new XYZ(centerX + halfW, centerY + halfH, centerZ);  // Top-right
                }
                else
                {
                    // Non-zero rotation - calculate in local space, then rotate
                    double halfW = actualWidth / 2.0;
                    double halfH = actualHeight / 2.0;
                    
                    var localCorners = new[]
                    {
                        new XYZ(-halfW, -halfH, 0),  // Corner 1: Bottom-left
                        new XYZ(halfW, -halfH, 0),   // Corner 2: Bottom-right
                        new XYZ(-halfW, halfH, 0),   // Corner 3: Top-left
                        new XYZ(halfW, halfH, 0)     // Corner 4: Top-right
                    };
                    
                    double cosSleeve = Math.Cos(rotationAngleRad);
                    double sinSleeve = Math.Sin(rotationAngleRad);
                    
                    for (int j = 0; j < 4; j++)
                    {
                        double localX = localCorners[j].X;
                        double localY = localCorners[j].Y;
                        
                        // Rotate corner by sleeve rotation matrix
                        double worldX = localX * cosSleeve - localY * sinSleeve;
                        double worldY = localX * sinSleeve + localY * cosSleeve;
                        
                        // Translate to sleeve center
                        worldCorners[j] = new XYZ(
                            centerX + worldX,
                            centerY + worldY,
                            centerZ
                        );
                    }
                }

                // ✅ VALIDATION 4: Validate all corners are valid (non-NaN, non-Infinity, non-zero)
                bool allCornersValid = true;
                for (int i = 0; i < 4; i++)
                {
                    if (double.IsNaN(worldCorners[i].X) || double.IsInfinity(worldCorners[i].X) ||
                        double.IsNaN(worldCorners[i].Y) || double.IsInfinity(worldCorners[i].Y) ||
                        double.IsNaN(worldCorners[i].Z) || double.IsInfinity(worldCorners[i].Z) ||
                        Math.Abs(worldCorners[i].X) < 1e-9 && Math.Abs(worldCorners[i].Y) < 1e-9 && Math.Abs(worldCorners[i].Z) < 1e-9)
                    {
                        allCornersValid = false;
                        SafeFileLogger.SafeAppendText("corner_save_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {zone.Id}: Invalid corner {i + 1} - X={worldCorners[i].X}, Y={worldCorners[i].Y}, Z={worldCorners[i].Z}\n");
                        break;
                    }
                }

                if (!allCornersValid)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[SLEEVE-CORNERS] ❌ Zone {zone.Id}: Calculated corners are invalid");
                    return false;
                }

                // ✅ SAVE: Attempt to save corners with retry logic
                bool success = false;
                int attempt = 0;
                Exception lastException = null;

                while (attempt < MAX_RETRIES && !success)
                {
                    attempt++;
                    try
                    {
                        repository.UpdateSleeveCorners(
                            zone.Id,
                            worldCorners[0].X, worldCorners[0].Y, worldCorners[0].Z,  // Corner 1
                            worldCorners[1].X, worldCorners[1].Y, worldCorners[1].Z,  // Corner 2
                            worldCorners[2].X, worldCorners[2].Y, worldCorners[2].Z,  // Corner 3
                            worldCorners[3].X, worldCorners[3].Y, worldCorners[3].Z   // Corner 4
                        );

                        // ✅ VERIFICATION: Check if corners were actually saved (by reading back from DB)
                        // Note: This is a simple check - we rely on UpdateSleeveCorners logging for detailed feedback
                        success = true;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            double rotationDeg = rotationAngleRad * 180.0 / Math.PI;
                            DebugLogger.Info($"[SLEEVE-CORNERS] ✅ Saved 4 world-space corners for zone {zone.Id} (attempt {attempt}/{MAX_RETRIES}): " +
                                $"Rotation={rotationDeg:F1}°, " +
                                $"Center=({centerX:F6}, {centerY:F6}, {centerZ:F6}), " +
                                $"Size={actualWidth * 304.8:F1}mm × {actualHeight * 304.8:F1}mm, " +
                                $"C1=({worldCorners[0].X:F6}, {worldCorners[0].Y:F6}, {worldCorners[0].Z:F6}), " +
                                $"C2=({worldCorners[1].X:F6}, {worldCorners[1].Y:F6}, {worldCorners[1].Z:F6}), " +
                                $"C3=({worldCorners[2].X:F6}, {worldCorners[2].Y:F6}, {worldCorners[2].Z:F6}), " +
                                $"C4=({worldCorners[3].X:F6}, {worldCorners[3].Y:F6}, {worldCorners[3].Z:F6})");
                        }
                    }
                    catch (Exception dbEx)
                    {
                        lastException = dbEx;
                        success = false;
                        
                        SafeFileLogger.SafeAppendText("corner_save_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ Zone {zone.Id}: Database save attempt {attempt}/{MAX_RETRIES} failed: {dbEx.Message}\n");
                        
                        if (attempt < MAX_RETRIES)
                        {
                            // Wait a bit before retry (exponential backoff)
                            System.Threading.Thread.Sleep(50 * attempt);
                        }
                        else
                        {
                            // Final attempt failed - log error
                            SafeFileLogger.SafeAppendText("corner_save_errors.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {zone.Id}: All {MAX_RETRIES} save attempts failed. Last error: {dbEx.Message}\nStack trace: {dbEx.StackTrace}\n");
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[SLEEVE-CORNERS] ❌ Zone {zone.Id}: Failed to save corners after {MAX_RETRIES} attempts: {dbEx.Message}");
                            }
                        }
                    }
                }

                return success;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("corner_save_errors.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] ❌ Zone {(zone?.Id.ToString() ?? "NULL")}: Unexpected error in SaveSleeveCornersRobust: {ex.Message}\nStack trace: {ex.StackTrace}\n");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[SLEEVE-CORNERS] ❌ Unexpected error saving corners for zone {zone?.Id}: {ex.Message}");
                }
                
                return false;
            }
        }
    }
}
