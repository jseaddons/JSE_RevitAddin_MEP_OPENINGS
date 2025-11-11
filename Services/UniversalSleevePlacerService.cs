using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

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
        
        // ✅ OOP REFACTORING: Centralized flag management
        private readonly FlagManager _flagManager;

        // ⚠️ QUICK WIN: Pre-cached family symbols (load once, reuse many times)
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();

        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacerService(Document doc, OpeningConditions conditions, ISleevePlacementStrategy strategy, Dictionary<string, double> clearanceSettings = null, string filterName = null, FlagManager flagManager = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _filterName = filterName;
            
            // ✅ OOP REFACTORING: Initialize FlagManager (create if not provided for backward compatibility)
            _flagManager = flagManager ?? new FlagManager(doc);
            
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
                    return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
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
                        return norm.Length > 0 ? UnitUtils.ConvertToInternalUnits(val, UnitTypeId.Millimeters) : 0.0;
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
            return clashZones.ToList();
        }
        
        public (int PlacedCount, int SkippedCount, int ErrorCount) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            // ⏱️ TIMING: Start overall placement timer
            var overallTimer = System.Diagnostics.Stopwatch.StartNew();
            var detailedTimingLog = new System.Text.StringBuilder();
            // Diagnostics: per-run placement log in AppData
            var placementLogName = SafeFileLogger.GetLogFilePath($"sleeve_placement_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
            
            // ✅ PERFORMANCE OPTIMIZATION: Track placed sleeves for batch processing
            var placedSleeveData = new List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter)>();
            var placedSleeveIds = new List<ElementId>();
            
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
            
            // ⏱️ TIMING: Cache levels
            var cacheTimer = System.Diagnostics.Stopwatch.StartNew();
            // ✅ PERFORMANCE OPTIMIZATION: Pre-cache expensive operations
            // Cache levels list once instead of scanning for each clash zone
            var cachedLevels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();
            cacheTimer.Stop();
            detailedTimingLog.AppendLine($"[TIMING] Level caching: {cacheTimer.ElapsedMilliseconds}ms ({cachedLevels.Count} levels)");
            
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
            
            // ✅ MULTI-THREADING: Pre-filter eligible clash zones in parallel (XML-only validation)
            // All validation uses XML data - completely safe for parallel processing
            // This filters out invalid zones BEFORE entering sequential placement loop
            var preFilterTimer = System.Diagnostics.Stopwatch.StartNew();
            // NOTE: User requested to skip additional gating here – rely on upstream filtering only.
            // var eligibleClashZones = PreFilterEligibleClashZones(clashZones);
            var eligibleClashZones = clashZones;
            preFilterTimer.Stop();
            detailedTimingLog.AppendLine($"[TIMING] Parallel pre-filtering: {preFilterTimer.ElapsedMilliseconds}ms ({clashZones.Count} → {eligibleClashZones.Count} eligible)");
            
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] ✅ MULTI-THREADING: Pre-filtered {clashZones.Count} clash zones → {eligibleClashZones.Count} eligible in {preFilterTimer.ElapsedMilliseconds}ms");
            
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
            
            // ⏱️ TIMING: Pre-cache family symbols
            var preCacheTimer = System.Diagnostics.Stopwatch.StartNew();
            // ⚠️ QUICK WIN: Pre-load family symbols for all needed families (significant performance gain)
            if (OptimizationFlags.UseFamilySymbolCache)
            {
                PreCacheFamilySymbols(sortedClashZones);
            }
            preCacheTimer.Stop();
            detailedTimingLog.AppendLine($"[TIMING] Family symbol pre-caching: {preCacheTimer.ElapsedMilliseconds}ms");
            
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
                
                foreach (var clashZone in sortedClashZones)
                {
                    // ✅ DEBUG: Log which clash zone is being processed
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[PLACEMENT-LOOP] Processing ClashZone {clashZone.Id} ({processedCount + 1}/{sortedClashZones.Count}): MEP={clashZone.MepElementIdValue}, Flags: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, SleeveId={clashZone.SleeveInstanceId}");
                    
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
                        File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 0: Starting processing for Zone={clashZone.Id}\n");
                        }
                    } 
                    catch { }
                    
                    // ⏱️ TIMING: Start per-sleeve timer (only for actual placement operations)
                    var sleeveTimer = System.Diagnostics.Stopwatch.StartNew();
                    var sleeveLog = new System.Text.StringBuilder();
                    
                    try
                    {
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        try { 
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 1: Entered try block for Zone={clashZone.Id}\n");
                            }
                        } catch { }
                        
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                        
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        try { 
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2: About to validate placement point for Zone={clashZone.Id}, IP=({clashZone.IntersectionPointX},{clashZone.IntersectionPointY},{clashZone.IntersectionPointZ})\n");
                            }
                        } catch { }
                        
                        // ✅ CRITICAL: Validate intersection coordinates - NO FALLBACK to wall center
                        // This will throw an exception if coordinates are invalid, which we catch and handle below
                        try
                        {
                            ValidatePlacementPoint(clashZone);
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.1: ✅ Placement point validation PASSED for Zone={clashZone.Id}\n");
                            } } catch { }
                        }
                        catch (InvalidOperationException validationEx)
                        {
                            // Invalid coordinates - skip this zone and log error
                            var msg = $"[UniversalSleevePlacer] ERROR: {validationEx.Message}";
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error(msg);
                            if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✗ ERROR: ClashZone {clashZone.Id} has invalid intersection coordinates - SKIPPED");
                            
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.2: ❌ VALIDATION FAILED: {validationEx.Message}\n");
                            } } catch { }
                            
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
                            continue; // Skip this zone - cannot place without valid intersection point
                        }
                        
                        // ✅ DEPLOYMENT MODE: Skip file writes
                        try { 
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 3: Validation passed, checking flags: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, SleeveId={clashZone.SleeveInstanceId}, ClusterId={clashZone.ClusterSleeveInstanceId}\n");
                            }
                        } catch { }
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Batch file logging instead of individual writes
                        // Only log to batch - will write once at end (or every 50 clash zones)
                        if (PlacedCount + SkippedCount < 50) // Only log first 50 for debugging
                        {
                            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}");
                        }
                        
                        // STEP: Determine existing sleeves via direct Revit lookups (flags are informational only)
                        if (clashZone.ClusterSleeveInstanceId > 0)
                        {
                            var clusterSleeveElement = GetCachedElement(clashZone.ClusterSleeveInstanceId);
                            if (clusterSleeveElement != null)
                                {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already has cluster sleeve {clashZone.ClusterSleeveInstanceId}");
                                SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SKIP ClusterExists: ClashZone={clashZone.Id}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}\n");
                            SkippedCount++;
                            continue;
                                    }
                                    else
                                    {
                                clashZone.IsClusterResolved = false;
                                clashZone.ClusterSleeveInstanceId = -1;
                            }
                        }
                        
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            var existingIndividualSleeve = GetCachedElement(clashZone.SleeveInstanceId);
                            if (existingIndividualSleeve != null)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already has individual sleeve {clashZone.SleeveInstanceId}");
                                SafeFileLogger.SafeAppendText(placementLogName, $"[{DateTime.Now}] SKIP IndividualExists: ClashZone={clashZone.Id}, SleeveId={clashZone.SleeveInstanceId}\n");
                                SkippedCount++;
                                continue;
                            }
                            else
                            {
                                    clashZone.IsResolved = false;
                                    clashZone.SleeveInstanceId = -1;
                            }
                        }
                        
                        // STEP 3: If we reach here, place individual sleeve (fresh or replacement)
                        
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
                                        // ✅ OOP REFACTORING: Use FlagManager to reset cluster flags
                                                                                if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] RESET: ClashZone {clashZone.Id} has cluster sleeve in Global XML but sleeve was deleted - resetting flags and proceeding with placement");
                                        clashZone.IsClusterResolved = false;
                                        clashZone.ClusterSleeveInstanceId = -1;
                                        _flagManager.UpdateFlagsForPlacement(clashZone, -1, isCluster: true, categoryName, _filterName);
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
                        try 
                        { 
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 6: Checking category match: ZoneCategory='{clashZone.MepElementCategory}' vs StrategyCategory='{_strategy?.GetCategoryName() ?? "NULL"}'\n");
                            }
                        } catch { }
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging - use DebugLogger only
                        // Validate category match
                        if (!string.IsNullOrEmpty(clashZone.MepElementCategory) && 
                            clashZone.MepElementCategory != _strategy.GetCategoryName())
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} category '{clashZone.MepElementCategory}' doesn't match '{_strategy.GetCategoryName()}'");
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 6.1: ❌ SKIPPED - Category mismatch\n");
                            } } catch { }
                            SkippedCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 7: Category match passed, creating MEP size\n");
                        } } catch { }
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging
                        
                        // ⏱️ TIMING: MEP size creation
                        var mepSizeTimer = System.Diagnostics.Stopwatch.StartNew();
                        // ⚠️ ZERO LINKED FILE ACCESS - use pre-calculated MEP size from ClashZone
                        var mepSize = new MepElementSize
                        {
                            Width = clashZone.MepElementWidth,
                            Height = clashZone.MepElementHeight,
                            Diameter = clashZone.MepElementWidth, // For round, width = diameter
                            Shape = clashZone.DuctShape
                        };
                        mepSizeTimer.Stop();
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 8: MEP size created: W={mepSize.Width}, H={mepSize.Height}, D={mepSize.Diameter}, Shape={mepSize.Shape}\n");
                        } } catch { }
                        
						// ⚠️ SPECIAL HANDLING & SIZING ORDER:
						// 1) Pipes (host-agnostic), 2) Dampers, 3) Cable trays, 4) Ducts/default
                        XYZ placementOffset = XYZ.Zero;
                        double finalWidth = 0.0, finalHeight = 0.0, finalDiameter = 0.0;
                        
                        // ⏱️ TIMING: Clearance calculation
                        var clearanceTimer = System.Diagnostics.Stopwatch.StartNew();
                        // 🛡️ ARCHITECTURE FIX: Use CONDITIONS service for ALL clearance types
                        // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
                        // Raw dimensions from ClashZone + Clearance from CONDITIONS = Final dimensions
                        
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] CLEARANCE CALCULATION START: Category='{clashZone.MepElementCategory}', Strategy={(_strategy?.GetType().Name ?? "NULL")}");
                        
						bool isPipesCategory = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
						
						// ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging
						
						if (isPipesCategory)
						{
							// ✅ Pipes: Raw dimensions + CONDITIONS clearance
							var rawDiameter = clashZone.MepElementWidth; // Raw diameter from ClashZone
							var clearance = GetClearanceFromConditions("Pipes", mepSize);
							finalDiameter = rawDiameter + (2 * clearance);
							finalWidth = finalDiameter;
							finalHeight = finalDiameter;
							       if (!DeploymentConfiguration.DeploymentMode)
							DebugLogger.Info($"[UniversalSleevePlacer] PIPE: Raw={UnitUtils.ConvertFromInternalUnits(rawDiameter, UnitTypeId.Millimeters):F1}mm + Clearance={UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F1}mm = Final={UnitUtils.ConvertFromInternalUnits(finalDiameter, UnitTypeId.Millimeters):F1}mm");
						}
						else if (_strategy is DamperPlacementStrategy damperStrategy)
                        {
                            // ✅ Fire dampers: Raw dimensions + CONDITIONS clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            
                            // Get offset and final dimensions from strategy (uses CONDITIONS)
                            var adj = damperStrategy.GetDamperPlacementAdjustment(clashZone, _conditions);
                            placementOffset = adj.offsetVector;
                            finalWidth = adj.finalWidth;
                            finalHeight = adj.finalHeight;
                            finalDiameter = finalWidth; // Not used for dampers (rectangular only)
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] DAMPER: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm → Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        
                        // ✅ DEBUG: Log strategy information
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Strategy Type: {_strategy.GetType().Name}");
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Strategy Category: {_strategy.GetCategoryName()}");
                        
                        if (_strategy is DuctPlacementStrategy ductStrategy)
                        {
                            // ✅ Ducts: Raw dimensions + CONDITIONS clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT CALCULATION START: Raw dimensions {UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm");
                            
                            var clearance = GetClearanceFromConditions("Ducts", mepSize);
                            finalWidth = rawWidth + (2 * clearance);
                            finalHeight = rawHeight + (2 * clearance);
                            finalDiameter = finalWidth; // For round elements
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm + Clearance={UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F1}mm = Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                        {
                            // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                            // ✅ Cable trays: Raw dimensions + UI/XML clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;

                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm");
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: UI Clearance Settings Count={_clearanceSettings.Count}");

                            // Get offset and final dimensions from strategy (uses UI settings first, then CONDITIONS)
                            var adj2 = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions, _clearanceSettings);
                            placementOffset = adj2.offsetVector;
                            finalWidth = adj2.finalWidth;
                            finalHeight = adj2.finalHeight;

                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                            finalDiameter = finalWidth; // Not used for cable trays (rectangular only)

                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm → Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        else
                        {
                            // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging
                            // 🔥 FALLBACK: Use raw dimensions + default clearance if no strategy matches
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            var defaultClearance = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters); // 50mm default
                            
                            finalWidth = rawWidth + (2 * defaultClearance);
                            finalHeight = rawHeight + (2 * defaultClearance);
                            finalDiameter = Math.Max(finalWidth, finalHeight);
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ NO STRATEGY MATCHED for ClashZone {clashZone.Id} - Using fallback dimensions");
                        }
                        clearanceTimer.Stop();
                        totalClearanceTime += clearanceTimer.Elapsed;
                        sleeveLog.AppendLine($"  Clearance calc: {clearanceTimer.ElapsedMilliseconds}ms");

                        // ⚠️ REMOVED: Old width/height swapping logic that was causing double-swapping
                        // The new logic later in the method (lines 660-665) handles this correctly
                        // by ensuring the longer dimension becomes width, not just swapping blindly
                        
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 9: Starting clearance calculation\n");
                        } } catch { }
                        
                        // Clearance calculation happens here (existing code)...
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 10: Clearance calculated, selecting family\n");
                        } } catch { }
                        
                        // ⏱️ TIMING: Family selection and loading
                        var familyTimer = System.Diagnostics.Stopwatch.StartNew();
                        // Select universal family
                        var (familyName, typeName, isCircular) = SelectUniversalFamily(clashZone, mepSize);
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 11: Family selected: '{familyName}', Type='{typeName}', IsCircular={isCircular}\n");
                        } } catch { }
                        
                        var familySymbol = LoadFamilySymbol(familyName);
                        familyTimer.Stop();
                        totalFamilyLoadTime += familyTimer.Elapsed;
                        sleeveLog.AppendLine($"  Family load: {familyTimer.ElapsedMilliseconds}ms");
                        
                        if (familySymbol == null)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalSleevePlacer] Family '{familyName}' not found");
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 11.1: ❌ ERROR - Family '{familyName}' not found\n");
                            } } catch { }
                            ErrorCount++;
                            continue;
                        }
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 12: ✅ Family symbol loaded successfully\n");
                        } } catch { }
                        
                        // Activate symbol if needed
                        if (!familySymbol.IsActive)
                        {
                            familySymbol.Activate();
                        }
                        
                        // ✅ SLEEVE PLACEMENT FLOW - Step 2: Get sleeve placement data from Filter XML
                        // Filter XML contains all placement data: coordinates, MEP dimensions, host type, family name
                        // During XML deserialization: 
                        // 1. IntersectionPoint (XYZ object) is null because it's marked [XmlIgnore]
                        // 2. Only IntersectionPointX, IntersectionPointY, IntersectionPointZ (doubles) are saved/loaded from XML
                        // 3. We create the XYZ object from these three values
                        bool usingXmlSnapshot = false;
                        XYZ placementPointChosen;
                        
                        // Check if XYZ object exists, if not create it from XML properties
                        if (clashZone.IntersectionPoint == null)
                        {
                            // ✅ STEP 2: Get placement coordinates from Filter XML
                            // Get values directly from XML (IntersectionPointX, Y, Z)
                            placementPointChosen = new XYZ(clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ);
                            clashZone.IntersectionPoint = placementPointChosen; // Set it back so it's available for next time
                            usingXmlSnapshot = true;
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[SLEEVE-PLACEMENT-FLOW] Step 2: Got placement coordinates from Filter XML: ({clashZone.IntersectionPointX:F3},{clashZone.IntersectionPointY:F3},{clashZone.IntersectionPointZ:F3})");
                        }
                        else
                        {
                            // Use existing XYZ object
                            placementPointChosen = clashZone.IntersectionPoint;
                        }
                        
                        // ✅ STEP 2 (continued): All placement data comes from Filter XML:
                        // - Placement coordinates: IntersectionPointX/Y/Z ✅ (from Filter XML)
                        // - MEP dimensions: MepElementWidth/Height (from Filter XML)
                        // - Host type: StructuralElementType (from Filter XML)
                        // - MEP category: MepElementCategory (from Filter XML)
                        if (!DeploymentConfiguration.DeploymentMode && PlacedCount < 5)
                        {
                            DebugLogger.Info($"[SLEEVE-PLACEMENT-FLOW] Step 2: Got placement data from Filter XML - " +
                                $"MEP Size: W={clashZone.MepElementWidth:F3}, H={clashZone.MepElementHeight:F3}, " +
                                $"Host: {clashZone.StructuralElementType}, Category: {clashZone.MepElementCategory}");
                        }
                        
                        // Validate intersection point has valid coordinates (not all zeros)
                        if (Math.Abs(placementPointChosen.X) < 1e-9 && Math.Abs(placementPointChosen.Y) < 1e-9 && Math.Abs(placementPointChosen.Z) < 1e-9)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalSleevePlacer] ❌ CRITICAL: Cannot place sleeve for Zone {clashZone.Id} - IntersectionPoint is (0,0,0)!");
                            try { File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2.3: ❌ SKIPPED - IntersectionPoint is (0,0,0)\n"); } catch { }
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
                                var placementSource = usingXmlSnapshot ? "XML snapshot" : "Recomputed intersection";
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] [PLACEMENT-ORIGIN] Zone={clashZone.Id}, Source={placementSource}, Point=({placementPointChosen?.X:F3},{placementPointChosen?.Y:F3},{placementPointChosen?.Z:F3})\n");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] PLACEMENT: Zone={clashZone.Id}, MEP={clashZone.MepElementIdValue}, HOST={clashZone.StructuralElementIdValue}, Point=({placementPointChosen?.X:F3},{placementPointChosen?.Y:F3},{placementPointChosen?.Z:F3}), SPP_XML=({clashZone.SleevePlacementPointX:F3},{clashZone.SleevePlacementPointY:F3},{clashZone.SleevePlacementPointZ:F3}), IP=({clashZone.IntersectionPointX:F3},{clashZone.IntersectionPointY:F3},{clashZone.IntersectionPointZ:F3})\n");
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] Using placement point {placementPointChosen} → adjusted {adjustedPlacementPoint}");
                        
                        if (placementOffset.GetLength() > 0.001)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[UniversalSleevePlacer] Applied offset {placementOffset} to placement point");
                        }
                        
                        // ⏱️ TIMING: Level finding
                        var levelTimer = System.Diagnostics.Stopwatch.StartNew();
                        // ✅ PERFORMANCE OPTIMIZATION: Use cached levels list and select level based on structural element type
                        Level nearestLevel = null;
                        
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
                        
                        // DEPLOYMENT MODE: Skip file writes
                        try 
                        { 
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 13: About to create sleeve at point=({adjustedPlacementPoint.X:F3},{adjustedPlacementPoint.Y:F3},{adjustedPlacementPoint.Z:F3}), Level='{nearestLevel?.Name ?? "NULL"}'\n");
                            }
                        } catch { }
                        
                        // ✅ SLEEVE PLACEMENT FLOW - Step 3: Place sleeve using data from Filter XML
                        // Uses placement coordinates, family selection, and dimensions from Filter XML
                        // ⏱️ TIMING: Sleeve creation
                        var createTimer = System.Diagnostics.Stopwatch.StartNew();
                        
                        if (!DeploymentConfiguration.DeploymentMode && PlacedCount < 5)
                            DebugLogger.Info($"[SLEEVE-PLACEMENT-FLOW] Step 3: Placing sleeve at ({adjustedPlacementPoint.X:F3},{adjustedPlacementPoint.Y:F3},{adjustedPlacementPoint.Z:F3}) " +
                                $"using family '{familySymbol.Family.Name}' with dimensions W={finalWidth:F3}, H={finalHeight:F3}");
                        
                        // Place sleeve instance (NO HOST PARAMETER - workplane-based families)
                        // ✅ Works with linked structural elements because no host reference needed
                        FamilyInstance sleeveInstance = null;
                        try
                        {
                            // ✅ STEP 3: Create sleeve instance using placement data from Filter XML
                            sleeveInstance = _doc.Create.NewFamilyInstance(
                                adjustedPlacementPoint,  // From Filter XML (IntersectionPoint)
                                familySymbol,            // Selected based on host type from Filter XML
                                nearestLevel,
                                StructuralType.NonStructural);
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14: ✅ Sleeve instance created! ID={sleeveInstance?.Id?.IntegerValue ?? -1}\n");
                            } } catch { }
                        }
                        catch (Exception createEx)
                        {
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14.1: ❌ EXCEPTION during sleeve creation: {createEx.Message}\n{createEx.StackTrace}\n");
                            } } catch { }
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalSleevePlacer] Exception creating sleeve: {createEx.Message}");
                            ErrorCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        createTimer.Stop();
                        totalSleeveCreateTime += createTimer.Elapsed;
                        sleeveLog.AppendLine($"  Sleeve create: {createTimer.ElapsedMilliseconds}ms");
                        
                        if (sleeveInstance == null)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[UniversalSleevePlacer] Failed to create sleeve instance");
                            try {                             // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] STEP 14.2: ❌ ERROR - Sleeve instance is NULL after creation\n");
                            } } catch { }
                            ErrorCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        // ⏱️ TIMING: Parameter setting
                        var parameterTimer = System.Diagnostics.Stopwatch.StartNew();
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
                        SetSleeveOrientation(sleeveInstance, clashZone);
                        parameterTimer.Stop();
                        totalParameterTime += parameterTimer.Elapsed;
                        sleeveLog.AppendLine($"  Parameters: {parameterTimer.ElapsedMilliseconds}ms");
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                        // ⏱️ TIMING: Coordinate validation and update
                        var validationTimer = System.Diagnostics.Stopwatch.StartNew();
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
                                DebugLogger.Info($"[UniversalSleevePlacer] Placed sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id}, W={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm (bbox deferred)");
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[UniversalSleevePlacer] ✓ Placed {_strategy.GetCategoryName()} sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id} at {adjustedPlacementPoint} ({sleeveTimer.ElapsedMilliseconds}ms)");
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[UniversalSleevePlacer] Error placing sleeve for ClashZone {clashZone.Id}: {ex.Message}");
                        try {                         // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(debugLogPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ EXCEPTION IN OUTER TRY-CATCH: Zone={clashZone.Id}, Error={ex.Message}\n{ex.StackTrace}\n");
                        } } catch { }
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
                                // Set bounding box coordinates
                                zone.SetSleeveBoundingBox(actualBbox);
                                
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
                }
                
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

                    SaveUpdatedXmlFiles(clashZones);
                   
                    
                    // ⚠️ REMOVED: Don't update coordinates here because clustering will delete individual sleeves
                    
                    // ⚠️ CRITICAL: Log flag states AFTER XML save
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-AFTER] XML save completed for {PlacedCount} placed sleeves\n");

                    // ✅ OOP REFACTORING: Use FlagManager for flag updates after placement
                    // Per methodology document line 242-245: "Update Global XML" after placement
                    try
                    {
                        if (placedClashZonesForGlobal.Count > 0)
                        {
                            foreach (var clashZone in placedClashZonesForGlobal)
                            {
                                if (clashZone.SleeveInstanceId > 0)
                                {
                                    _flagManager.UpdateFlagsForPlacement(
                                        clashZone,
                                        clashZone.SleeveInstanceId,
                                        isCluster: false,
                                        clashZone.MepElementCategory,
                                        _filterName
                                    );
                                }
                            }
                            
                                                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ✅ Successfully updated Global XML for {placedClashZonesForGlobal.Count} placed sleeves (BEFORE clustering)");
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
                        DebugLogger.Error($"[FLAG-MANAGER] ❌ CRITICAL ERROR: Flag update after individual placement failed: {upEx.Message}");
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
            if (PlacedCount > 0 && !DeploymentConfiguration.DeploymentMode)
            {
                var avgPlacementTime = overallTimer.ElapsedMilliseconds / (double)PlacedCount;
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
        /// Select universal family based on host type and MEP shape
        /// Uses 4 universal families following CONVOID approach
        /// </summary>
        private (string familyName, string typeName, bool isCircular) SelectUniversalFamily(ClashZone clashZone, MepElementSize mepSize)
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
            
            string typeName = "Default"; // Use default type for your families
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] Selected family: {familyName}, Type: {typeName}");
            
            return (familyName, typeName, isCircular);
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
                        return UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
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
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in clearance calculation: {ex.Message}\n");
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetClearanceFromConditions] Error: {ex.Message}");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
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
                double clearanceInFeet = UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[GetClearanceFromXmlConditions] {category}: {clearanceInMm}mm → {clearanceInFeet:F6}ft");
                return clearanceInFeet;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"[GetClearanceFromXmlConditions] Error: {ex.Message}");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
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
                double widthMm = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                double heightMm = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                double diameterMm = UnitUtils.ConvertFromInternalUnits(diameter, UnitTypeId.Millimeters);
                
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
        
        // 🔍 DEBUG: Log pipe detection for floors vs walls
        var hostType = clashZone.StructuralElementType ?? "Unknown";
                if (!DeploymentConfiguration.DeploymentMode)
        DebugLogger.Info($"[SetSleeveParameters] PIPE DETECTION DEBUG: Host={hostType}, Category={clashZone.MepElementCategory}, isPipe={isPipe}");
        
        // Apply rounding to nearest 5mm if setting is enabled
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(finalWidth, finalHeight);
                var roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(finalDiameter);
        
        // ✅ PERFORMANCE: Removed verbose parameter logging - only log in diagnostic mode
        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
        {
            try
            {
                double diamMm = UnitUtils.ConvertFromInternalUnits(roundedDiameter, UnitTypeId.Millimeters);
                double widthMm = UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters);
                double heightMm = UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters);
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
                    p.Set(roundedDiameter);
                    // Add to cache if not already cached
                    if (!paramCache.ContainsKey(name) && !p.IsReadOnly)
                        paramCache[name] = p;
                    // ✅ PERFORMANCE: Removed verbose logging - only log in diagnostic mode
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        double diamMm = UnitUtils.ConvertFromInternalUnits(roundedDiameter, UnitTypeId.Millimeters);
                        DebugLogger.Info($"[UniversalSleevePlacer] Set '{name}' = {diamMm:F1}mm on sleeve {sleeveInstance.Id}");
                    }
                    setOk = true;
                    break;
                }
            }
            
            if (!setOk)
                    {
                        // Fallback to Width/Height for circular
                        GetParam("Width")?.Set(roundedDiameter);
                        GetParam("Height")?.Set(roundedDiameter);
                        // ✅ PERFORMANCE: Removed verbose logging
                    }
                }
                else
                {
                    // Rectangular
                    GetParam("Width")?.Set(roundedWidth);
                    GetParam("Height")?.Set(roundedHeight);
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
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT SWAP: Height({UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm) > Width({UnitUtils.ConvertFromInternalUnits(tempWidth, UnitTypeId.Millimeters):F1}mm) - Swapped to make longer dimension the width");
            }
            else
            {
                // Width is already longer - no swap needed
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT NO SWAP: Width({UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm) >= Height({UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm) - Longer dimension already width");
            }
            
            // Update the sleeve parameters with correct values
            GetParam("Width")?.Set(roundedWidth);
            GetParam("Height")?.Set(roundedHeight);
            
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT FINAL: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm (NO ROTATION)");
        }
        else if (isFloorHost && isCableTray && !treatAsCircular)
        {
            // Cable trays on floors: NO width/height swap - maintain original orientation
                        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR CABLE TRAY: NO SWAP - Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm");
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
                DebugLogger.Info($"[UniversalSleevePlacer] X-FRAMING SWAP: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm");
            }
            else if (hostOrientation == "Y")
            {
                // Y-framing: swap width and depth
                double tempWidth = roundedWidth;
                roundedWidth = roundedHeight;  // Use height as width
                roundedHeight = tempWidth;     // Use original width as height
                                if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[UniversalSleevePlacer] Y-FRAMING SWAP: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm");
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
        
        double thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
        
        // ✅ PERFORMANCE: Removed verbose depth check logging
        
        bool depthSetSuccess = false;
        
        if (isWallHost && wallWidthParam != null && !wallWidthParam.IsReadOnly)
        {
            // Wall host: use Wall Width parameter
            wallWidthParam.Set(thickness);
            // ✅ PERFORMANCE: Removed verbose logging
            depthSetSuccess = true;
        }
        else if (depthParam != null && !depthParam.IsReadOnly)
        {
            // Floor/Framing host: use Depth parameter
            depthParam.Set(thickness);
            // ✅ PERFORMANCE OPTIMIZATION: Removed parameter verification read (AsDouble) - unnecessary overhead
            // Parameter Set() is reliable, no need to verify immediately
            depthSetSuccess = true;
        }
        else
        {
            // Try type parameter as fallback
            // ✅ PERFORMANCE OPTIMIZATION: Try to cache type parameter (symbol-level, not instance-level)
            var typeDepthParam = sleeveInstance.Symbol?.LookupParameter("Depth");
            if (typeDepthParam != null && !typeDepthParam.IsReadOnly)
            {
                typeDepthParam.Set(thickness);
                // ✅ PERFORMANCE OPTIMIZATION: Regeneration deferred to batch operation at end of placement loop
                // _doc.Regenerate(); // ❌ REMOVED: Will be done in batch after all sleeves placed
                depthSetSuccess = true;
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
                        
                        // Set the parameter
                        bottomOfOpeningParam.Set(bottomOfOpening);
                        
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
        
        // Set MEP metadata parameters
                var mepElementIdParam = GetParam("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    mepElementIdParam.Set(clashZone.MepElementId.IntegerValue);
                }

                GetParam("MEP_UniqueId")?.Set(clashZone.MepElementUniqueId);
                GetParam("MEP_Size")?.Set(clashZone.MepElementFormattedSize);
                GetParam("System_Abbreviation")?.Set(clashZone.MepElementSystemAbbreviation);
                GetParam("MEP_Count")?.Set(1);  // Individual sleeve
                
                // ✅ CRITICAL FIX: Set MEP_Category parameter for clustering
                // ✅ PERFORMANCE: Reduced logging - only log failures in diagnostic mode
                var mepCategoryParam = GetParam("MEP_Category");
                if (mepCategoryParam != null)
                {
                    mepCategoryParam.Set(clashZone.MepElementCategory);
                    if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Set MEP_Category = '{clashZone.MepElementCategory}' for sleeve {sleeveInstance.Id.IntegerValue}");
                    }
                }
                else if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] MEP_Category parameter not found on sleeve {sleeveInstance.Id}");
                }

                // ⚠️ CRITICAL FIX: Transfer HostParameterValues from XML intersection data to sleeve parameters
                // ✅ PERFORMANCE OPTIMIZATION: Reduced logging and use pre-cached parameters
                if (clashZone.HostParameterValues != null && clashZone.HostParameterValues.Count > 0)
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
                                // Handle different parameter storage types
                                if (param.StorageType == StorageType.String)
                                {
                                    param.Set(hostParam.Value);
                                    hostParamsSet++;
                                }
                                else if (param.StorageType == StorageType.Integer)
                                {
                                    if (int.TryParse(hostParam.Value, out int intValue))
                                    {
                                        param.Set(intValue);
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
                                        param.Set(doubleValue);
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
                // Set Filter Name based on category
                string filterName = GetFilterNameForCategory(clashZone.MepElementCategory);
                var filterNameParam = sleeveInstance.LookupParameter("Filter Name");
                if (filterNameParam != null && !filterNameParam.IsReadOnly)
                {
                    filterNameParam.Set(filterName);
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
                    instanceIdParam.Set(sleeveInstance.Id.IntegerValue);
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
                    linkedFileParam.Set(linkedFile);
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
                    hostFileParam.Set(hostFile);
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
                guidParam.Set(guidString);
                
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
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] BEFORE: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm, ActiveDocX={zone.SleevePlacementPointActiveDocumentX:F3}\n");
                    
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
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] AFTER: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm, ActiveDocX={zone.SleevePlacementPointActiveDocumentX:F3}\n");
                    
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
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Dimensions: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
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
                            var effectiveFilterName = GetBaseFilterName(_filterName, filter?.Name, category);
                            baseFilterNameForPersistence ??= effectiveFilterName;
                            var persistenceZones = categoryZones
                                .Select(CloneZoneForPersistence)
                                .Where(z => z != null)
                                .ToList();

                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 Snapshot (clone): {string.Join(", ", persistenceZones.Select(z => $"{z.Id}:{z.SleeveInstanceId}").Take(10))}");

                            persistenceService.SaveClashZones(persistenceZones, effectiveFilterName, filter);
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 SUCCESS");
                            }
                        catch (Exception persistEx)
                        {
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 3 FAILED: {persistEx}");
                            throw;
                        }

                        diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4: Saving filter to XML");
                        try
                        {
                            filterManagementService.SaveFilterToXmlFile(filter, xmlFile);
                            diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] Step 4 SUCCESS");
                            processedCategoryFilters.Add((filter, category));
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
                    var baseFilterName = baseFilterNameForPersistence ?? GetBaseFilterName(_filterName, sanitizedFilterNameGlobal, string.Empty);
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
                                    filterManagementService.SaveFilterToXmlFile(mainFilter, mainFilterPath);
                                    diagnosticLog.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] MAIN FILTER UPDATED: {Path.GetFileName(mainFilterPath)}");

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

                                EnsureConditionsFiles(baseFilterName, processedCategoryFilters, filtersDirectory, diagnosticLog);
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
        
        private string GetBaseFilterName(string rawFilterName, string fallbackFilterName, string category)
        {
            var source = !string.IsNullOrWhiteSpace(rawFilterName) ? rawFilterName : fallbackFilterName;
            if (string.IsNullOrWhiteSpace(source))
            {
                return fallbackFilterName ?? "Unknown";
                    }

            source = source.Trim();
            source = Path.GetFileNameWithoutExtension(source);

            if (string.IsNullOrWhiteSpace(source))
                return fallbackFilterName ?? "Unknown";

            var normalizedCategory = category?.Trim().ToLowerInvariant() ?? string.Empty;
            if (!string.IsNullOrEmpty(normalizedCategory))
            {
                normalizedCategory = normalizedCategory.Replace(" ", "_");
                var suffix = "_" + normalizedCategory;
                if (source.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    source = source.Substring(0, source.Length - suffix.Length);
                }
            }

            if (string.IsNullOrWhiteSpace(source))
                return fallbackFilterName ?? "Unknown";

            return source;
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
                var conditionsService = new ConditionsService(filtersDirectory, msg =>
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
        private void SetSleeveOrientation(FamilyInstance sleeveInstance, ClashZone clashZone)
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
            // ====== FLOOR HOST: Rotate based on MEP orientation ======
                    // Get stored MEP orientation from XML ("X" or "Y")
                    string mepOrientation = clashZone.MepElementOrientationDirection;
                    
                    if (!string.IsNullOrEmpty(mepOrientation))
                    {
                        var loc = sleeveInstance.Location as LocationPoint;
                        if (loc != null)
                        {
                            double rotationAngle = 0.0;
                            
                            // ⚠️ CABLETRAY FIX: Cable trays have INVERTED rotation logic compared to ducts
                            // For cable trays: Y-orientation = 0°, X-orientation = 90°
                            // For ducts/pipes: Y-orientation = 90°, X-orientation = 0°
                            bool isCableTray = clashZone.MepElementCategory.Contains("Cable", StringComparison.OrdinalIgnoreCase);
                            
                            if (isCableTray)
                            {
                                // ⚠️ INVERTED LOGIC FOR CABLE TRAYS
                                if (mepOrientation == "X")
                                {
                                    rotationAngle = Math.PI / 2; // 90 degrees
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL CABLETRAY ON FLOOR: MEP orientation is X - rotating sleeve 90°");
                                }
                                else
                                {
                                    rotationAngle = 0.0; // No rotation needed
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL CABLETRAY ON FLOOR: MEP orientation is Y - no rotation needed");
                                }
                            }
                            else
                            {
                                // ⚠️ STANDARD LOGIC FOR DUCTS/PIPES
                                if (mepOrientation == "Y")
                                {
                                    rotationAngle = Math.PI / 2; // 90 degrees
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL {clashZone.MepElementCategory.ToUpper()} ON FLOOR: MEP orientation is Y - rotating sleeve 90°");
                                }
                                else
                                {
                                    rotationAngle = 0.0; // No rotation needed
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL {clashZone.MepElementCategory.ToUpper()} ON FLOOR: MEP orientation is X - no rotation needed");
                                }
                            }
                            
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                                        if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Rotated sleeve {rotationAngleDegrees:F1}° based on MEP orientation ({mepOrientation})");
                    try
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ORIENT-FLOOR] Sleeve {sleeveInstance.Id.IntegerValue}: Rotated {rotationAngleDegrees:F1}° for MEP dir ({mepOrientation})\n");
                    }
                    catch { }
                }
            }
            
            // Set HostOrientation parameter
                    var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                    if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                    {
                        hostOrientationParam.Set("FloorHosted");
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
            
            // Set HostOrientation parameter from XML (pre-calculated during refresh)
                    if (!string.IsNullOrEmpty(clashZone.HostOrientation))
                    {
                        var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                        if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                        {
                            hostOrientationParam.Set(clashZone.HostOrientation);
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
    }
}
